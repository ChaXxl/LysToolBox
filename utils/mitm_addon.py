import json
import re
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional, List, Dict, Any, Union

import httpx
import openpyxl
import pandas as pd
import shortuuid
from loguru import logger
from lxml import etree
from mitmproxy import http
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from PySide6.QtCore import QThread, Signal

from utils.medicineID import MEDICINE_ID


class Addon(QThread):
    """
    MITM代理插件类，用于拦截和处理各电商平台的药品数据
    继承自QThread以便与Qt界面交互
    """

    add_text = Signal(str)  # 用于向UI发送文本信息的信号

    def __init__(self):
        """初始化Addon类"""
        super().__init__()

        # HTTP客户端，用于发送请求
        self.h = httpx.Client()

        # 搜索关键词
        self.keyword = None

        # Excel相关
        self.sheet = None
        self.workBook = None
        self.filename: Optional[Path] = None

        # 创建线程池，用于并行处理数据保存等耗时操作
        self.thread = ThreadPoolExecutor(max_workers=5)

        # 药品品牌名和药品名称，从关键词中解析
        self.brand_name = []
        self.medicine_name = ""

    def createExcel(self, filename: Path) -> None:
        """
        创建或加载Excel文件

        Args:
            filename: Excel文件路径
        """
        self.filename = filename

        # 判断文件是否存在, 不存在则新建
        if not filename.exists():
            self.workBook = openpyxl.Workbook()  # 创建一个工作簿对象
        else:
            self.workBook = openpyxl.load_workbook(
                filename, keep_vba=True
            )  # 打开Excel表格并保留宏

        self.sheet = self.workBook.active  # 选取第一个sheet

    def save_to_excel(self, datas: List[List[Any]], tag: Optional[str] = None) -> None:
        """
        保存数据到Excel文件

        Args:
            datas: 要保存的数据列表
            tag: 标签-指明哪个平台
        """
        # 数据为空则直接返回
        if not datas:
            return

        df = None

        # 判断文件是否存在, 不存在则新建
        if not self.filename.exists():
            workBook = openpyxl.Workbook()  # 创建一个工作簿对象
        else:
            workBook = openpyxl.load_workbook(
                self.filename, keep_vba=True
            )  # 打开Excel表格

            # 读取现有数据用于去重
            df = pd.read_excel(self.filename)

        sheet = workBook.active  # 选取第一个sheet

        max_row = sheet.max_row
        saved_count = 0

        # 表头
        headers = [
            "uuid",
            "药店名称",
            "店铺主页",
            "资质名称",
            "营业执照图片",
            "药品名",
            "药品ID",
            "药品图片",
            "原价",
            "挂网价格",
            "平台",
            "排查日期",
        ]

        # 如果是第一次保存数据, 就添加表头
        if max_row == 1:
            sheet.append(headers)

        save_flag = True

        for data in datas:
            # 重复数据不保存 - 根据药店名称、店铺主页、药品名、挂网价格、平台判断
            if df is not None:
                temp_df = df[
                    (df["药店名称"] == data[1])
                    & (df["店铺主页"] == data[2])
                    & (df["药品名"] == data[5])
                    & (df["挂网价格"] == float(data[8]))
                    & (df["平台"] == data[9])
                ]

                if not temp_df.empty:
                    continue

            # 生成一个短UUID
            short_uuid = shortuuid.uuid()
            data[0] = short_uuid

            sheet.append(data)
            saved_count += 1
            save_flag = False

        if save_flag:
            msg = f"{self.filename.stem} {tag}-没有数据需要保存"
            logger.info(msg)
            return

        # 保存工作簿
        workBook.save(self.filename)

        msg = f"\n\n{self.filename.stem} {tag}-保存了{saved_count}条, 数据总条数: {sheet.max_row - 1}\n\n"
        self.add_text.emit(msg)

    def check_brand_product_name(self, name: str) -> bool:
        """
        检查产品名称是否符合搜索条件

        Args:
            name: 产品名称

        Returns:
            bool: 是否符合搜索条件
        """
        # 关键词为空则返回False
        if not self.keyword:
            return False

        # 解析关键词，分割品牌名和药品名
        keywords = self.keyword.split(" ")
        self.brand_name, self.medicine_name = keywords[:-1], keywords[-1]

        # 取药品名的前3个字符进行匹配，如果药品名长度不足3则用全部
        medicine_name_temp = (
            self.medicine_name[:3]
            if len(self.medicine_name) > 2
            else self.medicine_name
        )

        # 药品名必须在产品名中
        if medicine_name_temp not in name:
            return False

        # 特殊处理"一口"品牌
        if "一口" in self.brand_name:
            return "一口" in name
        else:
            # 至少要包含其中一个品牌名
            for brand in self.brand_name:
                if brand in name:
                    return True

        return False

    def parsejd2HTML(self, html_str: str) -> None:
        """
        解析京东后30条商品数据

        Args:
            html_str: HTML字符串
        """
        # 将字符串转换为列表以便按行处理
        html_lines = html_str.split("\n")

        # 删除不需要的内容
        start_idx = -1
        end_idx = len(html_lines)

        for i, line in enumerate(html_lines):
            if "li data-sku=" in line:
                start_idx = i
                break

        for i, line in enumerate(html_lines[start_idx:], start=start_idx):
            if "<script>" in line:
                end_idx = i
                break

        if start_idx == -1:
            return

        # 提取有效HTML内容
        html_content = "\n".join(html_lines[start_idx:end_idx])

        # 解析HTML
        try:
            html = etree.HTML(html_content)
            li_elements = html.xpath("//li")
        except Exception as e:
            logger.error(f"解析京东HTML失败: {e}")
            return

        datas = []
        # 解析商品信息
        for li in li_elements:
            try:
                # 商品名称
                productName = li.xpath("string(./div/div[3]//em)").split("\n")[0]

                # 商品价格
                price = li.xpath(".//div/div[2]/strong/i/text()")[0]

                # 商品图片
                productImg = "https:" + li.xpath("./div[1]//img/@data-lazy-img")[0]

                # 店铺名称
                storeName = li.xpath('.//div[@class="p-shop"]/span/a/@title')[0]

                # 店铺链接
                storeUrl = "https:" + li.xpath(".//div/div[5]/span/a/@href")[0]

                # 转为字符串类型
                productName = str(productName)
                price = str(price)
                productImg = str(productImg)
                storeName = str(storeName)
                storeUrl = str(storeUrl)

                # 跳过乐药师大药房旗舰店的商品
                if storeName == "乐药师大药房旗舰店":
                    continue

                # 获取当前日期
                t = time.strftime("%Y-%m-%d", time.localtime())

                # 使用搜索关键词作为产品名
                productName = self.keyword

                # 获取药品ID
                medicine_id = MEDICINE_ID.get(productName, "")

                # 添加数据
                # [uuid, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品ID, 药品图片, 原价, 挂网价格, 平台, 排查日期]
                datas.append(
                    [
                        "",
                        storeName,
                        storeUrl,
                        "",
                        "",
                        productName,
                        medicine_id,
                        productImg,
                        "",
                        price,
                        "京东",
                        t,
                    ]
                )
            except Exception as e:
                logger.error(f"解析京东商品数据失败: {e}")
                continue

        msg = "\n准备保存京东后 30 条数据\n"
        self.add_text.emit(msg)

        # 保存数据
        self.save_to_excel(datas, "京东")

    def jd(self, res: str) -> None:
        """
        解析京东搜索结果页面

        Args:
            res: 响应内容
        """
        try:
            html = etree.HTML(res)
            datas = []

            for li in html.xpath('//div[@id="J_goodsList"]//li'):
                try:
                    # 药品名称
                    productName = li.xpath("string(./div/div[3]/a/em)")

                    # 检查是否符合搜索条件
                    if not self.check_brand_product_name(productName):
                        continue

                    # 提取价格、图片、店铺名称等信息
                    price = li.xpath("./div/div[2]/strong/i/text()")[0]
                    productImg = (
                        "https:"
                        + li.xpath('./div[1]/div[@class="p-img"]//img/@data-lazy-img')[
                            0
                        ]
                    )
                    storeName = li.xpath('./div[1]/div[@class="p-shop"]/span/a/@title')[
                        0
                    ]
                    storeUrl = "https:" + li.xpath("./div/div[5]/span/a/@href")[0]

                    # 转为字符串类型
                    productName = str(productName)
                    price = str(price)
                    productImg = str(productImg)
                    storeName = str(storeName)
                    storeUrl = str(storeUrl)

                    # 跳过乐药师大药房旗舰店的商品
                    if storeName == "乐药师大药房旗舰店":
                        continue

                    # 获取当前日期
                    t = time.strftime("%Y-%m-%d", time.localtime())

                    # 使用搜索关键词作为产品名
                    productName = self.keyword

                    # 获取药品ID
                    medicine_id = MEDICINE_ID.get(productName, "")

                    # 添加数据
                    datas.append(
                        [
                            "",
                            storeName,
                            storeUrl,
                            "",
                            "",
                            productName,
                            medicine_id,
                            productImg,
                            "",
                            price,
                            "京东",
                            t,
                        ]
                    )
                except Exception as e:
                    logger.error(f"解析京东商品失败: {e}")
                    continue

            # 保存数据
            self.save_to_excel(datas, "京东")
        except Exception as e:
            logger.error(f"解析京东页面失败: {e}")

    def jd_saveCertificate(
        self, platform: str, storeName: str, companyName: str, url: str
    ) -> None:
        """
        保存京东店铺营业执照信息

        Args:
            platform: 平台名称
            storeName: 店铺名称
            companyName: 公司名称
            url: 营业执照图片链接
        """
        flag = True

        # 搜索Excel表格第2列找到对应的店铺名称，将营业执照链接保存到Excel表格中
        for row in self.sheet.iter_rows(
            min_row=1, max_row=self.sheet.max_row, min_col=1, max_col=12
        ):
            if row[1].value == storeName:
                flag = False

                # 更新资质名称和营业执照图片
                self.sheet.cell(row[0].row, 4).value = companyName
                self.sheet.cell(row[0].row, 5).value = url

                msg = f"\n第 {row[0].row} 行 {storeName} {companyName} {url}"
                self.add_text.emit(msg)

        if flag:
            msg = f"\n在Excel中没找到该店铺: {storeName}\n"
            self.add_text.emit(msg)
            return

        # 保存工作簿
        self.workBook.save(self.filename)

    def jd_certificate(self, res: str, url: str) -> None:
        """
        解析京东店铺营业执照

        Args:
            res: 响应内容
            url: 请求URL
        """
        try:
            html = etree.HTML(res)

            # 提取公司名称
            companyName = html.xpath('//li[@class="noBorder"][2]/span/text()')
            if not companyName:
                return

            companyName = companyName[0]

            # 提取店铺名称
            try:
                storeName = re.findall(r'document\.title="(.*?)"', res)
                if not storeName:
                    return

                storeName = str(storeName[0]).strip()

                # 检查是否包含政策信息
                if "根据国家相关政策" in companyName:
                    return

            except Exception as e:
                logger.error(f"解析店铺名称失败: {e}")
                return

            # 保存营业执照信息
            self.jd_saveCertificate("京东", storeName, companyName, url)
        except Exception as e:
            logger.error(f"解析京东营业执照失败: {e}")

    def yfw(self, res: str) -> None:
        """
        解析药房网搜索结果

        Args:
            res: 响应内容
        """
        try:
            html = etree.HTML(res)
            datas = []

            # 提取商品列表
            li_elements = html.xpath('//*[@id="slist"]/ul//li')
            for li in li_elements:
                try:
                    # 提取店铺名、链接等信息
                    storeName = li.xpath('.//div[@class="clearfix"]/a/@title')[0]
                    storeUrl = (
                        "https:" + li.xpath('.//div[@class="clearfix"]/a/@href')[0]
                    )
                    productImg = (
                        "https:" + li.xpath('.//div[@class="img"]/a/img/@src')[0]
                    )
                    price = li.xpath(
                        './/div[@class="clearfix"]/a/@data-commodity_price'
                    )[0]

                    # 获取当前日期
                    t = time.strftime("%Y-%m-%d", time.localtime())

                    # 使用搜索关键词作为产品名
                    productName = self.keyword

                    # 获取药品ID
                    medicine_id = MEDICINE_ID.get(productName, "")

                    # 添加数据
                    datas.append(
                        [
                            "",
                            storeName,
                            storeUrl,
                            storeName,  # 资质名称默认使用店铺名
                            "",
                            productName,
                            medicine_id,
                            productImg,
                            "",
                            price,
                            "药房网",
                            t,
                        ]
                    )
                except Exception as e:
                    logger.error(f"解析药房网商品失败: {e}")
                    continue

            # 保存数据
            self.save_to_excel(datas, "药房网")
        except Exception as e:
            logger.error(f"解析药房网页面失败: {e}")

    def pdd(self, res: str) -> None:
        """
        解析拼多多搜索结果页面

        Args:
            res: 响应内容
        """
        datas = []

        try:
            # 使用正则表达式提取window.rawData
            raw_data_match = re.findall(r"window\.rawData=(.*?);document", res)
            if not raw_data_match:
                return

            # 解析JSON数据
            raw_data = json.loads("".join(raw_data_match))

            # 获取商品列表
            goods_list = (
                raw_data.get("stores", {})
                .get("store", {})
                .get("data", {})
                .get("ssrListData", {})
                .get("list", [])
            )

            for data in goods_list:
                try:
                    mall_id = str(data["mallEntrance"]["mall_id"])

                    # 跳过乐药师大药房旗舰店
                    if mall_id == "397292525":
                        continue

                    storeUrl = (
                        f"https://mobile.yangkeduo.com/mall_page.html?mall_id={mall_id}"
                    )
                    productName = data.get("goodsName", "")

                    # 检查是否符合搜索条件
                    if not self.check_brand_product_name(productName):
                        continue

                    productImg = data.get("imgUrl", "")
                    price = data.get("priceInfo", "")
                    t = time.strftime("%Y-%m-%d", time.localtime())

                    # 使用搜索关键词作为产品名
                    productName = self.keyword

                    # 获取药品ID
                    medicine_id = MEDICINE_ID.get(productName, "")

                    # 添加数据
                    datas.append(
                        [
                            "",
                            "",  # 店铺名称为空
                            storeUrl,
                            "",
                            "",
                            productName,
                            medicine_id,
                            productImg,
                            "",
                            price,
                            "拼多多",
                            t,
                        ]
                    )
                except Exception as e:
                    logger.error(f"解析拼多多商品失败: {e}")
                    continue
        except Exception as e:
            logger.error(f"解析拼多多页面失败: {e}")

        # 如果有数据则保存
        if datas:
            self.save_to_excel(datas, "拼多多")

    def pdd_xhr(self, res: Dict[str, Any]) -> None:
        """
        解析拼多多XHR数据

        Args:
            res: JSON响应数据
        """
        datas = []

        try:
            # 获取商品列表
            items = res.get("items", [])

            for item in items:
                try:
                    # 获取商品数据
                    data = item.get("item_data", {}).get("goods_model", {})

                    mall_id = data.get("mall_id")
                    storeUrl = (
                        f"https://mobile.yangkeduo.com/mall_page.html?mall_id={mall_id}"
                    )
                    productName = data.get("goods_name", "")

                    # 检查是否符合搜索条件
                    if not self.check_brand_product_name(productName):
                        continue

                    productImg = data.get("hd_thumb_url", "")
                    price = data.get("price_info", "")
                    t = time.strftime("%Y-%m-%d", time.localtime())

                    # 使用搜索关键词作为产品名
                    productName = self.keyword

                    # 获取药品ID
                    medicine_id = MEDICINE_ID.get(productName, "")

                    # 添加数据
                    datas.append(
                        [
                            "",
                            "",  # 店铺名称为空
                            storeUrl,
                            "",
                            "",
                            productName,
                            medicine_id,
                            productImg,
                            "",
                            price,
                            "拼多多",
                            t,
                        ]
                    )
                except Exception as e:
                    logger.error(f"解析拼多多XHR商品失败: {e}")
                    continue
        except Exception as e:
            logger.error(f"解析拼多多XHR数据失败: {e}")

        # 如果有数据则保存
        if datas:
            self.save_to_excel(datas, "拼多多")

    def meituan(self, res: Dict[str, Any]) -> None:
        """
        解析美团搜索结果

        Args:
            res: JSON响应数据
        """
        # 检查数据有效性
        if res.get("data") is None or isinstance(res.get("data"), str):
            return

        datas = []

        try:
            # 遍历模块列表
            for module in res.get("data", {}).get("module_list", []):
                try:
                    string_data = module.get("string_data")
                    if not string_data:
                        continue

                    # 解析JSON数据
                    data = json.loads(string_data)
                    storeName = data.get("name", "")  # 药店名称

                    # 跳过乐药师大药房旗舰店
                    if storeName == "乐药师大药房旗舰店":
                        continue

                    # 只处理快递电商店铺
                    temp_str = "快递电商"
                    if temp_str not in storeName:
                        continue

                    # 去除后缀
                    storeName = storeName.replace("（快递电商）", "")

                    # 处理产品列表
                    for product in data.get("product_list", []):
                        productName = product.get("product_name", "")  # 药品名称

                        # 检查是否符合搜索条件
                        if not self.check_brand_product_name(productName):
                            continue

                        productImg = product.get("picture", "")  # 药品图片
                        price = product.get("price", "")  # 挂网价格
                        original_price = product.get("original_price", "")  # 原价
                        t = time.strftime("%Y-%m-%d", time.localtime())  # 排查日期

                        # 使用搜索关键词作为产品名
                        productName = self.keyword

                        # 获取药品ID
                        medicine_id = MEDICINE_ID.get(productName, "")

                        # 添加数据
                        datas.append(
                            [
                                "",
                                storeName,
                                "",  # 店铺链接为空
                                "",
                                "",
                                productName,
                                medicine_id,
                                productImg,
                                original_price,
                                price,
                                "美团",
                                t,
                            ]
                        )
                except Exception as e:
                    logger.error(f"解析美团模块失败: {e}")
                    continue
        except Exception as e:
            logger.error(f"解析美团数据失败: {e}")

        # 如果有数据则保存并输出信息
        if datas:
            self.add_text.emit(str(datas))
            self.save_to_excel(datas, "美团")

    def taobao(self, res: str) -> None:
        """
        解析淘宝天猫搜索结果

        Args:
            res: 响应内容
        """
        datas = []

        try:
            # 将JSONP转为JSON
            res = re.sub(r"mtopjsonp\d+\(", "", res)
            res = res[:-1]  # 去掉最后一个括号

            # 解析JSON数据
            data = json.loads(res)

            # 获取商品列表
            items = data.get("data", {}).get("itemsArray", [])

            for item in items:
                try:
                    storeName = item.get("shopInfo", {}).get("title", "")  # 店铺名称

                    # 跳过乐药师大药房旗舰店
                    if storeName == "乐药师大药房旗舰店":
                        continue

                    storeUrl = "https:" + item.get("shopInfo", {}).get(
                        "url", ""
                    )  # 店铺链接
                    productName = item.get("title", "")  # 药品名称

                    # 提取中文字符
                    chinese_pattern = re.compile(r"[\u4e00-\u9fff]+", re.UNICODE)
                    chinese_text = re.findall(chinese_pattern, productName)

                    # 合并中文字符并截取
                    if chinese_text:
                        productName = "".join(chinese_text)
                        if len(productName) > 3:
                            productName = productName[:-3]

                    # 获取价格和图片
                    price = item.get("priceShow", {}).get("price", "")
                    productImg = item.get("pic_path", "")
                    t = time.strftime("%Y-%m-%d", time.localtime())

                    # 使用搜索关键词作为产品名
                    productName = self.keyword

                    # 获取药品ID
                    medicine_id = MEDICINE_ID.get(productName, "")

                    # 添加数据
                    datas.append(
                        [
                            "",
                            storeName,
                            storeUrl,
                            "",
                            "",
                            productName,
                            medicine_id,
                            productImg,
                            "",
                            price,
                            "淘宝天猫",
                            t,
                        ]
                    )
                except Exception as e:
                    logger.error(f"解析淘宝商品失败: {e}")
                    continue
        except Exception as e:
            logger.error(f"解析淘宝数据失败: {e}")

        # 如果有数据则保存
        if datas:
            self.save_to_excel(datas, "淘宝天猫")

    def ele(self, res: Dict[str, Any]) -> None:
        """
        解析饿了么搜索结果

        Args:
            res: JSON响应数据
        """
        datas = []

        try:
            # 获取结果数据
            data = res.get("data", {}).get("result", {})

            # 处理数据格式
            if isinstance(data, list) and data:
                data = data[0]

            # 遍历商品列表
            for item in data.get("listItems", []):
                try:
                    restaurant = item.get("info", {}).get("restaurant")

                    if restaurant is None:
                        continue

                    storeName = restaurant.get("name", "")  # 药店名称

                    # 跳过乐药师大药房旗舰店
                    if storeName == "乐药师大药房旗舰店":
                        continue

                    # 遍历食品列表（药品）
                    for food in item.get("info", {}).get("foods", []):
                        productName = food.get("name", "")  # 药品名

                        # 检查是否符合搜索条件
                        if not self.check_brand_product_name(productName):
                            continue

                        productImg = food.get("imagePath", "")  # 药品图片
                        price = food.get("price", "")  # 挂网价格
                        t = time.strftime("%Y-%m-%d", time.localtime())

                        # 使用搜索关键词作为产品名
                        productName = self.keyword

                        # 获取药品ID
                        medicine_id = MEDICINE_ID.get(productName, "")

                        # 添加数据
                        datas.append(
                            [
                                "",
                                storeName,
                                "",  # 店铺链接为空
                                "",
                                "",
                                productName,
                                medicine_id,
                                productImg,
                                "",
                                price,
                                "饿了么",
                                t,
                            ]
                        )
                except Exception as e:
                    logger.error(f"解析饿了么商品失败: {e}")
                    continue
        except Exception as e:
            logger.error(f"解析饿了么数据失败: {e}")

        # 如果有数据则保存
        if datas:
            self.save_to_excel(datas, "饿了么")

    def request(self, flow: http.HTTPFlow) -> None:
        """
        处理请求

        Args:
            flow: 请求流
        """
        url = flow.request.url

        # 记录京东API的Cookie
        if "api.m.jd.com" in url:
            logger.info(flow.request.headers.get("Cookie"))

    def response(self, flow: http.HTTPFlow) -> None:
        """
        处理响应

        Args:
            flow: 响应流
        """
        url = flow.request.url

        # 处理各平台的响应数据

        # 京东搜索结果
        if re.match("https://search.jd.com/Search", url):
            res = flow.response.text
            msg = f"\n京东 {url[:50]}\n"
            self.add_text.emit(msg)
            self.jd(res)

        # 京东后30条数据
        elif re.match(
            r"https://api.m.jd.com/\?appid=search-pc-java&functionId=pc_search_s_new*",
            url,
        ):
            res = flow.response.text
            if not res:
                return

            msg = f"京东后 30 条数据 {url[:50]}\n"
            self.add_text.emit(msg)
            self.parsejd2HTML(res)

        # 京东营业执照
        elif re.match("https://mall.jd.com/showLicence*", url):
            res = flow.response.text
            msg = f" 京东营业执照 {url[:50]}"
            self.add_text.emit(msg)
            self.jd_certificate(res, url)

        # 药房网
        elif re.match(r"https://www.yaofangwang.com/medicine/\d+/*", url):
            res = flow.response.text
            msg = f"药房网 {url[:50]}"
            self.add_text.emit(msg)
            self.yfw(res)

        # 拼多多搜索结果
        elif re.match(r"https://mobile.yangkeduo.com/search_result.html", url):
            res = flow.response.text
            if not res:
                return

            msg = f"\n拼多多 {url[:50]}\n"
            self.add_text.emit(msg)
            self.pdd(res)

        # 拼多多XHR数据
        elif re.match(r"https://mobile.yangkeduo.com/proxy/api/search*", url):
            try:
                res = flow.response.json()
                if not res:
                    return

                msg = f"\n拼多多 xhr {url[:50]}\n"
                self.add_text.emit(msg)
                self.pdd_xhr(res)
            except Exception as e:
                logger.error(f"解析拼多多XHR响应失败: {e}")
                return

        # 美团
        elif re.match("https://i.waimai.meituan.com/openh5/search/globalpage*", url):
            try:
                res = flow.response.json()
                msg = f"\n美团 {url[:50]}\n"
                self.add_text.emit(msg)
                self.meituan(res)
            except Exception as e:
                logger.error(f"解析美团响应失败: {e}")
                return

        # 淘宝天猫
        elif re.match(
            "https://h5api.m.taobao.com/h5/mtop.relationrecommend.wirelessrecommend.recommend/2.0/*",
            url,
        ):
            res = flow.response.text
            msg = f'\n淘宝天猫 {url.split("?")[0]}\n'
            self.add_text.emit(msg)
            self.taobao(res)

        # 饿了么
        elif re.match(
            "https://waimai-guide.ele.me/h5/mtop.relationrecommend.elemetinyapprecommend.recommend*",
            url,
        ):
            try:
                res = flow.response.json()

                # 检查数据有效性
                if (
                    not res
                    or not res.get("data")
                    or not res.get("data").get("result")
                    or not res.get("data").get("result")[0].get("listItems")
                ):
                    return

                msg = f'\n饿了么 {url.split("?")[0]}\n'
                self.add_text.emit(msg)
                self.ele(res)
            except Exception as e:
                logger.error(f"解析饿了么响应失败: {e}")
                return
