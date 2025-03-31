import json
import re
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional

import httpx
import openpyxl
import pandas as pd
import pyperclip
import shortuuid
from loguru import logger
from lxml import etree
from mitmproxy import http
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from PySide6.QtCore import QThread, Signal

from medicineID import MEDICINE_ID


class Addon(QThread):
    add_text = Signal(str)

    def __init__(self):
        super().__init__()

        self.h = httpx.Client()

        self.keyword = None

        self.sheet = None
        self.workBook = None

        self.filename: Optional[Path] = None

        self.thread = ThreadPoolExecutor(max_workers=5)  # 线程池

        # self.brand_name_not = ""
        # self.product_name_not = ""

    def createExcel(self, filename):
        """
        创建 Excel 文件
        :param filename: 文件名
        :return: 无
        """
        self.filename = filename

        # 判断文件是否存在, 不存在则新建
        if not filename.exists():
            self.workBook = openpyxl.Workbook()  # 创建一个工作簿对象
        else:
            self.workBook = openpyxl.load_workbook(
                filename, keep_vba=True
            )  # 打开 Excel 表格格

        self.sheet = self.workBook.active  # 选取第一个sheet

    # 保存数据到 Excel 文件
    def save_to_excel(self, datas: list, tag=None):
        """
        保存数据到 Excel 文件
        :param datas:
        :param tag: 标签-指明哪个平台
        :return: 无
        """
        df = None

        # 判断文件是否存在, 不存在则新建
        if not self.filename.exists():
            workBook = openpyxl.Workbook()  # 创建一个工作簿对象
        else:
            workBook = openpyxl.load_workbook(
                self.filename, keep_vba=True
            )  # 打开 Excel 表格格

            df = pd.read_excel(self.filename)

        sheet = workBook.active  # 选取第一个sheet

        if datas is None or len(datas) == 0:
            return

        max_row = sheet.max_row
        i = 1

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

        save_flag: bool = True

        for data in datas:
            # 重复数据不保存 - 根据药店名称、店铺主页、药品名、挂网价格、平台判断
            if df is not None:
                temp_df = df[
                    (df["药店名称"] == data[1])
                    & (df["店铺主页"] == data[2])
                    & (df["药品名"] == data[5])
                    &
                    # (df["药品图片"] == data[6]) &
                    (df["挂网价格"] == float(data[8]))
                    & (df["平台"] == data[9])
                ]

                if not temp_df.empty:
                    continue

            # 生成一个短 UUID
            short_uuid = shortuuid.uuid()
            data[0] = short_uuid

            sheet.append(data)
            i += 1

            save_flag = False

        if save_flag:
            msg = f"{self.filename.stem} {tag}-没有数据需要保存"
            logger.info(msg)
            return

        workBook.save(self.filename)

        msg = f"\n\n{self.filename.stem} {tag}-保存了{i - 1}条, 数据总条数: {sheet.max_row - 1}\n\n"

        self.add_text.emit(msg)

    def check_brand_product_name(self, name: str) -> bool:
        keywords = self.keyword.split(" ")
        self.brand_name, self.medicine_name = keywords[:-1], keywords[-1]

        medicine_name_temp = (
            self.medicine_name[:3]
            if len(self.medicine_name) > 2
            else self.medicine_name
        )

        # 药品名和品牌名包含其中一个即可
        if medicine_name_temp not in name:
            return False

        if "一口" in self.brand_name:
            return "一口" in name
        else:
            # 至少要包含其中一个品牌名
            for brand in self.brand_name:
                if brand in name:
                    return True

        return False

    def parsejd2HTML(self, html_str: str):
        """
        1. 一行行删除 jd.html 文件中的内容，直到 <li data-sku= 有这个标签，然后再删除这一行之前的内容，即可得到京东的商品列表
        2. 删除倒数几行的 script 标签
        """

        # 删除不需要的内容
        for i in range(len(html_str)):
            if "li data-sku=" in html_str[i]:
                html_str = html_str[i:]
                break

        for i in range(len(html_str)):
            if "<script>" in html_str[i]:
                html_str = html_str[:i]
                break

        # 解析文件
        html = etree.HTML(html_str)

        # 商品列表
        try:
            li = html.xpath("//li")
        except:
            return

        datas = []
        # 解析商品信息
        for i in li:
            # 商品名称
            productName = i.xpath("string(./div/div[3]//em)").split("\n")[0]

            # 商品价格
            price = i.xpath(".//div/div[2]/strong/i/text()")[0]

            # 商品图片
            productImg = "https:" + i.xpath("./div[1]//img/@data-lazy-img")[0]

            # 店铺名称
            storeName = i.xpath('.//div[@class="p-shop"]/span/a/@title')[0]

            # 店铺链接
            storeUrl = "https:" + i.xpath(".//div/div[5]/span/a/@href")[0]

            productName = str(productName)
            price = str(price)
            productImg = str(productImg)
            storeName = str(storeName)

            if storeName == "乐药师大药房旗舰店":
                continue

            storeUrl = str(storeUrl)
            t = time.strftime("%Y-%m-%d", time.localtime())

            # productName = self.product_name_not
            productName = self.keyword

            # 序号, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品图片, 原价, 挂网价格, 平台, 排查日期
            datas.append(
                [
                    "",
                    storeName,
                    storeUrl,
                    "",
                    "",
                    productName,
                    productImg,
                    "",
                    price,
                    "京东",
                    t,
                ]
            )

        msg = "\n准备保存京东后 30 条数据\n"
        self.add_text.emit(msg)

        # self.thread.submit(self.save_to_excel, datas, '京东')
        self.save_to_excel(datas, "京东")
        # self.save.to_excel(datas, "京东")

    # 京东
    def jd(self, res):
        html = etree.HTML(res)

        datas = []

        for li in html.xpath('//div[@id="J_goodsList"]//li'):
            # data_sku = li.get("data-sku")

            # 药品名称
            productName = li.xpath("string(./div/div[3]/a/em)")

            if not self.check_brand_product_name(productName):
                continue

            price = li.xpath("./div/div[2]/strong/i/text()")[0]

            productImg = (
                "https:"
                + li.xpath('./div[1]/div[@class="p-img"]//img/@data-lazy-img')[0]
            )  # .

            storeName = li.xpath('./div[1]/div[@class="p-shop"]/span/a/@title')[0]

            storeUrl = "https:" + li.xpath("./div/div[5]/span/a/@href")[0]

            productName = str(productName)
            price = str(price)
            productImg = str(productImg)
            storeName = str(storeName)

            if storeName == "乐药师大药房旗舰店":
                continue

            storeUrl = str(storeUrl)
            t = time.strftime("%Y-%m-%d", time.localtime())

            # productName = self.product_name_not
            productName = self.keyword

            # 药品ID
            medicine_id = MEDICINE_ID.get(productName, "")

            # 序号, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品ID, 药品图片, 原价, 挂网价格, 平台, 排查日期
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

        # self.thread.submit(self.save_to_excel, datas, "京东")
        self.save_to_excel(datas, "京东")
        # self.save.to_excel(datas, "京东")

    def jd_saveCertificate(self, platform, storeName, companyName, url):
        flag = True

        # 搜索 Excel 表格第 1 列找到对应的店铺名称以及第 8 列是拼多多，将营业执照链接保存到 Excel 表格中
        for row in self.sheet.iter_rows(
            min_row=1, max_row=self.sheet.max_row, min_col=0, max_col=11
        ):
            if row[1].value == storeName:
                flag = False

                self.sheet.cell(row[0].row, 4).value = companyName
                self.sheet.cell(row[0].row, 5).value = url

                msg = f"\n第 {row[0].row} 行 {storeName} {companyName} {url}"
                self.add_text.emit(msg)

        if flag:
            msg = f"\n 在 Excel 中没找到该店铺: {storeName}\n"
            self.add_text.emit(msg)

            return

        self.workBook.save(self.filename)

    def jd_certificate(self, res, url):
        html = etree.HTML(res)

        companyName = html.xpath('//li[@class="noBorder"][2]/span/text()')[0]

        try:
            storeName = re.findall(r'document\.title="(.*?)"', res)[0]
            storeName = str(storeName)
            storeName = storeName.strip()

            if "根据国家相关政策" in companyName:
                return

        except:
            return

        # self.thread.submit(self.jd_saveCertificate, '京东', storeName, companyName, url)
        self.jd_saveCertificate("京东", storeName, companyName, url)

    # 药房网
    def yfw(self, res):
        html = etree.HTML(res)

        datas = []

        li = html.xpath('//*[@id="slist"]/ul//li')
        for i in li:
            storeName = i.xpath('.//div[@class="clearfix"]/a/@title')[0]
            storeUrl = "https:" + i.xpath('.//div[@class="clearfix"]/a/@href')[0]

            # productName = i.xpath(
            #     './/div[@class="info"]/h3/a[@class="sc_medicine"]/@title'
            # )[0]
            # productName = self.brand_name_not + productName

            productName = self.keyword

            productImg = "https:" + i.xpath('.//div[@class="img"]/a/img/@src')[0]
            price = i.xpath('.//div[@class="clearfix"]/a/@data-commodity_price')[0]
            t = time.strftime("%Y-%m-%d", time.localtime())

            # productName = self.product_name_not
            productName = self.keyword

            # 药品ID
            medicine_id = MEDICINE_ID.get(productName, "")

            # 序号, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品ID, 药品图片, 原价, 挂网价格, 平台, 排查日期
            datas.append(
                [
                    "",
                    storeName,
                    storeUrl,
                    storeName,
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

        # self.thread.submit(self.save_to_excel, datas, '药房网')
        self.save_to_excel(datas, "药房网")
        # self.save.to_excel(datas, "药房网")

    # 拼多多
    def pdd(self, res):
        datas = []

        # 使用正则表达式提取 window.rawData
        res = re.findall(r"window\.rawData=(.*?);document", res)

        try:
            res = json.loads("".join(res))
        except:
            msg = "json.loads error"
            self.add_text.emit(msg)
            return

        # 解析数据
        for data in (
            res.get("stores")
            .get("store")
            .get("data")
            .get("ssrListData")
            .get("list", [])
        ):
            storeName = ""  # 店铺名称

            if storeName == "乐药师大药房旗舰店":
                continue

            mall_id = str(data["mallEntrance"]["mall_id"])

            # 乐药师大药房旗舰店
            if mall_id == "397292525":
                continue

            storeUrl = f"https://mobile.yangkeduo.com/mall_page.html?mall_id={mall_id}"  # 店铺链接

            productName = data.get("goodsName")  # 药品名称
            if not self.check_brand_product_name(productName):
                continue

            productImg = data.get("imgUrl")
            price = data.get("priceInfo")  # 拼团价
            t = time.strftime("%Y-%m-%d", time.localtime())  # 排查日期

            # productName = self.product_name_not

            productName = self.keyword

            # 药品ID
            medicine_id = MEDICINE_ID.get(productName, "")

            # 序号, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品ID, 药品图片, 原价, 挂网价格, 平台, 排查日期
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
                    "拼多多",
                    t,
                ]
            )

        if datas is None:
            return

        # self.thread.submit(self.save_to_excel, datas, '拼多多')
        self.save_to_excel(datas, "拼多多")
        # self.save.to_excel(datas, "拼多多")

    # 拼多多 xhr 数据
    def pdd_xhr(self, res):
        datas = []
        for data in res.get("items", []):
            data = data.get("item_data").get("goods_model")
            storeName = ""  # 店铺名称

            if storeName == "乐药师大药房旗舰店":
                continue

            mall_id = data.get("mall_id")
            storeUrl = f"https://mobile.yangkeduo.com/mall_page.html?mall_id={mall_id}"  # 店铺链接

            productName = data.get("goods_name")  # 药品名称

            if not self.check_brand_product_name(productName):
                continue

            productImg = data.get("hd_thumb_url")  # 药品图片
            price = data.get("price_info")  # 拼团价
            t = time.strftime("%Y-%m-%d", time.localtime())  # 排查日期

            # productName = self.product_name_not

            productName = self.keyword

            # 药品ID
            medicine_id = MEDICINE_ID.get(productName, "")

            # 序号, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品图片, 原价, 挂网价格, 平台, 排查日期
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
                    "拼多多",
                    t,
                ]
            )

        if datas is None:
            return

        # self.thread.submit(self.save_to_excel, datas, '拼多多')
        self.save_to_excel(datas, "拼多多")
        # self.save.to_excel(datas, "拼多多")

    def meituan(self, res):
        """
        解析美团
        :param res: 服务器的响应文本
        :return: 无
        """
        if res.get("data") is None or type(res.get("data")) is str:
            return

        datas = []

        for i in res.get("data").get("module_list"):
            string_data = i.get("string_data")

            data = json.loads(string_data)
            storeName = data.get("name")  # 药店名称

            if storeName == "乐药师大药房旗舰店":
                continue

            temp_str = "快递电商"
            if temp_str not in storeName:
                continue

            storeName = storeName.replace("（快递电商）", "")

            for product in data.get("product_list"):
                productName = product.get("product_name")  # 药品名称
                if not self.check_brand_product_name(productName):
                    continue

                # product_sku_id = product.get("product_sku_id")  # 商品 ID

                storeUrl = ""
                productImg = product.get("picture")  # 药品图片
                price = product.get("price")  # 挂网价格
                original_price = product.get("original_price")  # 原价
                t = time.strftime("%Y-%m-%d", time.localtime())  # 排查日期

                # productName = self.product_name_not

                productName = self.keyword

                # 药品ID
                medicine_id = MEDICINE_ID.get(productName, "")

                # 序号, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品名, 药品图片, 原价, 挂网价格, 平台, 排查日期
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
                        original_price,
                        price,
                        "美团",
                        t,
                        # product_sku_id,
                    ]
                )

        if datas is None or len(datas) == 0:
            return

        msg = datas
        self.add_text.emit(msg)

        # self.thread.submit(self.save_to_excel, datas, '美团')
        self.save_to_excel(datas, "美团")
        # self.save.to_excel(datas, "美团")

    def meituan_certificate(self, res):
        img_url = res.get("data").get("poi_qualify_details", [])[0].get("qualify_pic")

        pyperclip.copy(img_url)

        msg = img_url
        self.add_text.emit(msg)
        # logger.info(img_url)

    # 淘宝天猫
    def taobao(self, res):
        datas = []

        # 将 res 转为 json 格式
        res = re.sub(r"mtopjsonp\d+\(", "", res)

        # 把 res 最后一个括号去掉
        res = res[:-1]

        res = "".join(res)

        res = json.loads(res)

        for data in res.get("data").get("itemsArray"):
            storeName = data.get("shopInfo").get("title")  # 店铺名称
            if storeName == "乐药师大药房旗舰店":
                continue

            storeUrl = "https:" + data.get("shopInfo").get("url")  # 店铺链接

            # 提取药品名称
            productName = data.get("title")  # 药品名称

            # 正则表达式匹配中文字符
            chinese_pattern = re.compile(r"[\u4e00-\u9fff]+", re.UNICODE)

            # 使用正则表达式搜索并提取中文字符
            chinese_text = re.findall(chinese_pattern, productName)

            # 将提取出的中文字符合并为一个字符串
            productName = "".join(chinese_text)[:-3]

            # if not self.check_brand_product_name(productName):
            #     continue

            price = data.get("priceShow").get("price")
            productImg = data.get("pic_path")
            t = time.strftime("%Y-%m-%d", time.localtime())

            # productName = self.product_name_not

            productName = self.keyword

            # 药品ID
            medicine_id = MEDICINE_ID.get(productName, "")

            # 序号, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品ID, 药品图片, 原价, 挂网价格, 平台, 排查日期
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

        # self.thread.submit(self.save_to_excel, datas, '淘宝天猫')
        self.save_to_excel(datas, "淘宝天猫")
        # self.save.to_excel(datas, "淘宝天猫")

    # 饿了么
    def ele(self, res):
        datas = []

        data = res.get("data", {}).get("result", {})
        if isinstance(data, list):
            data = data[0]

        for items in data.get("listItems", []):
            resturant = items.get("info").get("restaurant")

            if resturant is None:
                continue

            storeName = resturant.get("name")  # 药店名称

            if storeName == "乐药师大药房旗舰店":
                continue

            for food in items.get("info").get("foods"):
                productName = food.get("name")  # 药品名
                if not self.check_brand_product_name(productName):
                    continue

                storeUrl = ""
                productImg = food.get("imagePath")  # 药品图片
                price = food.get("price")  # 挂网价格
                t = time.strftime("%Y-%m-%d", time.localtime())

                # productName = self.product_name_not

                productName = self.keyword

                # 药品ID
                medicine_id = MEDICINE_ID.get(productName, "")

                # 序号, 药店名称, 店铺主页, 资质名称, 营业执照图片, 药品名, 药品ID, 药品图片, 原价, 挂网价格, 平台, 排查日期
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
                        "饿了么",
                        t,
                    ]
                )

        if datas is None:
            return

        # self.thread.submit(self.save_to_excel, datas, '饿了么')
        self.save_to_excel(datas, "饿了么")
        # self.save.to_excel(datas, "饿了么")

    def request(self, flow: http.HTTPFlow) -> None:
        url = flow.request.url

        if "api.m.jd.com" in url:
            logger.info(flow.request.headers.get("Cookie"))

    def response(self, flow: http.HTTPFlow):
        url = flow.request.url
        headers = flow.response.headers

        # 京东
        if re.match("https://search.jd.com/Search", url):
            res = flow.response.text
            msg = f"\n京东 {url[:50]}\n"
            self.add_text.emit(msg)

            self.jd(res)

        # 京东 后 30 条数据
        elif re.match(
            r"https://api.m.jd.com/\?appid=search-pc-java&functionId=pc_search_s_new*",
            url,
        ):
            res = flow.response.text

            if res is None:
                return

            msg = f"京东后 30 条数据 {url[:50]}\n"
            self.add_text.emit(msg)

            self.parsejd2HTML(res)

        # 京东营业执照
        elif re.match("https://mall.jd.com/showLicence*", url):
            res = flow.response.text

            msg = f" 京东营业执照 {url[:50]}"
            self.add_text.emit(msg)

            # self.thread.submit(self.jd_certificate, res, url)
            self.jd_certificate(res, url)

        # 药房网
        elif re.match(r"https://www.yaofangwang.com/medicine/\d+/*", url):
            res = flow.response.text
            msg = f"药房网 {url[:50]}"
            self.add_text.emit(msg)

            self.yfw(res)

        # 验证码
        elif re.findall("sys/vc/createVerifyCode.html", url):
            # p.screenshot('ocr.png', region=(743, 564, 130, 50))
            # with open('./ocr.png', 'rb')as f:
            #     img_bytes = f.read()
            # code = ocr.classification(img_bytes)
            # conn.set('code', code)
            ...

        # 拼多多
        elif re.match(r"https://mobile.yangkeduo.com/search_result.html", url):
            res = flow.response.text

            if res is None:
                return

            # logger.info(f'\n拼多多 {url[:50]}\n')
            msg = f"\n拼多多 {url[:50]}\n"
            self.add_text.emit(msg)

            self.pdd(res)

        # 拼多多 xhr 数据
        elif re.match(r"https://mobile.yangkeduo.com/proxy/api/search*", url):
            res = flow.response.json()

            if res is None:
                return

            msg = f"\n拼多多 xhr {url[:50]}\n"
            self.add_text.emit(msg)

            self.pdd_xhr(res)

        # 美团
        elif re.match("https://i.waimai.meituan.com/openh5/search/globalpage*", url):
            res = flow.response.json()
            msg = f"\n美团 {url[:50]}\n"
            self.add_text.emit(msg)

            self.meituan(res)

        # 美团营业执照
        elif re.match("https://yiyao-h5.meituan.com/wedrug/v2/poi/qualification*", url):
            res = flow.response.json()

            msg = f"\n美团营业执照 {url[:50]}\n"
            self.add_text.emit(msg)

            # self.thread.submit(self.meituan_certificate, res)
            self.meituan_certificate(res)

        # 淘宝天猫
        elif re.match(
            "https://h5api.m.taobao.com/h5/mtop.relationrecommend.wirelessrecommend.recommend/2.0/*",
            url,
        ):
            res = flow.response.text

            msg = f'\n淘宝天猫 {url.split("?")[0]}\n'
            self.add_text.emit(msg)

            # self.thread.submit(self.taobao, res)
            self.taobao(res)

        # 饿了么
        elif re.match(
            "https://waimai-guide.ele.me/h5/mtop.relationrecommend.elemetinyapprecommend.recommend*",
            url,
        ):
            # res = flow.response.text
            try:
                res = flow.response.json()
            except:
                return

            if (
                res is None
                or res.get("data") is None
                or res.get("data").get("result") is None
            ):
                return

            if res.get("data").get("result")[0].get("listItems") is None:
                return

            msg = f'\n饿了么 {url.split("?")[0]}\n'
            self.add_text.emit(msg)

            self.ele(res)
