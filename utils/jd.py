import random
import time
from pathlib import Path

from DrissionPage import Chromium
from DrissionPage.common import Keys
from lxml import etree
from PySide6.QtCore import Signal

import shortuuid
from utils.medicineID import MEDICINE_ID
from utils.save import Save


class JD:
    logInfo = Signal(str)

    def __init__(self, save_dir: Path):
        self.medicine_name = None
        self.brand_name = None
        self.keyword = None
        self.save_dir = save_dir

        self.save = Save()

        self.bro = Chromium()

    def check_brand_product_name(self, name: str) -> bool:
        medicine_name_temp = (
            self.medicine_name[:3]
            if len(self.medicine_name) > 2
            else self.medicine_name
        )

        # 药品名和品牌名包含其中一个即可
        if medicine_name_temp in name:
            return True

        # 至少要包含其中一个品牌名
        for brand in self.brand_name:
            if brand in name:
                return True

        return False

    @staticmethod
    def extract_data(html_obj, xpath_str: str):
        try:
            return html_obj.xpath(xpath_str)[0]
        except Exception as e:
            return ""

    def parse_search(self, html_str: str, filename: Path):
        html = etree.HTML(html_str)

        datas = []

        try:
            for li in html.xpath('//div[@id="J_goodsList"]//li'):
                # 药品名称
                productName = li.xpath("string(./div/div[3]/a/em)")

                if not self.check_brand_product_name(productName):
                    continue

                price = self.extract_data(li, "./div/div[2]/strong/i/text()")

                productImg = "https:" + self.extract_data(
                    li, './div[1]/div[@class="p-img"]//img/@data-lazy-img'
                )

                storeName = self.extract_data(
                    li, './div[1]/div[@class="p-shop"]/span/a/@title'
                )
                if storeName == "乐药师大药房旗舰店":
                    continue

                storeUrl = "https:" + self.extract_data(li, "./div/div[5]/span/a/@href")

                productName = self.keyword

                # 获取药品ID
                medicine_id = MEDICINE_ID.get(productName, "")

                price = str(price)
                productImg = str(productImg)
                storeName = str(storeName)

                storeUrl = str(storeUrl)
                t = time.strftime("%Y-%m-%d", time.localtime())

                # 添加数据
                # [uuid, 药店名称, 店铺主页, 资质名称, 药品名, 药品ID, 药品图片, 挂网价格, 平台, 排查日期]
                datas.append(
                    [
                        shortuuid.uuid(),
                        storeName,
                        storeUrl,
                        "",
                        productName,
                        medicine_id,
                        productImg,
                        price,
                        "京东",
                        t,
                    ]
                )

            if not datas:
                return

            self.save.to_excel(filename, datas, "京东")

        except Exception as e:
            self.logInfo.emit(f"{self.keyword} {e}")

    def parse_xhr(self, html_str: str, filename: Path):
        """
        1. 一行行删除 jd.html 文件中的内容，直到 <li data-sku= 有这个标签，然后再删除这一行之前的内容，即可得到京东的商品列表
        2. 删除倒数几行的 script 标签
        """
        content = html_str

        try:
            # 删除不需要的内容
            for i in range(len(content)):
                if "li data-sku=" in content[i]:
                    content = content[i:]
                    break

            for i in range(len(content)):
                if "<script>" in content[i]:
                    content = content[:i]
                    break

            # 解析文件
            html = etree.HTML(content)

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

                if not self.check_brand_product_name(productName):
                    continue

                # 商品价格
                price = i.xpath(".//div/div[2]/strong/i/text()")[0]

                # 商品图片
                productImg = "https:" + i.xpath("./div[1]//img/@data-lazy-img")[0]

                # 店铺名称
                storeName = i.xpath('.//div[@class="p-shop"]/span/a/@title')[0]
                if storeName == "乐药师大药房旗舰店":
                    continue

                # 店铺链接
                storeUrl = "https:" + i.xpath(".//div/div[5]/span/a/@href")[0]

                productName = self.keyword

                # 获取药品ID
                medicine_id = MEDICINE_ID.get(productName, "")

                price = str(price)
                productImg = str(productImg)
                storeName = str(storeName)

                storeUrl = str(storeUrl)
                t = time.strftime("%Y-%m-%d", time.localtime())

                # 添加数据
                # [uuid, 药店名称, 店铺主页, 资质名称, 药品名, 药品ID, 药品图片, 挂网价格, 平台, 排查日期]
                datas.append(
                    [
                        shortuuid.uuid(),
                        storeName,
                        storeUrl,
                        "",
                        productName,
                        medicine_id,
                        productImg,
                        price,
                        "京东",
                        t,
                    ]
                )

            if not datas:
                return

            self.save.to_excel(filename, datas, "京东")

        except Exception as e:
            self.logInfo.emit(f"解析京东搜索结果出错 {self.keyword} {e}")

    def scroll_down(self, tab):
        for _ in range(random.randint(4, 12)):
            tab.scroll.down(random.randint(50, 1000))
            time.sleep(random.uniform(0.0, 2.0))

        self.logInfo.emit("滑动到最底部")
        tab.scroll.to_bottom()

    def search(self, keyword: str):
        self.keyword = keyword

        keywords = self.keyword.split(" ")
        self.brand_name, self.medicine_name = keywords[:-1], keywords[-1]

        filename = self.save_dir / f"{self.keyword}.xlsx"

        self.save.logInfo = self.logInfo

        tab = self.bro.latest_tab

        tab.listen.start("search.jd.com/Search")

        self.logInfo.emit("\n\n打开京东首页")
        tab.get("https://www.jd.com/")

        # 输入搜索关键字
        self.logInfo.emit(f"搜索关键字: {self.keyword}")
        ele_input = tab.ele("#key", timeout=60)
        ele_input.click.multi(2)
        ele_input.input(self.keyword)
        ele_input.input(Keys.ENTER)

        # 监听搜索结果
        self.logInfo.emit("开始监听搜索结果...")

        res = tab.listen.wait(timeout=9)

        if res:
            html_str = res.response.body
        else:
            html_str = tab.html

        self.logInfo.emit("监听到搜索结果")

        # 解析搜索结果
        self.logInfo.emit("解析京东搜索结果")
        self.parse_search(html_str, filename=filename)

        # 往下滑动
        tab.listen.set_targets(
            "api.m.jd.com/?appid=search-pc-java&functionId=pc_search_s_new"
        )

        # 滑动
        self.logInfo.emit("往下滑动")
        self.scroll_down(tab)

        if not res:
            return

        for package in tab.listen.steps(timeout=2):
            res = package.response.body

            if not res or len(res) < 140:
                continue

            # 解析 xhr 结果
            self.logInfo.emit("解析京东 xhr 结果")
            self.parse_xhr(res, filename=filename)

        # 检查是否有下一页
        ele_next_page = tab.ele(".pn-next", timeout=2)

        while ele_next_page:
            try:
                self.logInfo.emit("点击下一页")
                ele_next_page = tab.ele(".pn-next", timeout=2)

                if ele_next_page:
                    ele_next_page.click()

                    # 滑动
                    self.logInfo.emit("往下滑动")
                    self.scroll_down(tab)

            except Exception as e:
                self.logInfo.emit(f"点击下一页出错")
                print(e)

        for package in tab.listen.steps(timeout=2):
            res = package.response.body

            if not res or len(res) < 140:
                continue

            # 解析 xhr 结果
            self.logInfo.emit("解析京东 xhr 结果")
            self.parse_xhr(res, filename=filename)
