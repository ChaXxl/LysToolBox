import json
import random
import re
import time
from pathlib import Path

import shortuuid
from DrissionPage import Chromium
from PySide6.QtCore import Signal

from utils.medicineID import MEDICINE_ID
from utils.save import Save


class TB:
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

    def parse(self, html_str: str, filename: Path = None):
        if type(html_str) is not str:
            return

        datas = []

        try:
            # 将 res 转为 json 格式
            res = re.sub(r"mtopjsonp\d+\(", "", html_str)

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

                if not self.check_brand_product_name(productName):
                    continue

                price = data.get("priceShow").get("price")
                productImg = data.get("pic_path")
                t = time.strftime("%Y-%m-%d", time.localtime())

                productName = self.keyword

                # 获取药品ID
                medicine_id = MEDICINE_ID.get(productName, "")

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
                        "淘宝",
                        t,
                    ]
                )

            if not datas:
                return

            self.save.to_excel(filename, datas, "淘宝天猫")

        except Exception as e:
            self.logInfo.emit(f"解析淘宝搜索结果出错: {self.keyword} {e}")

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

        # 监听搜索结果
        tab.listen.start(
            "h5api.m.taobao.com/h5/mtop.relationrecommend.wirelessrecommend.recommend/2.0"
        )

        self.logInfo.emit("\n\n打开淘宝首页")
        tab.get(
            f"https://s.taobao.com/search?commend=all&ie=utf8&initiative_id=tbindexz_20170306&page=1&preLoadOrigin=https%3A%2F%2Fwww.taobao.com&q={self.keyword}&search_type=item&sourceId=tb.index&spm=a21bo.jianhua%2Fa.201856.d13&ssid=s5-e&tab=all"
        )

        try:
            # 检测是否有弹窗
            if tab.ele("@text()=恭喜获得惊喜红包", timeout=0.5):
                tab.ele(
                    'xpath://*[@id="J_TBPC_POP_home"]/div/div/div/div/div/div/div[4]',
                    timeout=0.5,
                ).click()

            if tab.ele("@text()=抢年货购物惊喜券", timeout=0.5):
                tab.ele(".cpCloseIcon", timeout=0.5).click()

            if tab.ele("@text()=百亿补贴红包", timeout=0.5):
                tab.ele(".closeIconWrapper", timeout=0.5).click()

        except Exception as e:
            self.logInfo.emit(f"关闭弹窗出错: {e}")

        while tab.ele("text:验证码", timeout=0.5):
            ...

        while tab.ele("#nocaptcha", timeout=0.5):
            ...

        while tab.ele("text:验证码", timeout=0.5):
            ...

        while tab.ele("#nocaptcha", timeout=0.5):
            ...

        for package in tab.listen.steps(timeout=12):
            res = package.response.body

            if not res or len(res) < 200:
                continue

            # 解析搜索结果
            self.logInfo.emit("解析淘宝搜索结果")
            self.parse(res)

        # 检查是否还有下一页
        ele_next_page = tab.ele("@@tag()=button@@text():下一页", timeout=0.5)
        if not ele_next_page or "disabled" in ele_next_page.attrs:
            return

        # 如果出现验证码，则等待
        while tab.ele("text:验证码", timeout=0.5):
            ...

        while tab.ele("#nocaptcha", timeout=0.5):
            ...

        # 往下滑动
        self.logInfo.emit("往下滑动")
        self.scroll_down(tab)

        ele_next_page = tab.ele("@@tag()=button@@text():下一页", timeout=0.5)

        # 检查是否还有下一页
        while ele_next_page and "disabled" not in ele_next_page.attrs:
            ele_next_page = tab.ele("@@tag()=button@@text():下一页", timeout=0.5)

            try:
                # 如果有下一页, 就点击下一页
                if ele_next_page:
                    ele_next_page.click()

                    # 检查是否还有下一页
                    if not ele_next_page or "disabled" not in ele_next_page.attrs:
                        # 滑动
                        self.logInfo.emit("往下滑动")
                        self.scroll_down(tab)

            except Exception as e:
                self.logInfo.emit(f"点击下一页出错: {e}")
                print(e)

        for package in tab.listen.steps(timeout=2):
            html_str = package.response.body

            if not html_str:
                continue

            # 解析 xhr 结果
            self.logInfo.emit("解析淘宝搜索结果")
            self.parse(html_str, filename=filename)
