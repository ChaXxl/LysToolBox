# coding:utf-8
import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional, Union, override

import ddddocr
from PIL import Image

if not hasattr(Image, "ANTIALIAS"):
    setattr(Image, "ANTIALIAS", Image.LANCZOS)

import pandas as pd
from DrissionPage import Chromium
from DrissionPage.common import Keys
from lxml import etree
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import (BodyLabel, InfoBar, InfoBarPosition, ProgressBar,
                            PushButton, TextEdit)

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEditExcelDir
from view.interface.gallery_interface import GalleryInterface


class SearchJdCertWorker(QThread):
    logInfo = Signal(str)
    setProgress = Signal(int)
    setProgressInfo = Signal(int, int)

    def __init__(self, excel_path: Path):
        super().__init__()

        self.excel_path = excel_path

        # ocr 模型
        self.ocr = ddddocr.DdddOcr()

        # 初始化浏览器实例
        self.bro = Chromium()
        self.tab = self.bro.latest_tab

        # 预先编译正则表达式
        self.store_name_pattern = re.compile(r'document\.title="(.*?)"')

    def filter_data(self, excel_file: Path) -> list[str]:
        """
        从 Excel 文件中过滤店铺主页不为空、资质名称为空、平台为京东的记录
        """
        try:
            df = pd.read_excel(
                excel_file,
                usecols=["店铺主页", "资质名称", "平台"],
                dtype={"店铺主页": str, "资质名称": str, "平台": str},
            )

            mask = (
                df["店铺主页"].notna() & df["资质名称"].isna() & (df["平台"] == "京东")
            )

            return df.loc[mask, "店铺主页"].tolist()
        except Exception as e:
            self.logInfo.emit(f"读取 {excel_file} 文件失败: {e}")
            return []

    def ocr_classification(self, image: bytes) -> Union[str, dict]:
        """
        使用 ocr 模型识别验证码
        """
        try:
            return self.ocr.classification(image)
        except Exception as e:
            self.logInfo.emit(f"识别验证码失败: {e}")
            return ""

    def parse(self, res: str) -> None:
        """
        解析营业执照页面，提取店铺名和公司名，并保存到 Excel
        """
        try:
            html = etree.HTML(res)
            companyName = html.xpath('//li[@class="noBorder"][2]/span/text()')[0]

            storeName = self.store_name_pattern.search(res)
            storeName = storeName.group(1).strip() if storeName else ""

            if not storeName or "根据国家相关政策" in companyName:
                self.logInfo.emit(f"店铺名称或公司名称为空: {storeName} {companyName}")
                return

            self.write_to_excel(storeName, companyName)
        except Exception as e:
            self.logInfo.emit(f"解析营业执照页面失败: {e}")

    def write_to_excel(self, storeName: str, companyName: str) -> None:
        """
        将店铺和公司名写入 Excel 文件
        """
        if self.excel_path.is_file():
            try:
                df = pd.read_excel(self.excel_path)
                updates = 0

                # 检查药店名称是否是 storeName, 如果是并且资质名称列为空, 则更新
                for index, row in df.iterrows():
                    if row["药店名称"] != storeName or pd.notna(row["资质名称"]):
                        continue

                    updates += 1
                    df.loc[index, "资质名称"] = companyName
                    df.to_excel(self.excel_path, index=False)

                if updates:
                    self.logInfo.emit(
                        f"{self.excel_path.stem} 更新了 {updates} 行, {storeName} 的资质名称为 {companyName}"
                    )
            except Exception as e:
                self.logInfo.emit(f"写入 {self.excel_path.stem} 文件失败: {e}")
        elif self.excel_path.is_dir():
            for file in self.excel_path.glob("*.xlsx"):
                if any(keyword in file.stem for keyword in ["~", "对照", "排查"]):
                    continue

                try:
                    df = pd.read_excel(file)
                    updates = 0

                    # 检查药店名称是否是 storeName, 如果是并且资质名称列为空, 则更新
                    for index, row in df.iterrows():
                        if row["药店名称"] != storeName or pd.notna(row["资质名称"]):
                            continue

                        updates += 1
                        df.loc[index, "资质名称"] = companyName
                        df.to_excel(file, index=False)

                    if updates:
                        self.logInfo.emit(
                            f"{file.stem} 更新了 {updates} 行, {storeName} 的资质名称为 {companyName}"
                        )
                except Exception as e:
                    self.logInfo.emit(f"写入 {file.stem} 文件失败: {e}")

    def process_url(self, url: str) -> None:
        """
        处理单个 URL，完成验证码输入及数据抓取
        """
        try:
            self.tab.listen.start("mall.jd.com/showLicence", method="POST")
            self.tab.get(url)

            # 输入验证码
            verifyCode_input = self.tab("#verifyCode", timeout=10)
            verifyCodeImg = self.tab("#verifyCodeImg", timeout=1)

            img = verifyCodeImg.src()
            verifyCode = self.ocr_classification(img)
            verifyCode_input.input(verifyCode).input(Keys.ENTER)

            # 验证码错误, 重新识别
            while self.tab("#verifyCode_error", timeout=3):
                img = verifyCodeImg.src()
                verifyCode = self.ocr_classification(img)
                verifyCode_input.input(verifyCode).input(Keys.ENTER)

            # 获取数据包并解析
            res = self.tab.listen.wait(timeout=2)
            if res:
                self.parse(res.response.body)
        except Exception as e:
            self.logInfo.emit(f"处理 {url} 出错: {e}")

    @override
    def run(self):
        urls = set()

        if self.excel_path.is_file():
            urls.update(
                url.replace("index", "showLicence").replace("?from=pc", "")
                for url in self.filter_data(self.excel_path)
            )
        elif self.excel_path.is_dir():
            # 从 Excel 文件获取店铺主页的 url
            for file in self.excel_path.glob("*.xlsx"):
                if any(keyword in file.stem for keyword in ["~", "对照", "排查"]):
                    continue

                urls.update(
                    url.replace("index", "showLicence").replace("?from=pc", "")
                    for url in self.filter_data(file)
                )

        self.setProgressInfo.emit(0, len(urls))

        for idx, url in enumerate(urls):
            # 更新进度条
            self.setProgress.emit(idx + 1 / len(urls) * 100)
            self.setProgressInfo.emit(idx + 1, len(urls))

            self.process_url(url)


class SearchJdCertInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__(title="查找京东的资质名称", parent=parent)

        self.view = QWidget(self)

        # 状态提示
        self.stateTooltip = None

        # 界面的垂直布局
        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout = QHBoxLayout()
        self.hBoxLayout_progress = QHBoxLayout()

        self.label_excel_path = BodyLabel(text="Excel 文件或者 Excel 所在文件夹: ")

        # 显示 Excel 文件所在文件夹的文本框
        self.lineEdit_excel_path = DropableLineEditExcelDir()
        self.lineEdit_excel_path.textChanged.connect(
            lambda: cfg.set(
                cfg.searchJdCert_excel_path, self.lineEdit_excel_path.text()
            )
        )
        self.lineEdit_excel_path.setPlaceholderText("请选择或者拖入 Excel 文件或文件夹")

        # 选择路径的按钮
        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_excel_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件")
            )
        )

        # 开始搜索按钮
        self.btn_search = PushButton(text="开始查找")
        self.btn_search.clicked.connect(self.search)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        # 进度条
        self.progressBar = ProgressBar()

        # 进度提示标签
        self.label_progress = BodyLabel(text="0/0")

        self.hBoxLayout.addWidget(self.label_excel_path)
        self.hBoxLayout.addWidget(self.lineEdit_excel_path)
        self.hBoxLayout.addWidget(self.btn_select_path)

        self.hBoxLayout_progress.addWidget(self.progressBar)
        self.hBoxLayout_progress.addWidget(self.label_progress)

        self.vBoxLayout.addLayout(self.hBoxLayout)

        self.vBoxLayout.addWidget(self.btn_search)
        self.vBoxLayout.addWidget(self.textEdit_log)

        # 进度条布局
        self.vBoxLayout.addLayout(self.hBoxLayout_progress)

        self.__initWidget()

        self.worker: Optional[SearchJdCertWorker] = None

        self.lineEdit_excel_path.setText(cfg.searchJdCert_excel_path.value)

    def __initWidget(self):
        self.view.setObjectName("")
        self.setObjectName("searchJdCertInterface")

        self.setWidget(self.view)
        self.setWidgetResizable(True)

    def __initLayout(self):
        self.hBoxLayout.setSpacing(8)

    def createErrorInfoBar(self, title, content):
        """
        创建错误信息栏
        :param title:
        :param content:
        :return:
        """
        InfoBar.error(
            title=title,
            content=content,
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.BOTTOM_RIGHT,
            duration=2000,
            parent=self,
        )

    def createSuccessInfoBar(self, title, content):
        """
        创建成功信息栏
        :param title:
        :param content:
        :return:
        """
        InfoBar.success(
            title=title,
            content=content,
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=2000,
            parent=self,
        )

    @Slot(str)
    def logInfo(self, info):
        """
        打印日志
        """
        self.textEdit_log.append(info)

    @Slot(int)
    def setProgress(self, value):
        """
        设置进度条进度
        """
        self.progressBar.setValue(value)

    @Slot(int, int)
    def setProgressInfo(self, value, total):
        """
        设置进度条进度
        """
        self.label_progress.setText(f"{value}/{total}")

    @Slot()
    def finished(self):
        """处理完成"""
        self.lineEdit_excel_path.setEnabled(True)
        self.btn_select_path.setEnabled(True)
        self.btn_search.setEnabled(True)

        self.createSuccessInfoBar("成功", "处理完成")

    def search(self):
        self.textEdit_log.clear()

        excel_path = Path(self.lineEdit_excel_path.text())

        if not excel_path.exists():
            self.createErrorInfoBar("错误", "Excel 文件不存在")
            return

        self.lineEdit_excel_path.setEnabled(False)
        self.btn_select_path.setEnabled(False)
        self.btn_search.setEnabled(False)

        self.textEdit_log.clear()

        self.worker = SearchJdCertWorker(excel_path)
        self.worker.logInfo.connect(self.logInfo)
        self.worker.setProgress.connect(self.setProgress)
        self.worker.setProgressInfo.connect(self.setProgressInfo)
        self.worker.finished.connect(self.finished)
        self.worker.start()
