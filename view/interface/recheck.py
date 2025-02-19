# coding:utf-8
from datetime import datetime
from pathlib import Path
from typing import Optional, override

import pandas as pd
from DrissionPage import Chromium
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import (
    BodyLabel,
    InfoBar,
    InfoBarPosition,
    ProgressBar,
    PushButton,
    TextEdit,
)

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEditDir, DropableLineEditExcel
from view.interface.gallery_interface import GalleryInterface


class ReCheckWorker(QThread):
    logInfo = Signal(str)
    setProgress = Signal(int)
    setProgressInfo = Signal(int, int)

    def __init__(self, keywords_path: Path, output_dir: Path):
        super().__init__()

        self.keywords_path = keywords_path
        self.output_dir = output_dir

        self.bro = Chromium()

    def jd(self, store_url: str, medicine_name: str, store_name: str) -> bool:
        res: bool = False

        tab = self.bro.latest_tab

        # 等待用户登录
        self.logInfo.emit("请在浏览器中登录京东账号")
        tab.get("https://www.jd.com/")
        tab.ele("tag:a@class=nickname", timeout=70)

        # 进入店铺搜索药品
        new_tab = self.bro.new_tab(store_url)

        new_tab.ele("#key01", timeout=5).input(medicine_name)
        new_tab.ele(".button01", timeout=5).click()

        # 抱歉，没有找到
        if not new_tab.ele("text:抱歉，没有找到与", timeout=3):
            res = True

        self.bro.close_tabs(new_tab)

        return res

    def tb(self, store_url: str, medicine_name: str, store_name: str):
        res: bool = False

        tab = self.bro.latest_tab

        # 等待用户登录
        self.logInfo.emit("请在浏览器中登录淘宝账号")
        tab.get("https://www.taobao.com/")
        tab.ele("tag:a@class=site-nav-login-info-nick", timeout=70)

        # 进入店铺搜索药品
        tab.listen.start("h5api.m.taobao.com/h5/mtop.taobao.shop.simple.item.fetch")
        new_tab_id = self.bro.new_tab(store_url)
        tab.ele("text=搜索宝贝", timeout=10).input(medicine_name)
        tab.ele("搜本店", timeout=10).click()

        for package in tab.listen.steps():
            # 如果 data.data 里面有数据，说明找到了
            if package.response.body["data"]["data"]:
                res = True

        self.bro.close_tabs(new_tab_id)

        return res

    @override
    def run(self):
        start = datetime.now()

        df = pd.read_excel(self.keywords_path)

        # 只挑选平台为 京东 或者 淘宝 的数据
        df = df[df["平台"].isin(["京东", "淘宝"])]

        # 去重
        df.drop_duplicates(subset=["店铺主页"], keep="first", inplace=True)

        self.logInfo.emit(f"共有 {df.shape[0]} 间店铺需要复查")

        process_count = 0

        for i, row in df.iterrows():
            try:
                process_count += 1

                self.setProgress.emit(process_count / df.shape[0] * 100)
                self.setProgressInfo.emit(i + 1, df.shape[0])

                res: bool = True

                if "京东" == row["平台"]:
                    res = self.jd(row["店铺主页"], row["药品名"], row["药店名称"])
                elif "淘宝" == row["平台"]:
                    res = self.tb(row["店铺主页"], row["药品名"], row["药店名称"])

                # 如果 res 为 False, 说明对应平台下架了该药品, 则把该行移除
                if not res:
                    df = df.drop(i)

                # 保存结果
                df.to_excel(
                    self.output_dir / "复查结果.xlsx", index=False, engine="openpyxl"
                )
            except Exception as e:
                self.logInfo.emit(f"{row["店铺主页"]} 复查失败: {e}")
                continue

        # 保存结果
        df.to_excel(self.output_dir / "复查结果.xlsx", index=False, engine="openpyxl")

        self.logInfo.emit(f"\n耗时: {datetime.now() - start}")


class ReCheckInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("复查数据", parent=parent)

        self.view = QWidget(self)

        # 状态提示
        self.stateTooltip = None

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout = QHBoxLayout()
        self.hBoxLayout_output = QHBoxLayout()
        self.hBoxLayout_progress = QHBoxLayout()

        # Excel 文件所在文件夹的文本框
        self.label_onnx = BodyLabel(text="要复查药品的 Excel 文件: ")
        self.lineEdit_keywordPath = DropableLineEditExcel()
        self.lineEdit_keywordPath.setPlaceholderText(
            "请选择或者拖入复查药品的 Excel 文件"
        )
        self.lineEdit_keywordPath.textChanged.connect(
            lambda: cfg.set(cfg.recheck_excel_path, self.lineEdit_keywordPath.text())
        )

        self.label_output = BodyLabel(text="输出文件夹: ")
        # 输出文件夹的文本框
        self.lineEdit_output_path = DropableLineEditDir()
        self.lineEdit_output_path.setPlaceholderText("请选择或者拖入输出文件夹")
        self.lineEdit_output_path.textChanged.connect(
            lambda: cfg.set(cfg.recheck_output_path, self.lineEdit_output_path.text())
        )

        # 选择药品 Excel 文件的按钮
        self.btn_select_keyword_path = PushButton(text="···")
        self.btn_select_keyword_path.clicked.connect(
            lambda: self.lineEdit_keywordPath.setText(
                QFileDialog.getOpenFileName(
                    self, "选择文件", filter="Excel 文件(*.xlsx)"
                )[0]
            )
        )

        # 选择输出路径的按钮
        self.btn_select_output_path = PushButton(text="···")
        self.btn_select_output_path.clicked.connect(
            lambda: self.lineEdit_output_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 开始按钮
        self.btn_download = PushButton(text="开始")
        self.btn_download.clicked.connect(self.start)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        # 进度条
        self.progressBar = ProgressBar()

        # 进度提示标签
        self.label_progress = BodyLabel(text="0/0")

        self.hBoxLayout.addWidget(self.label_onnx)
        self.hBoxLayout.addWidget(self.lineEdit_keywordPath)
        self.hBoxLayout.addWidget(self.btn_select_keyword_path)

        self.hBoxLayout_output.addWidget(self.label_output)
        self.hBoxLayout_output.addWidget(self.lineEdit_output_path)
        self.hBoxLayout_output.addWidget(self.btn_select_output_path)

        self.hBoxLayout_progress.addWidget(self.progressBar)
        self.hBoxLayout_progress.addWidget(self.label_progress)

        self.vBoxLayout.addLayout(self.hBoxLayout)
        self.vBoxLayout.addLayout(self.hBoxLayout_output)

        self.vBoxLayout.addWidget(self.btn_download)
        self.vBoxLayout.addWidget(self.textEdit_log)

        # 进度条布局
        self.vBoxLayout.addLayout(self.hBoxLayout_progress)

        self.__initWidget()

        # 从配置文件中读取路径
        self.lineEdit_keywordPath.setText(cfg.recheck_excel_path.value)
        self.lineEdit_output_path.setText(cfg.recheck_output_path.value)

        self.worker: Optional[ReCheckWorker] = None

    def __initWidget(self):
        self.view.setObjectName("")
        self.setObjectName("ReCheckInterface")

        self.setWidget(self.view)
        self.setWidgetResizable(True)

    def createErrorInfoBar(self, title, content):
        """
        创建错误信息栏
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
        创建错误信息栏
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
        self.lineEdit_keywordPath.setEnabled(True)
        self.btn_select_keyword_path.setEnabled(True)

        self.lineEdit_output_path.setEnabled(True)
        self.btn_select_output_path.setEnabled(True)

        self.btn_download.setEnabled(True)

        if self.stateTooltip is not None:
            self.stateTooltip.hide()

        self.createSuccessInfoBar("完成", "复查完成 ✅")

    def start(self):
        self.textEdit_log.clear()

        # 检查是否选择了 药品 Excel 文件
        keywords_path = self.btn_select_keyword_path.text()
        if not keywords_path:
            self.createErrorInfoBar("错误", "请选择 onnx 模型文件")
            return

        # 检查是否选择了输出文件夹
        output_dir = self.lineEdit_output_path.text()
        if not output_dir:
            self.createErrorInfoBar("错误", "请选择输出文件夹")
            return

        keywords_path = Path(self.lineEdit_keywordPath.text())
        output_dir = Path(self.lineEdit_output_path.text())

        self.lineEdit_keywordPath.setEnabled(False)
        self.btn_select_keyword_path.setEnabled(False)

        self.lineEdit_output_path.setEnabled(False)
        self.btn_select_output_path.setEnabled(False)

        self.btn_download.setEnabled(False)

        self.worker = ReCheckWorker(keywords_path, output_dir)

        self.worker.logInfo.connect(self.logInfo)
        self.worker.finished.connect(self.finished)
        self.worker.setProgress.connect(self.setProgress)
        self.worker.setProgressInfo.connect(self.setProgressInfo)

        self.worker.start()
