# coding:utf-8
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from typing import Optional, override

import httpx
import pandas as pd
from loguru import logger
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import (BodyLabel, InfoBar, InfoBarPosition, ProgressBar,
                            PushButton, TextEdit)

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEdit
from view.interface.gallery_interface import GalleryInterface


class ImagesDownloader(QThread):
    logInfo = Signal(str)
    setProgress = Signal(int)
    setProgressInfo = Signal(int, int)

    def __init__(self, root_dir: Path):
        super(ImagesDownloader, self).__init__()

        self.root_dir = root_dir

        # 统计下载图片的数量
        self.download_count = 0

        self.total_rows = 0

        # httpx 客户端
        self.session = httpx.Client()

    @staticmethod
    def count_time(func):
        def wrapper(*args, **kwargs):
            start = datetime.now()
            func(*args, **kwargs)
            print(f"\033[38;5;208m\n\n\t耗时: {datetime.now() - start}\n\033[0m")

        return wrapper

    # 根据图片 url 下载图片
    def download_img(self, session: httpx.Client, img_url: str) -> bytes:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
        }

        res = session.get(img_url, headers=headers, follow_redirects=True)
        res.raise_for_status()  # 如果请求失败，抛出异常
        return res.content

    # 将图片写入文件
    def write_img(self, content: bytes, filename: Path):
        try:
            with open(filename, mode="wb") as f:
                f.write(content)
        except Exception as e:
            logger.error(f"图片保存失败: {filename} {e}")

    def process_excel_file(self, excel_path):
        # 创建保存图片的目录
        save_dir = self.root_dir / excel_path.stem
        save_dir.mkdir(parents=True, exist_ok=True)

        try:
            df = pd.read_excel(excel_path, usecols=["uuid", "药品图片"])

            tasks = []

            for row in df.itertuples(index=False):
                uuid = row[0]  # uuid
                img_url = str(row[1])  # 药品图片 url

                # 不是有效的 url
                if not str(img_url).startswith("http"):
                    continue

                filename = (
                    save_dir / f'{excel_path.stem}_{uuid}.{str(img_url).split(".")[-1]}'
                )

                if filename.exists():
                    msg = f"\t图片已存在: {excel_path.stem} {uuid}"
                    logger.info(msg)
                    self.logInfo.emit(msg)
                    continue

                tasks.append((img_url, filename))

            # 多线程下载图片
            with ThreadPoolExecutor() as t:
                futures = [
                    t.submit(self.download_and_save_img, img_url, filename)
                    for img_url, filename in tasks
                ]

                # 等待所有任务完成
                for future in futures:
                    future.result()

        except Exception as e:
            msg = f"处理失败: {excel_path} {e}"
            logger.error(msg)
            self.logInfo.emit(msg)

    def download_and_save_img(self, img_url, filename: Path):
        try:
            content = self.download_img(self.session, img_url)
            self.write_img(content, filename)

            self.download_count += 1

            # 更新进度条
            self.setProgress.emit(self.download_count / self.total_rows * 100)
            self.setProgressInfo.emit(self.download_count, self.total_rows)

        except Exception as e:
            filename.touch()
            msg = f"下载失败: {filename.stem} {img_url} {e}"
            logger.error(msg)
            self.logInfo.emit(msg)

    def count_rows(self) -> int:
        total_rows = 0

        for excelFile in self.root_dir.glob("*.xlsx"):
            if excelFile.stem.startswith("~"):
                continue

            df = pd.read_excel(excelFile, usecols=["uuid"])

            total_rows += int(df.shape[0])

        return total_rows

    @override
    def run(self):
        self.total_rows = self.count_rows()

        self.setProgressInfo.emit(0, self.total_rows)
        self.logInfo.emit(f"开始下载图片, 共 {self.total_rows} 张\n")

        for excel_file in self.root_dir.glob("*.xlsx"):
            if (
                excel_file.stem.startswith("~")
                or "对照" in excel_file.stem
                or "排查" in excel_file.stem
            ):
                continue

            self.process_excel_file(excel_file)

        self.session.close()

        msg = f"\n下载图片完成, 共下载 {self.download_count}  张图片. 应该有要下载 {self.total_rows} 张"
        self.logInfo.emit(msg)


class ImagesDownloadInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("图片下载", parent=parent)

        self.view = QWidget(self)

        # 状态提示
        self.stateTooltip = None

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout = QHBoxLayout()
        self.hBoxLayout_progress = QHBoxLayout()

        self.label_excel_path = BodyLabel(text="Excel 文件所在文件夹: ")

        # 显示 Excel 文件所在文件夹的文本框
        self.lineEdit_excel_path = DropableLineEdit()
        self.lineEdit_excel_path.setPlaceholderText(
            "请选择或者拖入 Excel 文件所在文件夹"
        )
        self.lineEdit_excel_path.textChanged.connect(
            lambda: cfg.set(cfg.downloadImg_img_path, self.lineEdit_excel_path.text())
        )

        # 选择路径的按钮
        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_excel_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 下载按钮
        self.btn_download = PushButton(text="下载")
        self.btn_download.clicked.connect(self.start_download)

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
        self.vBoxLayout.addWidget(self.btn_download)
        self.vBoxLayout.addWidget(self.textEdit_log)

        # 进度条布局
        self.vBoxLayout.addLayout(self.hBoxLayout_progress)

        self.__initWidget()

        # 从配置文件中读取路径
        self.lineEdit_excel_path.setText(cfg.downloadImg_img_path.value)

        """
         一些变量
        """
        self.worker: Optional[ImagesDownloader] = None

    def __initWidget(self):
        self.view.setObjectName("图片下载")
        self.setObjectName("ImagesDownloadInterface")

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
        self.lineEdit_excel_path.setEnabled(True)
        self.btn_download.setEnabled(True)

        if self.stateTooltip is not None:
            self.stateTooltip.hide()

        self.createSuccessInfoBar("完成", "数据爬取完成 ✅")

    def start_download(self):
        # 检查是否选择了文件夹
        self.textEdit_log.clear()

        excel_path = self.lineEdit_excel_path.text()
        if not excel_path:
            self.createErrorInfoBar("错误", "请选择 Excel 文件所在文件夹")
            return

        excel_path = Path(self.lineEdit_excel_path.text())

        self.lineEdit_excel_path.setEnabled(False)
        self.btn_download.setEnabled(False)

        # 创建下载器
        self.worker = ImagesDownloader(excel_path)

        self.worker.logInfo.connect(self.logInfo)
        self.worker.finished.connect(self.finished)
        self.worker.setProgress.connect(self.setProgress)
        self.worker.setProgressInfo.connect(self.setProgressInfo)

        self.worker.start()
