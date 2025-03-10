# coding:utf-8
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Optional, override

import filetype
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import (BodyLabel, InfoBar, InfoBarPosition, ProgressBar,
                            PushButton, TextEdit)

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEditDir
from view.interface.gallery_interface import GalleryInterface


class FixWorker(QThread):
    logInfo = Signal(str)
    setProgress = Signal(int)
    setProgressInfo = Signal(int, int)

    # 定义支持的文件格式集合
    SUPPORTED_FORMATS = {".jpeg", ".jpg", ".webp", ".png", ".avif", ".gif"}

    def __init__(self, root_dir: Path):
        super().__init__()
        self.root_dir = root_dir

    @override
    def run(self):
        # 记录开始时间
        start_time = datetime.now()

        files = [
            f for f in self.root_dir.rglob("*") if f.suffix in self.SUPPORTED_FORMATS
        ]

        for idx, file_path in enumerate(files):
            try:
                self.setProgress.emit((idx + 1) * 100 // len(files))
                self.setProgressInfo.emit(idx + 1, len(files))

                # 检测文件格式
                mime_type = filetype.guess_mime(str(file_path))
                if mime_type is None:
                    self.logInfo.emit(f"文件格式检测失败: {file_path}")
                    continue

                target_format = mime_type.split("/")[-1]

                if mime_type == "image/jpeg" and file_path.suffix == ".jpg":
                    continue

                if file_path.suffix == f".{target_format}":
                    continue

                # 构造目标文件路径
                target_path = file_path.with_suffix(f".{target_format}")

                # 重命名文件
                file_path.rename(target_path)

            except Exception as e:
                self.logInfo.emit(f"处理失败: {file_path} - {str(e)}")
                continue

        # 记录处理时间
        self.logInfo.emit(f"\n耗时: {datetime.now() - start_time}")


class FixImageSuffixInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__(title="修正图片后缀名", parent=parent)

        self.view = QWidget(self)

        # 状态提示
        self.stateTooltip = None

        # 界面的垂直布局
        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout = QHBoxLayout()
        self.hBoxLayout_progress = QHBoxLayout()

        self.label_img_path = BodyLabel(text="图片所在文件夹: ")

        # 显示 Excel 文件所在文件夹的文本框
        self.lineEdit_img_path = DropableLineEditDir()
        self.lineEdit_img_path.textChanged.connect(
            lambda: cfg.set(cfg.fiximgsuffix_excel_path, self.lineEdit_img_path.text())
        )
        self.lineEdit_img_path.setPlaceholderText("请选择或者拖入 Excel 文件所在文件夹")

        # 选择路径的按钮
        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_img_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 刷新按钮
        self.btn_refresh = PushButton(text="开始")
        self.btn_refresh.clicked.connect(self.start)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        # 进度条
        self.progressBar = ProgressBar()

        # 进度提示标签
        self.label_progress = BodyLabel(text="0/0")

        self.hBoxLayout.addWidget(self.label_img_path)
        self.hBoxLayout.addWidget(self.lineEdit_img_path)
        self.hBoxLayout.addWidget(self.btn_select_path)

        self.hBoxLayout_progress.addWidget(self.progressBar)
        self.hBoxLayout_progress.addWidget(self.label_progress)

        self.vBoxLayout.addLayout(self.hBoxLayout)

        self.vBoxLayout.addWidget(self.btn_refresh)
        self.vBoxLayout.addWidget(self.textEdit_log)

        self.vBoxLayout.addLayout(self.hBoxLayout_progress)

        self.__initWidget()

        self.worker: Optional[FixWorker] = None

        self.lineEdit_img_path.setText(cfg.fiximgsuffix_excel_path.value)

    def __initWidget(self):
        self.view.setObjectName("")
        self.setObjectName("FixImageSuffixInterface")

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
        self.lineEdit_img_path.setEnabled(True)
        self.btn_select_path.setEnabled(True)
        self.btn_refresh.setEnabled(True)

        self.createSuccessInfoBar("成功", "处理完成")

    def start(self):
        """格式化 Excel 文件"""
        self.textEdit_log.clear()

        excel_dir = Path(self.lineEdit_img_path.text())

        if excel_dir.is_file():
            return

        if not excel_dir.exists():
            self.createErrorInfoBar("错误", "文件夹不存在")
            return

        self.lineEdit_img_path.setEnabled(False)
        self.btn_select_path.setEnabled(False)
        self.btn_refresh.setEnabled(False)

        self.textEdit_log.clear()

        self.worker = FixWorker(excel_dir)

        self.worker.logInfo.connect(self.logInfo)
        self.worker.finished.connect(self.finished)
        self.worker.setProgress.connect(self.setProgress)
        self.worker.setProgressInfo.connect(self.setProgressInfo)

        self.worker.start()
