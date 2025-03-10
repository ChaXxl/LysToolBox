# coding:utf-8
from pathlib import Path
from typing import Optional, override

import pandas as pd
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import BodyLabel, InfoBar, InfoBarPosition, PushButton, TextEdit

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEditDir
from view.interface.gallery_interface import GalleryInterface


class MergedExcelFilesWorker(QThread):
    """
    合并 Excel 文件的工作线程
    """

    logInfo = Signal(str)

    def __init__(self, excel_path: Path, output_path: Path):
        super().__init__()
        self.excel_path = excel_path
        self.output_path = output_path

    @override
    def run(self):
        # 获取 Excel 文件列表
        excel_files = list(Path(self.excel_path).glob("*.xlsx"))
        self.logInfo.emit(f"获取到 {len(excel_files)} 个 Excel 文件")

        # 合并 Excel 文件
        merged_excel_file = self.output_path / "合并.xlsx"

        # 创建一个空的列表, 用来存储非空的数据
        dataframes = []

        try:
            # 遍历所有需要合并的文件
            for f in excel_files:
                df = pd.read_excel(f)
                if not df.empty and not df.isnull().all().all():
                    dataframes.append(df)

            if not dataframes:
                self.logInfo.emit("没有找到需要合并的文件")
                return

            # 合并所有非空的数据
            merged_df = pd.concat(dataframes)

            # 去除重复的行
            merged_df = merged_df.drop_duplicates()

            # 保存合并后的数据
            merged_df.to_excel(merged_excel_file, index=False)
            self.logInfo.emit(f"保存合并后的数据到: {merged_excel_file}")

        except Exception as e:
            self.logInfo.emit(f"合并 Excel 文件失败: {e}")


class MergedExcelFilesInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("合并 Excel 文件", parent=parent)

        self.view = QWidget(self)

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout_input = QHBoxLayout()
        self.hBoxLayout_output = QHBoxLayout()

        # Excel文件所在文件夹
        self.label_excel_path = BodyLabel(text="Excel文件所在文件夹: ")
        self.lineEdit_excel_path = DropableLineEditDir()
        self.lineEdit_excel_path.setPlaceholderText("请选择或者拖入Excel文件所在文件夹")
        self.lineEdit_excel_path.textChanged.connect(
            lambda: cfg.set(
                cfg.mergedExcelFiles_excel_path, self.lineEdit_excel_path.text()
            )
        )

        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_excel_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 输出文件夹
        self.label_output_path = BodyLabel(text="输出文件夹: ")
        self.lineEdit_output_path = DropableLineEditDir()
        self.lineEdit_output_path.setPlaceholderText("请选择或者拖入输出文件夹")
        self.lineEdit_output_path.textChanged.connect(
            lambda: cfg.set(
                cfg.mergedExcelFiles_output_path, self.lineEdit_output_path.text()
            )
        )

        self.btn_select_output_path = PushButton(text="···")
        self.btn_select_output_path.clicked.connect(
            lambda: self.lineEdit_output_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 按钮 用于开始合并
        self.btn_merged = PushButton(text="合并")
        self.btn_merged.clicked.connect(self.merge)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        self.hBoxLayout_input.addWidget(self.label_excel_path)
        self.hBoxLayout_input.addWidget(self.lineEdit_excel_path)
        self.hBoxLayout_input.addWidget(self.btn_select_path)

        self.hBoxLayout_output.addWidget(self.label_output_path)
        self.hBoxLayout_output.addWidget(self.lineEdit_output_path)
        self.hBoxLayout_output.addWidget(self.btn_select_output_path)

        self.vBoxLayout.addLayout(self.hBoxLayout_input)
        self.vBoxLayout.addLayout(self.hBoxLayout_output)
        self.vBoxLayout.addWidget(self.btn_merged)
        self.vBoxLayout.addWidget(self.textEdit_log)

        self.__initWidget()

        self.worker: Optional[MergedExcelFilesWorker] = None

        self.lineEdit_excel_path.setText(cfg.mergedExcelFiles_excel_path.value)
        self.lineEdit_output_path.setText(cfg.mergedExcelFiles_output_path.value)

    def __initWidget(self):
        self.view.setObjectName("MergedExcelFilesInterface")
        self.setObjectName("MergedExcelFilesInterface")

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
        创建成功信息栏
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

    @Slot()
    def finished(self):
        self.lineEdit_excel_path.setEnabled(True)
        self.lineEdit_output_path.setEnabled(True)

        self.btn_select_path.setEnabled(True)
        self.btn_merged.setEnabled(True)

        self.createSuccessInfoBar("完成", "合并完成")

    def merge(self):
        self.textEdit_log.clear()

        # 检查是否选择了文件夹
        excel_path = self.lineEdit_excel_path.text()
        output_path = self.lineEdit_output_path.text()

        if not excel_path:
            self.createErrorInfoBar("错误", "请选择Excel文件所在文件夹")
            return

        if not output_path:
            self.createErrorInfoBar("错误", "请选择输出文件夹")
            return

        # 创建工作线程
        self.worker = MergedExcelFilesWorker(Path(excel_path), Path(output_path))
        self.worker.finished.connect(self.finished)
        self.worker.logInfo.connect(self.logInfo)
        self.worker.start()
