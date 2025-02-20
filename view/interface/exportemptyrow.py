# coding:utf-8
from pathlib import Path
from typing import Optional, override

import pandas as pd
from PySide6.QtCore import QThread, Slot, Qt, Signal
from PySide6.QtWidgets import QHBoxLayout, QVBoxLayout, QWidget, QFileDialog
from qfluentwidgets import (
    BodyLabel,
    PushButton,
    TextEdit,
    InfoBar,
    InfoBarPosition,
)
from view.components.dropable_lineEdit import (
    DropableLineEditDir,
    DropableLineEditExcelDir,
)
from view.interface.gallery_interface import GalleryInterface
from common.config import cfg


class ExportEmptyRowInterfaceWorker(QThread):
    logInfo = Signal(str)

    def __init__(self, excel_path: Path, output_path: Path):
        super().__init__()
        self.excel_path = excel_path
        self.output_path = output_path

    @override
    def run(self):
        df_list = []

        if self.excel_path.is_file():
            self.logInfo.emit(f"正在导出 {self.excel_path.stem} ...")

            df_list.append(
                pd.read_excel(
                    self.excel_path,
                    usecols=["药店名称", "店铺主页", "资质名称", "平台"],
                )
            )
        else:
            for excel_file in self.excel_path.glob("*.xlsx"):
                self.logInfo.emit(f"{excel_file.stem} ...")

                if any(keyword in excel_file.stem for keyword in ["~", "对照", "排查"]):
                    continue

                df_list.append(
                    pd.read_excel(
                        excel_file, usecols=["药店名称", "店铺主页", "资质名称", "平台"]
                    )
                )

        if df_list:
            df = pd.concat(df_list, ignore_index=True)  # 合并所有 DataFrame
            df = df[df["资质名称"].isna()].drop_duplicates()  # 筛选 + 去重
            df.to_excel(f"{self.output_path}/导出结果.xlsx", index=False)  # 保存结果


class ExportEmptyRowInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("导出资质名称为空的行数", parent=parent)

        self.view = QWidget(self)

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout_input = QHBoxLayout()
        self.hBoxLayout_output = QHBoxLayout()

        # 选择 Excel 文件所在文件夹
        self.label_excel_path = BodyLabel(text="Excel 文件所在文件夹或者文件: ")
        self.lineEdit_excel_path = DropableLineEditExcelDir()
        self.lineEdit_excel_path.setPlaceholderText(
            "请选择或者拖入上一次的 Excel 文件所在文件夹或者文件"
        )

        self.lineEdit_excel_path.textChanged.connect(
            lambda: cfg.set(
                cfg.exportEmptyRow_excel_path, self.lineEdit_excel_path.text()
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
                cfg.exportEmptyRow_output_path, self.lineEdit_output_path.text()
            )
        )

        self.btn_select_output_path = PushButton(text="···")
        self.btn_select_output_path.clicked.connect(
            lambda: self.lineEdit_output_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 按钮 用于开始导出
        self.btn_export = PushButton(text="导出")
        self.btn_export.clicked.connect(self.export)

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
        self.vBoxLayout.addWidget(self.btn_export)
        self.vBoxLayout.addWidget(self.textEdit_log)

        self.__initWidget()

        self.worker: Optional[ExportEmptyRowInterfaceWorker] = None

        self.lineEdit_excel_path.setText(cfg.exportEmptyRow_excel_path.value)
        self.lineEdit_output_path.setText(cfg.exportEmptyRow_output_path.value)

    def __initWidget(self):
        self.view.setObjectName("ExportEmptyRowInterface")
        self.setObjectName("ExportEmptyRowInterface")

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

    @Slot()
    def finished(self):
        self.label_excel_path.setEnabled(True)
        self.lineEdit_excel_path.setEnabled(True)
        self.lineEdit_output_path.setEnabled(True)

        self.btn_select_path.setEnabled(True)
        self.btn_export.setEnabled(True)

        self.createSuccessInfoBar("完成", "导出成功 ✅")

    def export(self):
        self.textEdit_log.clear()

        # 检查是否选择了文件夹
        excel_path = self.lineEdit_excel_path.text()
        output_path = self.lineEdit_output_path.text()

        if not excel_path:
            self.createErrorInfoBar("错误", "请选择Excel文件或者所在文件夹")
            return

        if not output_path:
            self.createErrorInfoBar("错误", "请选择输出文件夹")
            return

        # 创建工作线程
        self.worker = ExportEmptyRowInterfaceWorker(Path(excel_path), Path(output_path))
        self.worker.logInfo.connect(self.logInfo)
        self.worker.finished.connect(self.finished)
        self.worker.start()
