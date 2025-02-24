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


class IncrementalDatasWorker(QThread):
    """
    统计新增加的数据的工作线程
    """

    logInfo = Signal(str)

    def __init__(self, excel_path1: Path, excel_path2: Path, output_path: Path):
        super().__init__()
        self.excel_path1 = excel_path1
        self.excel_path2 = excel_path2
        self.output_path = output_path

    @override
    def run(self):
        try:
            # 读取上一次的 Excel 文件
            if self.excel_path1.is_file():
                df1 = pd.read_excel(self.excel_path1)
            else:
                dataframes: list[pd.DataFrame] = []
                for file in self.excel_path1.glob("*.xlsx"):
                    df = pd.read_excel(file)
                    if not df.empty and not df.isnull().all().all():
                        dataframes.append(df)

                if not dataframes:
                    self.logInfo.emit("没有找到上一次需要合并的文件")
                    return

                df1 = pd.concat(dataframes)

            # 读取这次的 Excel 文件
            if self.excel_path2.is_file():
                df2 = pd.read_excel(self.excel_path2)
            else:
                dataframes: list[pd.DataFrame] = []
                for file in self.excel_path2.glob("*.xlsx"):
                    df = pd.read_excel(file)
                    if not df.empty and not df.isnull().all().all():
                        dataframes.append(df)

                if not dataframes:
                    self.logInfo.emit("没有找到这次的 Excel 文件")
                    return

                df2 = pd.concat(dataframes)

            # 根据这几列(药店名称、店铺主页、资质名称、药品名 平台)来筛选出 df2 比 df1 多了哪些行
            key_columns = ["药店名称", "店铺主页", "资质名称", "药品名", "平台"]

            # 将多列的组合变成元组，再进行筛选
            df_incremental = df2[
                ~df2[key_columns]
                .apply(tuple, axis=1)
                .isin(df1[key_columns].apply(tuple, axis=1))
            ]

            # 保存新增加的数据
            df_incremental.to_excel(self.output_path / "新增加的数据.xlsx", index=False)
        except Exception as e:
            self.logInfo.emit(f"统计新增加的数据失败 ❌: {e}")


class IncrementalDatasInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("统计新增加的数据", parent=parent)

        self.view = QWidget(self)

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout_input1 = QHBoxLayout()
        self.hBoxLayout_input2 = QHBoxLayout()
        self.hBoxLayout_output = QHBoxLayout()

        # 选择上一次的 Excel 文件所在文件夹
        self.label_excel_path1 = BodyLabel(
            text="上一次的 Excel 文件所在文件夹或者文件: "
        )
        self.lineEdit_excel_path1 = DropableLineEditExcelDir()
        self.lineEdit_excel_path1.setPlaceholderText(
            "请选择或者拖入上一次的 Excel 文件所在文件夹或者文件"
        )

        self.lineEdit_excel_path1.textChanged.connect(
            lambda: cfg.set(
                cfg.incrementalDatas_excel_path1, self.lineEdit_excel_path1.text()
            )
        )

        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_excel_path1.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 选择这次的 Excel 文件所在文件夹
        self.label_excel_path2 = BodyLabel(text="这次的 Excel 文件所在文件夹或者文件: ")
        self.lineEdit_excel_path2 = DropableLineEditExcelDir()
        self.lineEdit_excel_path2.setPlaceholderText(
            "请选择或者拖入这次的 Excel 文件所在文件夹或者文件"
        )

        self.lineEdit_excel_path2.textChanged.connect(
            lambda: cfg.set(
                cfg.incrementalDatas_excel_path2, self.lineEdit_excel_path2.text()
            )
        )

        self.btn_select_path2 = PushButton(text="···")
        self.btn_select_path2.clicked.connect(
            lambda: self.lineEdit_excel_path2.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 输出文件夹
        self.label_output_path = BodyLabel(text="输出文件夹: ")
        self.lineEdit_output_path = DropableLineEditDir()
        self.lineEdit_output_path.setPlaceholderText("请选择或者拖入输出文件夹")

        self.lineEdit_output_path.textChanged.connect(
            lambda: cfg.set(
                cfg.incrementalDatas_output_path, self.lineEdit_output_path.text()
            )
        )

        self.btn_select_output_path = PushButton(text="···")
        self.btn_select_output_path.clicked.connect(
            lambda: self.lineEdit_output_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 按钮 用于开始合并
        self.btn_incremental = PushButton(text="统计")
        self.btn_incremental.clicked.connect(self.incremental)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        self.hBoxLayout_input1.addWidget(self.label_excel_path1)
        self.hBoxLayout_input1.addWidget(self.lineEdit_excel_path1)
        self.hBoxLayout_input1.addWidget(self.btn_select_path)

        self.hBoxLayout_input2.addWidget(self.label_excel_path2)
        self.hBoxLayout_input2.addWidget(self.lineEdit_excel_path2)
        self.hBoxLayout_input2.addWidget(self.btn_select_path2)

        self.hBoxLayout_output.addWidget(self.label_output_path)
        self.hBoxLayout_output.addWidget(self.lineEdit_output_path)
        self.hBoxLayout_output.addWidget(self.btn_select_output_path)

        self.vBoxLayout.addLayout(self.hBoxLayout_input1)
        self.vBoxLayout.addLayout(self.hBoxLayout_input2)
        self.vBoxLayout.addLayout(self.hBoxLayout_output)
        self.vBoxLayout.addWidget(self.btn_incremental)
        self.vBoxLayout.addWidget(self.textEdit_log)

        self.__initWidget()

        self.worker: Optional[IncrementalDatasWorker] = None

        self.lineEdit_excel_path1.setText(cfg.incrementalDatas_excel_path1.value)
        self.lineEdit_excel_path2.setText(cfg.incrementalDatas_excel_path2.value)
        self.lineEdit_output_path.setText(cfg.incrementalDatas_output_path.value)

    def __initWidget(self):
        self.view.setObjectName("IncrementalDatasInterface")
        self.setObjectName("IncrementalDatasInterface")

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
        self.lineEdit_excel_path1.setEnabled(True)
        self.lineEdit_excel_path2.setEnabled(True)
        self.lineEdit_output_path.setEnabled(True)

        self.btn_select_path.setEnabled(True)
        self.btn_select_path2.setEnabled(True)
        self.btn_incremental.setEnabled(True)

        self.createSuccessInfoBar("完成", "统计新增加的数据完成")

    def incremental(self):
        self.textEdit_log.clear()

        # 检查是否选择了文件夹
        excel_path1 = self.lineEdit_excel_path1.text()
        excel_path2 = self.lineEdit_excel_path2.text()
        output_path = self.lineEdit_output_path.text()

        if not excel_path1:
            self.createErrorInfoBar("错误", "请选择Excel文件所在文件夹")
            return

        if not excel_path2:
            self.createErrorInfoBar("错误", "请选择这次的Excel文件所在文件夹")
            return

        if not output_path:
            self.createErrorInfoBar("错误", "请选择输出文件夹")
            return

        # 创建工作线程
        self.worker = IncrementalDatasWorker(
            Path(excel_path1), Path(excel_path2), Path(output_path)
        )
        self.worker.logInfo.connect(self.logInfo)
        self.worker.finished.connect(self.finished)
        self.worker.start()
