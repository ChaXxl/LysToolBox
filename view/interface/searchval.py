# coding:utf-8
from pathlib import Path
from typing import Optional, Union, override

import pandas as pd
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtGui import QDropEvent
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import (
    BodyLabel,
    ComboBox,
    InfoBar,
    InfoBarPosition,
    PushButton,
    SwitchButton,
    TextEdit,
)

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEdit
from view.interface.gallery_interface import GalleryInterface


class SearchWorker(QThread):
    logInfo = Signal(str)

    def __init__(
        self, root_dir: Path, search_val: [str], search_column: str, recursive: bool
    ):
        super().__init__()

        self.root_dir = root_dir
        self.search_val = search_val
        self.search_column = search_column
        self.recursive = recursive

    @override
    def run(self):
        try:
            # 读取 root_dir 下所有 Excel 文件
            if self.recursive:
                excel_files = list(self.root_dir.rglob("*.xlsx"))
            else:
                excel_files = list(self.root_dir.glob("*.xlsx"))

            if not excel_files:
                self.logInfo.emit("没有找到 Excel 文件")
                return

            for file in excel_files:
                try:
                    df = pd.read_excel(file, engine="openpyxl")

                    if self.search_column not in df.columns:
                        continue

                    for idx, row in df.iterrows():
                        for val in self.search_val:
                            if val not in str(row[self.search_column]):
                                continue

                            # 打印在哪个文件、第几行找到
                            self.logInfo.emit(f"\n {file.stem}, {idx + 1} 行: {val}.")
                except Exception as e:
                    continue

        except Exception as e:
            self.logInfo.emit(f"失败: {e}")


class SearchValInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("查找值", parent=parent)

        self.view = QWidget(self)

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout = QHBoxLayout()
        self.hBoxLayout_search = QHBoxLayout()

        # 选择路径的按钮
        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_excel_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 输入查找值
        self.lineEdit_search_val = TextEdit()
        self.lineEdit_search_val.setPlaceholderText("请输入要查找的值, 一行一个")

        self.label_excel_path = BodyLabel(text="Excel 文件所在文件夹: ")

        # 显示 Excel 文件所在文件夹的文本框
        self.lineEdit_excel_path = DropableLineEdit()
        self.lineEdit_excel_path.setPlaceholderText("请选择 Excel 文件所在文件夹")
        self.lineEdit_excel_path.textChanged.connect(
            lambda: cfg.set(cfg.searchval_excel_path, self.lineEdit_excel_path.text())
        )

        self.label_search_column = BodyLabel(text="查找列: ")
        self.label_search_column.setMaximumWidth(60)

        # 选择在哪些列中查找: uuid	药店名称	店铺主页	资质名称
        self.comboBox = ComboBox()
        self.comboBox.addItem("uuid")
        self.comboBox.addItem("药店名称")
        self.comboBox.addItem("店铺主页")
        self.comboBox.addItem("资质名称")

        # 是否递归搜索
        self.label_recursive = BodyLabel(text="是否递归搜索: ")
        self.label_recursive.setMaximumWidth(90)
        self.switchButton = SwitchButton()
        self.switchButton.setOnText("")
        self.switchButton.setOffText("")
        self.switchButton.setMaximumWidth(80)

        # 查找按钮
        self.btn_search = PushButton(text="查找")
        self.btn_search.clicked.connect(self.search_val)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        self.hBoxLayout.addWidget(self.label_excel_path)
        self.hBoxLayout.addWidget(self.lineEdit_excel_path)
        self.hBoxLayout.addWidget(self.btn_select_path)

        self.hBoxLayout_search.addWidget(self.label_search_column)
        self.hBoxLayout_search.addWidget(self.comboBox)
        self.hBoxLayout_search.addWidget(self.label_recursive)
        self.hBoxLayout_search.addWidget(self.switchButton)
        self.hBoxLayout_search.addWidget(self.btn_search)

        self.vBoxLayout.addLayout(self.hBoxLayout)
        self.vBoxLayout.addWidget(self.lineEdit_search_val)
        self.vBoxLayout.addLayout(self.hBoxLayout_search)
        self.vBoxLayout.addWidget(self.textEdit_log)

        self.__initWidget()

        # 从配置文件中读取路径
        self.lineEdit_excel_path.setText(cfg.searchval_excel_path.value)

        """
         一些变量
        """
        self.worker: Optional[SearchWorker] = None

    def __initWidget(self):
        self.view.setObjectName("查找值")
        self.setObjectName("SearchValInterface")

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
    def finish(self):
        self.lineEdit_excel_path.setEnabled(True)
        self.btn_select_path.setEnabled(True)
        self.lineEdit_search_val.setEnabled(True)
        self.comboBox.setEnabled(True)
        self.btn_search.setEnabled(True)

        self.createSuccessInfoBar("完成", "查找完成")

    def search_val(self):
        self.textEdit_log.clear()

        # 检查是否选择了文件夹
        excel_path = self.lineEdit_excel_path.text()
        if not excel_path:
            self.createErrorInfoBar("错误", "请选择 Excel 文件所在文件夹")
            return

        # 检查是否输入了查找值
        search_val: str = self.lineEdit_search_val.toPlainText().strip()
        if not search_val:
            self.createErrorInfoBar("错误", "请输入要查找的值")
            return

        search_val: [str] = search_val.split("\n")

        # 看用户选择了哪个作为要查找的列
        search_column = self.comboBox.currentText()

        excel_path = Path(self.lineEdit_excel_path.text())

        # 获取是否递归搜索
        recursive = self.switchButton.isChecked()

        self.lineEdit_excel_path.setEnabled(False)
        self.btn_select_path.setEnabled(False)
        self.lineEdit_search_val.setEnabled(False)
        self.comboBox.setEnabled(False)
        self.btn_search.setEnabled(False)

        self.worker = SearchWorker(excel_path, search_val, search_column, recursive)
        self.worker.logInfo.connect(self.logInfo)
        self.worker.finished.connect(self.finish)
        self.worker.start()
