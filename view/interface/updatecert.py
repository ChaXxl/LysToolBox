# coding:utf-8
from pathlib import Path
from typing import override

import pandas as pd
import psycopg as pg
from psycopg import sql
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QVBoxLayout, QWidget
from qfluentwidgets import (
    BodyLabel,
    InfoBar,
    InfoBarPosition,
    LineEdit,
    PasswordLineEdit,
    PushButton,
    TextEdit,
)

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEdit
from view.interface.gallery_interface import GalleryInterface


class UpdateWorker(QThread):
    logInfo = Signal(str)

    def __init__(self, db_config: dict, excel_path: Path):
        super().__init__()

        self.db_config = db_config
        self.excel_path = excel_path

    @override
    def run(self):
        try:
            conn = pg.connect(**self.db_config)
            cursor = conn.cursor()

            df = pd.read_excel(self.excel_path, usecols=["药店名称", "资质名称"])
            df = df[(df["药店名称"] != "") & df["资质名称"].notna()]

            if df.empty:
                self.logInfo.emit("没有数据需要更新")
                return

            update_query = (
                "UPDATE store_info SET qualification_name = %s WHERE store_name = %s"
            )
            cursor.executemany(update_query, df.values.tolist())
            conn.commit()

            rowcount = cursor.rowcount

            self.logInfo.emit(f"\n更新了 {rowcount} 条数据")
        except Exception as e:
            self.logInfo.emit(f"失败: {e}")
        finally:
            conn.close()


class UpdateCertInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__(title="更新数据库的资质名称", parent=parent)

        self.view = QWidget(self)

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout_db = QHBoxLayout()
        self.hBoxLayout_file = QHBoxLayout()

        # host
        self.lineEdit_host = LineEdit()
        self.lineEdit_host.setPlaceholderText("IP 地址")
        self.lineEdit_host.setText(cfg.updateCert_host.value)
        self.lineEdit_host.textChanged.connect(
            lambda: cfg.set(cfg.updateCert_host, self.lineEdit_host.text())
        )

        # 数据库名称
        self.lineEdit_dbname = LineEdit()
        self.lineEdit_dbname.setPlaceholderText("数据库名称")
        self.lineEdit_dbname.setText(cfg.updateCert_dbname.value)
        self.lineEdit_dbname.textChanged.connect(
            lambda: cfg.set(cfg.updateCert_dbname, self.lineEdit_dbname.text())
        )

        # 用户名
        self.lineEdit_user = LineEdit()
        self.lineEdit_user.setPlaceholderText("用户名")
        self.lineEdit_user.setText(cfg.updateCert_user.value)
        self.lineEdit_user.textChanged.connect(
            lambda: cfg.set(cfg.updateCert_user, self.lineEdit_user.text())
        )

        # 密码
        self.lineEdit_password = PasswordLineEdit()
        self.lineEdit_password.setPlaceholderText("密码")
        self.lineEdit_password.setText(cfg.updateCert_password.value)
        self.lineEdit_password.textChanged.connect(
            lambda: cfg.set(cfg.updateCert_password, self.lineEdit_password.text())
        )

        # 测试连接按钮
        self.btn_test_connection = PushButton(text="测试连接")
        self.btn_test_connection.clicked.connect(self.testConnection)

        self.label_excel_path = BodyLabel(text="Excel 文件所在文件夹: ")

        self.label_excel_path = BodyLabel(text="Excel 文件所在文件夹: ")

        # 显示 Excel 文件所在文件夹的文本框
        self.lineEdit_excel_path = DropableLineEdit()
        self.lineEdit_excel_path.setPlaceholderText("请选择或者拖入 Excel 文件")
        self.lineEdit_excel_path.textChanged.connect(
            lambda: cfg.set(cfg.updateCert_excel_path, self.lineEdit_excel_path.text())
        )

        # 选择路径的按钮
        self.btn_select_path = PushButton("···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_excel_path.setText(
                QFileDialog.getOpenFileName(self, "选择文件夹")[0]
            )
        )

        # 开始按钮
        self.btn_start = PushButton(text="开始")
        self.btn_start.clicked.connect(self.update)

        self.hBoxLayout_db.addWidget(self.lineEdit_host)
        self.hBoxLayout_db.addWidget(self.lineEdit_dbname)
        self.hBoxLayout_db.addWidget(self.lineEdit_user)
        self.hBoxLayout_db.addWidget(self.lineEdit_password)
        self.hBoxLayout_db.addWidget(self.btn_test_connection)

        self.hBoxLayout_file.addWidget(self.label_excel_path)
        self.hBoxLayout_file.addWidget(self.lineEdit_excel_path)
        self.hBoxLayout_file.addWidget(self.btn_select_path)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        self.vBoxLayout.setSpacing(30)

        self.vBoxLayout.addLayout(self.hBoxLayout_db)

        self.vBoxLayout.addLayout(self.hBoxLayout_file)

        self.vBoxLayout.addWidget(self.btn_start)
        self.vBoxLayout.addWidget(self.textEdit_log)

        self.__initWidget()

        self.lineEdit_excel_path.setText(cfg.updateCert_excel_path.value)

    def __initWidget(self):
        self.view.setObjectName("")
        self.setObjectName("UpdateCertInterface")

        self.setWidget(self.view)
        self.setWidgetResizable(True)

    def __initLayout(self):
        self.hBoxLayout_db.setSpacing(8)

    @Slot(str)
    def logInfo(self, info):
        """
        打印日志
        """
        self.textEdit_log.append(info)

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

    def testConnection(self):
        host = self.lineEdit_host.text()
        dbname = self.lineEdit_dbname.text()
        user = self.lineEdit_user.text()
        password = self.lineEdit_password.text()

        if not all([host, dbname, user, password]):
            self.createErrorInfoBar("错误", "请填写完整的数据库信息")
            return

        try:
            conn = pg.connect(
                host=host,
                dbname=dbname,
                user=user,
                password=password,
            )

            cursor = conn.cursor()
            cursor.execute("SELECT 1")
            conn.commit()
            self.createSuccessInfoBar("成功", "连接成功")
        except Exception as e:
            self.createErrorInfoBar("失败", f"连接失败: {e}")

    def update(self):
        host = self.lineEdit_host.text()
        dbname = self.lineEdit_dbname.text()
        user = self.lineEdit_user.text()
        password = self.lineEdit_password.text()

        # 检查是否有空值
        if not all([host, dbname, user, password]):
            self.createErrorInfoBar("错误", "请填写完整的数据库信息")
            return

        root_dir = self.lineEdit_excel_path.text()
        if not root_dir:
            self.createErrorInfoBar("错误", "请选择 Excel 文件")
            return

        db_config = {
            "host": host,
            "dbname": dbname,
            "user": user,
            "password": password,
        }

        root_dir = Path(self.lineEdit_excel_path.text())

        self.worker = UpdateWorker(db_config, root_dir)
        self.worker.logInfo.connect(self.logInfo)
        self.worker.start()
