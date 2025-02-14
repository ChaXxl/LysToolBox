# coding:utf-8
from pathlib import Path
from typing import Optional, Tuple, override

import openpyxl
import psycopg
import psycopg as pg
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import (BodyLabel, InfoBar, InfoBarPosition, LineEdit,
                            PasswordLineEdit, PushButton, TextEdit)

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEdit
from view.interface.gallery_interface import GalleryInterface


class GetQuaNameFromDB(QThread):
    logInfo = Signal(str)

    def __init__(self, db_config: dict, excel_path: Path):
        super().__init__()

        self.db_config = db_config
        self.excel_path = excel_path

        self.conn: Optional[psycopg.BaseConnection] = None
        self.cursor: Optional[psycopg.Cursor] = None

        # 统计处理的行数
        self.processed_rows = 0

    def query_db(self, query: str, params: tuple):
        """通用的数据库查询函数"""
        try:
            self.cursor.execute(query, params)
            # 获取查询结果
            result = self.cursor.fetchone()  # 如果是单行结果，使用 fetchone()
            return result
        except Exception as e:
            self.logInfo.emit(f"数据查询失败: {query} {params} {e}")
            return None

    def get_qualification_name(self, store_name: str) -> Optional[str]:
        """根据药店名称查询资质名称"""
        query = """SELECT qualification_name FROM store_info WHERE store_name = %s;"""
        res = self.query_db(query, (store_name,))
        return res[0] if res else ""

    def get_qualification_name_pdd(self, store_homepage: str) -> Tuple[str, str]:
        """查询拼多多店铺资质名称"""
        query = "SELECT store_name, qualification_name FROM store_info WHERE store_homepage = %s;"
        res = self.query_db(query, (store_homepage,))
        return (res[0], res[1]) if res else ("", "")

    def process_row(
        self, store_name: str, store_homepage: str, row: tuple, platform: str
    ):
        """处理单行数据, 查询并返回需要更新的值"""
        if platform == "拼多多":
            store_name, qualification_name = self.get_qualification_name_pdd(
                store_homepage
            )
        else:
            qualification_name = self.get_qualification_name(store_name)

        return store_name, qualification_name if qualification_name else ""

    def readExcel(self, excel_file: Path):
        """读取并处理 Excel 文件"""
        try:
            workbook = openpyxl.load_workbook(excel_file)
            sheet = workbook.active

            # 遍历 Excel 行，更新资质名称
            for row in sheet.iter_rows(min_row=2):
                store_name = row[1].value  # 药店名称在第二列
                store_homepage = row[2].value  # 店铺主页在第三列
                platform = row[9].value
                qualification_name = row[3].value  # 资质名称在第四列

                # 已有资质名称，不需要更新
                if (
                    qualification_name is not None
                    and qualification_name != ""
                    and qualification_name != "NaN"
                ):
                    continue

                store_name, qualification_name = self.process_row(
                    store_name, store_homepage, row, platform
                )

                if not qualification_name:
                    continue

                # 更新 Excel 表格中的资质名称
                row[1].value = store_name
                row[3].value = qualification_name

                # 记录处理的行数
                self.processed_rows += 1

                self.logInfo.emit(
                    f"\t\t更新 {excel_file.stem}{row[0].row}行的  {store_name} 的资质名称为 {qualification_name}"
                )

            workbook.save(excel_file)

        except Exception as e:
            self.logInfo.emit(f"处理失败: {excel_file.name} {e}")

    def process_all_excels(self):
        """处理所有 Excel 文件"""
        if self.excel_path.is_file():
            self.readExcel(self.excel_path)
            return

        for excel_file in self.excel_path.rglob("*.xlsx"):
            if any(keyword in excel_file.stem for keyword in ["~", "对照", "排查"]):
                continue

            self.readExcel(excel_file)

    @override
    def run(self):
        """运行主程序"""
        self.conn = pg.connect(**self.db_config)
        self.cursor = self.conn.cursor()

        try:
            self.process_all_excels()
        finally:
            self.conn.close()

        self.logInfo.emit(f"\n处理完成, 共更新 {self.processed_rows} 行数据")


class WriteExcelInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__(title="从数据库查询资质写入 Excel", parent=parent)

        self.view = QWidget(self)

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout_db = QHBoxLayout()
        self.hBoxLayout_file = QHBoxLayout()

        # host
        self.lineEdit_host = LineEdit()
        self.lineEdit_host.setPlaceholderText("IP 地址")
        self.lineEdit_host.setText(cfg.writeExcel_host.value)
        self.lineEdit_host.textChanged.connect(
            lambda: cfg.set(cfg.writeExcel_host, self.lineEdit_host.text())
        )

        # 数据库名称
        self.lineEdit_dbname = LineEdit()
        self.lineEdit_dbname.setPlaceholderText("数据库名称")
        self.lineEdit_dbname.setText(cfg.writeExcel_dbname.value)
        self.lineEdit_dbname.textChanged.connect(
            lambda: cfg.set(cfg.writeExcel_dbname, self.lineEdit_dbname.text())
        )

        # 用户名
        self.lineEdit_user = LineEdit()
        self.lineEdit_user.setPlaceholderText("用户名")
        self.lineEdit_user.setText(cfg.writeExcel_user.value)
        self.lineEdit_user.textChanged.connect(
            lambda: cfg.set(cfg.writeExcel_user, self.lineEdit_user.text())
        )

        # 密码
        self.lineEdit_password = PasswordLineEdit()
        self.lineEdit_password.setPlaceholderText("密码")
        self.lineEdit_password.setText(cfg.writeExcel_password.value)
        self.lineEdit_password.textChanged.connect(
            lambda: cfg.set(cfg.writeExcel_password, self.lineEdit_password.text())
        )

        # 测试连接按钮
        self.btn_test_connection = PushButton(text="测试连接")
        self.btn_test_connection.clicked.connect(self.testConnection)

        self.label_excel_path = BodyLabel(text="Excel 文件所在文件夹: ")

        # 显示 Excel 文件所在文件夹的文本框
        self.lineEdit_excel_path = DropableLineEdit()
        self.lineEdit_excel_path.setPlaceholderText(
            "请选择或者拖入 Excel 文件所在文件夹"
        )
        self.lineEdit_excel_path.textChanged.connect(
            lambda: cfg.set(cfg.writeExcel_excel_path, self.lineEdit_excel_path.text())
        )

        # 选择路径的按钮
        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_excel_path.setText(
                QFileDialog.getOpenFileName(self, "选择文件夹")[0]
            )
        )

        # 开始按钮
        self.btn_start = PushButton(text="开始")
        self.btn_start.clicked.connect(self.start)

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

        self.worker: Optional[GetQuaNameFromDB] = None

        self.lineEdit_excel_path.setText(cfg.writeExcel_excel_path.value)

    def __initWidget(self):
        self.view.setObjectName("")
        self.setObjectName("WriteExcelInterface")

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

    def start(self):
        self.textEdit_log.clear()

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
            self.createErrorInfoBar("错误", "请选择 Excel 文件所在文件夹")
            return

        db_config = {
            "host": host,
            "dbname": dbname,
            "user": user,
            "password": password,
        }

        root_dir = Path(self.lineEdit_excel_path.text())

        self.worker = GetQuaNameFromDB(db_config, root_dir)
        self.worker.logInfo.connect(self.logInfo)
        self.worker.start()
