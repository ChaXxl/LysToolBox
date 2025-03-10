# coding:utf-8
import asyncio
import subprocess
from datetime import datetime
from pathlib import Path
from platform import system
from typing import Optional, override

from mitmproxy.options import Options
from mitmproxy.tools.dump import DumpMaster
from openpyxl.reader.excel import load_workbook
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import (BodyLabel, InfoBar, InfoBarPosition, ProgressBar,
                            PushButton, TextEdit, TogglePushButton)

from common.config import cfg
from utils.mitm_addon import Addon
from view.components.dropable_lineEdit import (DropableLineEditDir,
                                               DropableLineEditExcel)
from view.interface.gallery_interface import GalleryInterface


class MitmProxySearchWorker(QThread):
    logInfo = Signal(str)
    setProgress = Signal(int)
    setProgressInfo = Signal(int, int)

    def __init__(
        self,
        excel_path: Optional[Path],
        keyword: str,
        output_dir: Path,
        proxy_ip: str,
        proxy_port: int,
    ):
        super().__init__()

        self.excel_path = excel_path
        self.keyword = keyword
        self.output_dir = output_dir

        self.addon = Addon()
        self.addon.add_text = self.logInfo
        self.options = Options(listen_host=proxy_ip, listen_port=proxy_port)
        self.m: Optional[DumpMaster] = None

    async def start_mitm(self):        
        self.m = DumpMaster(options=self.options)
        self.m.addons.add(self.addon)
        await self.m.run()

    @override
    def run(self):
        start = datetime.now()

        asyncio.run(self.start_mitm())

        self.logInfo.emit(f"\n耗时: {datetime.now() - start}")


class MitmProxySearchInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("通过 MitmProxy 代理搜索", parent=parent)

        self.view = QWidget(self)

        # 状态提示
        self.stateTooltip = None

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout_proxy = QHBoxLayout()
        self.hBoxLayout_excel = QHBoxLayout()
        self.hBoxLayout_output = QHBoxLayout()
        self.hBoxLayout_keyword = QHBoxLayout()
        self.hBoxLayout_progress = QHBoxLayout()

        # 设置代理
        self.label_proxy = BodyLabel(text="设置系统代理: ")
        self.lineEdit_proxy = DropableLineEditExcel()
        self.lineEdit_proxy.setPlaceholderText("请设置代理, 如 127.0.0.1:9999")
        self.lineEdit_proxy.textChanged.connect(
            lambda: cfg.set(cfg.mitmProxySearch_host, self.lineEdit_proxy.text())
        )

        self.btn_setProxy = TogglePushButton(text="打开代理")
        self.btn_setProxy.clicked.connect(self.on_btn_clicked_setProxy)

        # 待搜索药品的 Excel 文件
        self.label_excelPath = BodyLabel(text="待搜索药品的 Excel 文件: ")
        self.lineEdit_excelPath = DropableLineEditExcel()
        self.lineEdit_excelPath.setPlaceholderText("请选择或者拖入药品的 Excel 文件")
        self.lineEdit_excelPath.textChanged.connect(
            lambda: cfg.set(
                cfg.mitmProxySearch_excel_path, self.lineEdit_excelPath.text()
            )
        )

        self.btn_select_excel_path = PushButton(text="···")
        self.btn_select_excel_path.clicked.connect(
            lambda: self.lineEdit_excelPath.setText(
                QFileDialog.getOpenFileName(
                    self, "选择文件", filter="Excel 文件(*.xlsx)"
                )[0]
            )
        )

        # 输出文件夹
        self.label_output = BodyLabel(text="输出文件夹: ")
        self.lineEdit_output_path = DropableLineEditDir()
        self.lineEdit_output_path.setPlaceholderText("请选择或者拖入输出文件夹")
        self.lineEdit_output_path.textChanged.connect(
            lambda: cfg.set(
                cfg.mitmProxySearch_output_path, self.lineEdit_output_path.text()
            )
        )

        # 选择输出路径的按钮
        self.btn_select_output_path = PushButton(text="···")
        self.btn_select_output_path.clicked.connect(
            lambda: self.lineEdit_output_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 关键词
        self.label_keyword = BodyLabel(text="关键词: ")
        self.label_keyword.setMaximumWidth(100)
        self.lineEdit_keyword = DropableLineEditExcel()
        self.lineEdit_keyword.setPlaceholderText("")
        self.lineEdit_keyword.setMaximumWidth(500)
        self.lineEdit_keyword.textChanged.connect(
            lambda: cfg.set(cfg.mitmProxySearch_keyword, self.lineEdit_keyword.text())
        )

        # 修改
        self.btn_next = PushButton(text="修改")
        self.btn_next.clicked.connect(self.set_keyword)

        # 开始按钮
        self.btn_start = TogglePushButton(text="开始")
        self.btn_start.clicked.connect(self.start)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        # 进度条
        self.progressBar = ProgressBar()

        # 进度提示标签
        self.label_progress = BodyLabel(text="0/0")

        # 布局-设置代理
        self.hBoxLayout_proxy.addWidget(self.label_proxy)
        self.hBoxLayout_proxy.addWidget(self.lineEdit_proxy)
        self.hBoxLayout_proxy.addWidget(self.btn_setProxy)

        # 布局-待搜索药品的 Excel 文件
        self.hBoxLayout_excel.addWidget(self.label_excelPath)
        self.hBoxLayout_excel.addWidget(self.lineEdit_excelPath)
        self.hBoxLayout_excel.addWidget(self.btn_select_excel_path)

        # 布局-输出文件夹
        self.hBoxLayout_output.addWidget(self.label_output)
        self.hBoxLayout_output.addWidget(self.lineEdit_output_path)
        self.hBoxLayout_output.addWidget(self.btn_select_output_path)

        # 布局-关键词
        self.hBoxLayout_keyword.addWidget(self.label_keyword)
        self.hBoxLayout_keyword.addWidget(self.lineEdit_keyword)
        self.hBoxLayout_keyword.addWidget(self.btn_next)
        self.hBoxLayout_keyword.addWidget(self.btn_start)

        # 布局-进度条
        self.hBoxLayout_progress.addWidget(self.progressBar)
        self.hBoxLayout_progress.addWidget(self.label_progress)

        # 垂直布局
        self.vBoxLayout.addLayout(self.hBoxLayout_proxy)
        self.vBoxLayout.addLayout(self.hBoxLayout_excel)
        self.vBoxLayout.addLayout(self.hBoxLayout_output)
        self.vBoxLayout.addLayout(self.hBoxLayout_keyword)

        self.vBoxLayout.addWidget(self.textEdit_log)

        # 进度条
        self.vBoxLayout.addLayout(self.hBoxLayout_progress)

        self.__initWidget()

        # 从配置文件中读取路径
        self.lineEdit_proxy.setText(cfg.mitmProxySearch_host.value)
        self.lineEdit_excelPath.setText(cfg.mitmProxySearch_excel_path.value)
        self.lineEdit_output_path.setText(cfg.mitmProxySearch_output_path.value)
        self.lineEdit_keyword.setText(cfg.mitmProxySearch_keyword.value)

        self.worker: Optional[MitmProxySearchWorker] = None

        self.proxy_ip: str = ""
        self.proxy_port: str = ""

        self.btn_start_flag = False
        self.btn_setProxy_flag = False

        self.worker: Optional[MitmProxySearchWorker] = None

    def __initWidget(self):
        self.view.setObjectName("")
        self.setObjectName("MitmProxySearchInterface")

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

        # 滑动到底部
        self.textEdit_log.verticalScrollBar().setValue(
            self.textEdit_log.verticalScrollBar().maximum()
        )

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
        # 启用控件
        self.lineEdit_proxy.setEnabled(True)

        self.lineEdit_excelPath.setEnabled(True)
        self.btn_select_excel_path.setEnabled(True)

        self.lineEdit_output_path.setEnabled(True)
        self.btn_select_output_path.setEnabled(True)

        self.btn_start.setText("开始")
        self.btn_start_flag = False

        if self.stateTooltip is not None:
            self.stateTooltip.hide()

        self.createSuccessInfoBar("完成", "完成")

    @Slot()
    def set_keyword(self):
        """
        设置关键词
        """
        keyword = self.lineEdit_keyword.text()
        self.worker.addon.keyword = keyword

        filename = Path(self.lineEdit_output_path.text()) / f"{keyword}.xlsx"
        self.worker.addon.createExcel(filename)

        self.textEdit_log.append(f"\n\nExcel 保存在：{filename}\n\n")
        
    def on_btn_clicked_setProxy(self):
        """
        设置代理
        :return:
        """
        # Windows
        if system() == "Windows":
            from winproxy import ProxySetting

            p = ProxySetting()

            if self.btn_setProxy_flag:
                # 现在是关闭代理
                self.lineEdit_proxy.setEnabled(True)
                self.btn_setProxy.setText("打开代理")

                p.enable = False
                self.btn_setProxy_flag = False
            else:
                # 现在是把代理打开
                self.lineEdit_proxy.setEnabled(False)
                self.btn_setProxy.setText("关闭代理")

                proxy = self.lineEdit_proxy.text()
                p.server = proxy
                p.enable = True
                self.btn_setProxy_flag = True

            p.registry_write()

        # MacOS
        elif system() == "Darwin":
            if self.btn_setProxy_flag:
                #  关闭代理
                self.lineEdit_proxy.setEnabled(True)
                self.btn_setProxy.setText("打开代理")
                self.btn_setProxy_flag = False

                subprocess.run(["networksetup", "-setwebproxystate", "AX88179A", "off"])
                subprocess.run(
                    ["networksetup", "-setsecurewebproxystate", "AX88179A", "off"]
                )
                subprocess.run(
                    ["networksetup", "-setsocksfirewallproxystate", "AX88179A", "off"]
                )

            else:
                #  设置代理
                self.lineEdit_proxy.setEnabled(False)
                self.btn_setProxy.setText("关闭代理")
                self.btn_setProxy_flag = True

                subprocess.run(
                    ["networksetup", "-setwebproxy", "AX88179A", "127.0.0.1", "9999"]
                )
                subprocess.run(
                    [
                        "networksetup",
                        "-setsecurewebproxy",
                        "AX88179A",
                        "127.0.0.1",
                        "9999",
                    ]
                )
                subprocess.run(
                    [
                        "networksetup",
                        "-setsocksfirewallproxy",
                        "AX88179A",
                        "127.0.0.1",
                        "9999",
                    ]
                )

    def start(self):
        if self.btn_start_flag:
            # 启用控件
            self.lineEdit_excelPath.setEnabled(True)
            self.btn_select_excel_path.setEnabled(True)

            self.lineEdit_output_path.setEnabled(True)
            self.btn_select_output_path.setEnabled(True)

            self.btn_start.setText("开始")
            self.btn_start_flag = False

        else:
            self.textEdit_log.clear()

            if not self.btn_setProxy_flag:
                self.createErrorInfoBar("错误", "请先设置代理")
                return

            # 检查是否选择了输出文件夹
            output_dir = self.lineEdit_output_path.text()
            if not output_dir:
                self.createErrorInfoBar("错误", "请选择输出文件夹")
                return

            # 检查是否选择了待搜索药品的 Excel 文件
            excel_path = self.lineEdit_excelPath.text()

            # 关键词
            keyword = self.lineEdit_keyword.text()

            if not keyword:
                self.createErrorInfoBar("错误", "请输入关键词")
                return

            # if not any([excel_path, keyword]):
            #     self.createErrorInfoBar(
            #         "错误", "请选择待搜索药品的 Excel 文件或者关键词"
            #     )
            #     return
            #
            # if all([excel_path, keyword]):
            #     self.createErrorInfoBar(
            #         "错误", "只能选择待搜索药品的 Excel 文件或者关键词"
            #     )
            #     return

            # 获取代理 IP 和端口
            proxy = self.lineEdit_proxy.text().strip()
            proxy_ip, proxy_port = proxy.split(":")

            # 禁用控件
            self.lineEdit_proxy.setEnabled(False)

            self.lineEdit_excelPath.setEnabled(False)
            self.btn_select_excel_path.setEnabled(False)

            self.lineEdit_output_path.setEnabled(False)

            self.btn_start.setText("停止")

            # if excel_path:
            #     self.worker = MitmProxySearchWorker(
            #         Path(excel_path), "", Path(output_dir), proxy_ip, int(proxy_port)
            #     )
            # elif keyword:
            self.worker = MitmProxySearchWorker(
                None, keyword, Path(output_dir), proxy_ip, int(proxy_port)
            )

            self.btn_start_flag = True

            filename = Path(output_dir) / f"{keyword}.xlsx"

            self.textEdit_log.append(f"\n\nExcel 保存在：{filename}\n\n")

            self.worker.logInfo.connect(self.logInfo)
            self.worker.finished.connect(self.finished)
            self.worker.setProgress.connect(self.setProgress)
            self.worker.setProgressInfo.connect(self.setProgressInfo)

            self.worker.addon.createExcel(filename)
            self.worker.addon.keyword = keyword

            self.worker.start()
