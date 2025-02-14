# coding:utf-8
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import BodyLabel, PushButton, TextBrowser

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEditDir
from view.interface.gallery_interface import GalleryInterface


class AnalysisWorker(QThread):
    total_counts_signal = Signal(str)
    empty_counts_signal = Signal(str)
    platform_counts_signal = Signal(str)

    def __init__(self, root_dir: Path):
        super().__init__()

        self.root_dir = root_dir

    def run(self):
        total_files = 0
        total_counts_dict = {}
        empty_counts_dict = {}
        platform_counts_dict = {}

        for excel_file in self.root_dir.glob("*.xlsx"):
            if any(keyword in excel_file.stem for keyword in ["~", "对照", "排查"]):
                continue

            total_files += 1

            df = pd.read_excel(excel_file, usecols=["资质名称", "平台"])

            # 统计总行数
            total_counts_dict[excel_file.stem] = df.shape[0]

            # 统计资质名称为空的行数
            empty_counts = int(df["资质名称"].isna().sum())
            empty_counts_dict[excel_file.stem] = empty_counts

            # 统计各平台资质名称为空的行数
            mask = df["资质名称"].isna()
            count_dict = df.loc[mask, "平台"].value_counts().to_dict()

            for platform, count in count_dict.items():
                if platform not in platform_counts_dict:
                    platform_counts_dict[platform] = 0
                platform_counts_dict[platform] += count

        # 计算总和
        total_counts = sum(total_counts_dict.values())
        empty_counts = sum(empty_counts_dict.values())
        platform_counts = sum(platform_counts_dict.values())

        # 对字典从大到小排序
        total_counts_dict = sorted(
            total_counts_dict.items(), key=lambda x: x[1], reverse=True
        )
        empty_counts_dict = sorted(
            empty_counts_dict.items(), key=lambda x: x[1], reverse=True
        )
        platform_counts_dict = sorted(
            platform_counts_dict.items(), key=lambda x: x[1], reverse=True
        )

        # 打印结果
        self.total_counts_signal.emit(
            f"textEdit_total_counts <font color='#D15051'>共有 {total_files} 个文件, 总行数: {total_counts}</font>\n"
        )
        for excel_file, count in total_counts_dict:
            self.total_counts_signal.emit(
                f"textEdit_total_counts {excel_file}: {count}"
            )

        self.empty_counts_signal.emit(
            f"textEdit_empty_counts <font color='#D15051'>资质名称为空的总行数: {empty_counts}</font>\n"
        )
        for platform, count in empty_counts_dict:
            self.empty_counts_signal.emit(f"textEdit_empty_counts {platform}: {count}")

        self.platform_counts_signal.emit(
            f"textEdit_platform_counts <font color='#D15051'>各平台资质名称为空的总行数: {platform_counts}</font>\n"
        )
        for platform, count in platform_counts_dict:
            self.platform_counts_signal.emit(
                f"textEdit_platform_counts {platform}: {count}"
            )


class StatisticsInterface(GalleryInterface):
    total_counts_signal = Signal(str)
    empty_counts_signal = Signal(str)
    platform_counts_signal = Signal(str)

    def __init__(self, parent=None):
        super().__init__(title="统计数据", parent=parent)

        self.view = QWidget(self)

        # 状态提示
        self.stateTooltip = None

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout = QHBoxLayout()
        self.hLayout_count = QHBoxLayout()

        self.label_excel_path = BodyLabel(text="Excel 文件所在文件夹: ")

        # 显示 Excel 文件所在文件夹的文本框
        self.lineEdit_excel_path = DropableLineEditDir()
        self.lineEdit_excel_path.setPlaceholderText(
            "请选择或者拖入 Excel 文件所在文件夹"
        )
        self.lineEdit_excel_path.textChanged.connect(self.analysis)
        self.lineEdit_excel_path.textChanged.connect(
            lambda: cfg.set(cfg.statistics_excel_path, self.lineEdit_excel_path.text())
        )

        # 选择路径的按钮
        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_excel_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 刷新按钮
        self.btn_refresh = PushButton(text="刷新")
        self.btn_refresh.clicked.connect(self.analysis)

        # 统计总行数
        self.textEdit_total_counts = TextBrowser()
        self.textEdit_total_counts.setPlaceholderText("显示总行数")

        # 统计总的空行数
        self.textEdit_empty_counts = TextBrowser()
        self.textEdit_empty_counts.setPlaceholderText("显示空行数")

        #  统计各平台空行数
        self.textEdit_platform_counts = TextBrowser()
        self.textEdit_platform_counts.setPlaceholderText("显示各平台空行数")

        self.hLayout_count.addWidget(self.textEdit_total_counts)
        self.hLayout_count.addWidget(self.textEdit_empty_counts)
        self.hLayout_count.addWidget(self.textEdit_platform_counts)

        # 横向布局添加控件
        self.hBoxLayout.addWidget(self.label_excel_path)
        self.hBoxLayout.addWidget(self.lineEdit_excel_path)
        self.hBoxLayout.addWidget(self.btn_select_path)

        # 纵向布局添加布局
        self.vBoxLayout.addLayout(self.hBoxLayout)
        self.vBoxLayout.addWidget(self.btn_refresh)
        self.vBoxLayout.addLayout(self.hLayout_count)

        self.__initWidget()

        self.worker = None

        self.lineEdit_excel_path.setText(cfg.statistics_excel_path.value)

    def __initWidget(self):
        self.view.setObjectName("统计数据")
        self.setObjectName("StatisticsInterface")

        self.setWidget(self.view)
        self.setWidgetResizable(True)

    @Slot()
    def logInfo(self, msg: str):
        """
        打印日志
        """
        i = msg.split()[0]
        msg = msg.replace(i, "")

        # 去掉开头的空格
        if msg.startswith(" "):
            msg = msg[1:]

        if not hasattr(self, i):
            return

        obj: TextBrowser = getattr(self, i)

        # 保存当前滚动条位置
        scroll_position = obj.verticalScrollBar().value()

        if "<font" in msg:
            obj.setMarkdown(msg)
            obj.append("")
        else:
            obj.append(msg)

        # 恢复滚动条位置
        obj.verticalScrollBar().setValue(scroll_position)

    def analysis(self):
        """
        分析
        """
        root_dir = Path(self.lineEdit_excel_path.text())

        if not root_dir or not root_dir.exists():
            return

        self.textEdit_total_counts.clear()
        self.textEdit_empty_counts.clear()
        self.textEdit_platform_counts.clear()

        self.worker = AnalysisWorker(root_dir)
        self.worker.total_counts_signal.connect(self.logInfo)
        self.worker.empty_counts_signal.connect(self.logInfo)
        self.worker.platform_counts_signal.connect(self.logInfo)

        self.worker.start()
