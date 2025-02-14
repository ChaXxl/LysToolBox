# coding:utf-8
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Optional, override

import imageio.v3 as iio
from PIL import Image
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from qfluentwidgets import (
    BodyLabel,
    InfoBar,
    InfoBarPosition,
    ProgressBar,
    PushButton,
    TextEdit,
)

from common.config import cfg
from view.components.dropable_lineEdit import DropableLineEdit
from view.interface.gallery_interface import GalleryInterface


class TransferWorker(QThread):
    logInfo = Signal(str)
    setProgress = Signal(int)
    setProgressInfo = Signal(int, int)

    # 定义支持的文件格式集合
    SUPPORTED_FORMATS = {".jpeg", ".jpg", ".webp", ".png", ".avif", ".gif"}

    def __init__(self, root_dir: Path):
        super().__init__()
        self.root_dir = root_dir

        self.cur = 0
        self.total = 0

    @staticmethod
    def detect_format(file_path: Path) -> Optional[str]:
        """
        使用 file 命令检测文件格式
        :param file_path: 要检测的文件路径
        :return: MIME类型字符串，如果检测失败返回None
        """
        result = subprocess.run(
            [
                "file",
                "--mime-type",
                str(file_path),
            ],
            capture_output=True,
            text=True,
        )
        mime_type = result.stdout.split(":")[-1].strip()
        return mime_type

    @staticmethod
    def fill_transparent_background(img: Image.Image) -> Image.Image:
        """
        为带有透明通道的图片添加白色背景
        :param img: 原始图片对象
        :return: 处理后的图片对象
        """
        if img.mode in ("RGBA", "LA"):
            # 创建白色背景
            background = Image.new("RGB", img.size, (255, 255, 255))

            # 将原图覆盖到白色背景上
            background.paste(img, mask=img.split()[-1])

            return background
        return img

    def fixSuffix(self):
        files = [
            f for f in self.root_dir.rglob("*") if f.suffix in self.SUPPORTED_FORMATS
        ]

        for file_path in files:
            try:
                # 更新进度条
                self.cur += 0.5
                val = int(self.cur / self.total * 100)
                self.setProgress.emit(val)
                self.setProgressInfo.emit(self.cur, self.total)

                # 检测文件格式
                mime_type = self.detect_format(file_path)
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
                self.logInfo.emit(f"修改后缀失败: {file_path} - {str(e)}")

    def convert(self):
        files = [
            f for f in self.root_dir.rglob("*") if f.suffix in self.SUPPORTED_FORMATS
        ]

        for file_path in files:
            try:
                # 更新进度条
                self.cur += 0.5
                val = int(self.cur / self.total * 100)
                self.setProgress.emit(val)
                self.setProgressInfo.emit(self.cur, self.total)

                # 检测文件格式
                mime_type = self.detect_format(file_path)

                # 构造目标文件路径
                target_path = file_path.with_suffix(".jpg")

                # 根据不同的文件类型进行相应的转换处理
                match mime_type:
                    case "image/webp":
                        try:
                            with Image.open(file_path) as img:
                                # 转换为 JPEG 格式并保存
                                img.convert("RGB").save(
                                    target_path, "JPEG", quality=100
                                )

                        except Exception as e:
                            if mime_type == "image/webp":
                                img = iio.imread(file_path, index=0)
                                iio.imwrite(target_path, img)

                    case "image/avif":
                        # 使用 avifdec 将 avif 文件转换为 png, 再将 png 转换为 jpg
                        subprocess.run(["avifdec", str(file_path), str(target_path)])

                    case "image/png":
                        with Image.open(file_path) as img:
                            img = self.fill_transparent_background(img)
                            img.convert("RGB").save(target_path, "JPEG", quality=100)

                    case "image/jpeg":
                        if file_path.suffix.lower() == ".jpg":
                            continue

                        if file_path.suffix.lower() == ".jpeg":
                            file_path.rename(file_path.with_suffix(".jpg"))
                            continue

                        with Image.open(file_path) as img:
                            img.save(target_path, "JPEG", quality=100)

                    case "image/jpg":
                        if file_path.suffix.lower() == ".jpg":
                            continue

                        with Image.open(file_path) as img:
                            img.save(target_path, "JPEG", quality=100)

                    case _:
                        self.logInfo.emit(
                            f"不支持的文件格式: {file_path} ({mime_type})"
                        )
                        continue

            except Exception as e:
                self.logInfo.emit(f"出错 - {e}")

            else:
                if target_path.exists():
                    # 转换成功后删除源文件
                    file_path.unlink(missing_ok=True)

    @override
    def run(self):
        start = datetime.now()

        files = [
            f for f in self.root_dir.rglob("*") if f.suffix in self.SUPPORTED_FORMATS
        ]

        # 统计总共有多少张图片要处理
        self.total = len(files)
        self.setProgressInfo.emit(0, self.total)

        count_before = len(files)
        self.logInfo.emit(f"共有 {count_before} 张图片\n")

        # 修改后缀
        self.fixSuffix()

        # 转换格式
        self.convert()

        files = [
            f for f in self.root_dir.rglob("*") if f.suffix in self.SUPPORTED_FORMATS
        ]
        count_after = len(files)

        self.logInfo.emit(
            f"\n耗时: {datetime.now() - start}. 修改前有 {count_before} 张图片, 修改后有 {count_after} 张图片"
        )


class ImgFormatTransInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("图片格式转换", parent=parent)

        self.view = QWidget(self)

        # 状态提示
        self.stateTooltip = None

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout = QHBoxLayout()
        self.hBoxLayout_progress = QHBoxLayout()

        self.label_img_path = BodyLabel(text="图片所在文件夹: ")

        # 显示 Excel 文件所在文件夹的文本框
        self.lineEdit_img_path = DropableLineEdit()
        self.lineEdit_img_path.setPlaceholderText("请选择或者拖入图片所在文件夹")
        self.lineEdit_img_path.textChanged.connect(
            lambda: cfg.set(cfg.downloadImg_img_path, self.lineEdit_img_path.text())
        )

        # 选择路径的按钮
        self.btn_select_path = PushButton(text="···")
        self.btn_select_path.clicked.connect(
            lambda: self.lineEdit_img_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 下载按钮
        self.btn_download = PushButton(text="转换")
        self.btn_download.clicked.connect(self.start)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        # 进度条
        self.progressBar = ProgressBar()

        # 设置取值范围
        self.progressBar.setRange(0, 100)

        # 进度提示标签
        self.label_progress = BodyLabel(text="0/0")

        self.hBoxLayout.addWidget(self.label_img_path)
        self.hBoxLayout.addWidget(self.lineEdit_img_path)
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
        self.lineEdit_img_path.setText(cfg.downloadImg_img_path.value)

        self.worker: Optional[TransferWorker] = None

    def __initWidget(self):
        self.view.setObjectName("")
        self.setObjectName("ImgFormatTransInterface")

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
        self.lineEdit_img_path.setEnabled(True)
        self.btn_select_path.setEnabled(True)
        self.btn_download.setEnabled(True)

        if self.stateTooltip is not None:
            self.stateTooltip.hide()

        self.createSuccessInfoBar("完成", "格式转换完成 ✅")

    def start(self):
        self.textEdit_log.clear()

        # 检查是否选择了文件夹
        root_dir = self.lineEdit_img_path.text()
        if not root_dir:
            self.createErrorInfoBar("错误", "请选择 Excel 文件所在文件夹")
            return

        root_dir = Path(self.lineEdit_img_path.text())

        self.lineEdit_img_path.setEnabled(False)
        self.btn_select_path.setEnabled(False)
        self.btn_download.setEnabled(False)

        self.worker = TransferWorker(root_dir)

        self.worker.logInfo.connect(self.logInfo)
        self.worker.finished.connect(self.finished)
        self.worker.setProgress.connect(self.setProgress)
        self.worker.setProgressInfo.connect(self.setProgressInfo)

        self.worker.start()
