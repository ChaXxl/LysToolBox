# coding:utf-8
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from typing import Optional, override
from threading import Lock

import pillow_avif
import filetype
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

        # 添加文件缓存
        self._files_cache = []

        # 初始化锁
        self.lock = Lock()

    def _scan_files(self):
        """扫描所有支持的图片文件并缓存结果"""
        if not self._files_cache:
            self._files_cache = [
                f
                for f in self.root_dir.rglob("*")
                if f.suffix.lower() in self.SUPPORTED_FORMATS
            ]
        return self._files_cache

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

    def process_file(self, file_path: Path) -> None:
        """统一处理单个文件的逻辑"""
        try:
            #  获取锁来修改共享资源
            with self.lock:
                self.cur += 1
                val = int(self.cur / self.total * 100)
                self.setProgress.emit(val)
                self.setProgressInfo.emit(self.cur, self.total)

            # 检测文件格式
            mime_type = filetype.guess_mime(str(file_path))
            if mime_type is None:
                self.logInfo.emit(f"文件格式检测失败: {file_path}")
                return

            # 处理 JPEG 文件
            if mime_type == "image/jpeg":
                if file_path.suffix.lower() == ".jpeg":
                    file_path.rename(file_path.with_suffix(".jpg"))
                return

            # 处理需要转换的文件
            if mime_type != "image/jpeg":
                target_path = file_path.with_suffix(".jpg")
                try:
                    if mime_type == "image/webp":
                        try:
                            self._convert_image(file_path, target_path)
                        except Exception:
                            img = iio.imread(file_path, index=0)
                            iio.imwrite(target_path, img)
                    else:
                        self._convert_image(file_path, target_path)

                    # 转换成功后删除源文件
                    if target_path.exists():
                        file_path.unlink(missing_ok=True)
                except Exception as e:
                    self.logInfo.emit(f"转换失败: {file_path} - {str(e)}")

        except Exception as e:
            self.logInfo.emit(f"处理失败: {file_path} - {str(e)}")

    def _convert_image(self, source_path: Path, target_path: Path):
        """
        统一的图片转换处理函数
        """
        with Image.open(source_path) as img:
            if img.mode in ("RGBA", "LA"):
                img = self.fill_transparent_background(img)
            img.convert("RGB").save(target_path, "JPEG", quality=100)

    @override
    def run(self):
        start = datetime.now()

        # 扫描文件并缓存
        files = self._scan_files()

        # 统计总共有多少张图片要处理
        self.total = len(files)

        count_before = self.total

        self.setProgressInfo.emit(0, self.total)
        self.logInfo.emit(f"共有 {count_before} 张图片\n")

        # 使用多线程处理文件
        with ThreadPoolExecutor(max_workers=min(32, (self.total + 3) // 4)) as t:
            list(t.map(self.process_file, files))

        # 重新统计处理后的文件数
        self._files_cache = []  # 清除缓存以重新扫描
        count_after = len(self._scan_files())

        self.logInfo.emit(
            f"\n耗时: {datetime.now() - start}. 转换前有 {count_before} 张图片, 转换后有 {count_after} 张图片"
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

        # 按钮 用于开始转换
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
