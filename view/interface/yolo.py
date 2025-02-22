# coding:utf-8
from datetime import datetime
from pathlib import Path
from typing import Optional, override

import cv2
import numpy as np
import onnxruntime as ort
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
from utils.classnames import CLASS_NAMES
from view.components.dropable_lineEdit import DropableLineEditDir, DropableLineEditOnnx
from view.interface.gallery_interface import GalleryInterface


class YoloWorker(QThread):
    logInfo = Signal(str)
    setProgress = Signal(int)
    setProgressInfo = Signal(int, int)

    # 定义支持的文件格式集合
    SUPPORTED_FORMATS = {".jpeg", ".jpg", ".webp", ".png", ".avif", ".gif"}

    def __init__(
        self,
        img_dir: Path,
        model_path: Path,
        output_dir: Path,
        conf_thresh=0.85,
        iou_thresh=0.5,
    ):
        super().__init__()

        self.img_dir = img_dir
        self.model_path = model_path
        self.output_dir = output_dir
        self.output_dir.mkdir(exist_ok=True)

        self.conf_thresh = conf_thresh
        self.iou_thresh = iou_thresh

        self.classes = CLASS_NAMES

        self.colors = np.random.randint(
            0, 255, size=(len(self.classes), 3), dtype="uint8"
        )

        # 加载 ONNX 模型
        self.session = ort.InferenceSession(
            model_path,
            providers=(
                ["CUDAExecutionProvider", "CPUExecutionProvider"]
                if ort.get_device() == "GPU"
                else ["CPUExecutionProvider"]
            ),
        )
        self.input_shape = self.session.get_inputs()[0].shape[
            2:4
        ]  # 模型输入尺寸 (H, W)

    def letterbox(
        self,
        img: cv2.Mat,
        new_shape=(640, 640),
        color=(114, 114, 114),
        auto=False,
        scaleFill=False,
        scaleup=True,
    ):
        """
        将图像进行 letterbox 填充，保持纵横比不变，并缩放到指定尺寸。
        """
        shape = img.shape[:2]  # 当前图像的宽高

        if isinstance(new_shape, int):
            new_shape = (new_shape, new_shape)

        # 计算缩放比例
        r = min(
            new_shape[0] / shape[0], new_shape[1] / shape[1]
        )  # 选择宽高中最小的缩放比
        if not scaleup:  # 仅缩小，不放大
            r = min(r, 1.0)

        # 缩放后的未填充尺寸
        new_unpad = (int(round(shape[1] * r)), int(round(shape[0] * r)))

        # 计算需要的填充
        dw, dh = (
            new_shape[1] - new_unpad[0],
            new_shape[0] - new_unpad[1],
        )  # 计算填充的尺寸
        dw /= 2  # padding 均分
        dh /= 2

        # 缩放图像
        if shape[::-1] != new_unpad:  # 如果当前图像尺寸不等于 new_unpad，则缩放
            img = cv2.resize(img, new_unpad, interpolation=cv2.INTER_LINEAR)

        # 为图像添加边框以达到目标尺寸
        top, bottom = int(round(dh)), int(round(dh))
        left, right = int(round(dw)), int(round(dw))
        img = cv2.copyMakeBorder(
            img, top, bottom, left, right, cv2.BORDER_CONSTANT, value=color
        )

        # 确保填充后的图像尺寸为 640x640
        img = cv2.resize(
            img, (new_shape[1], new_shape[0]), interpolation=cv2.INTER_LINEAR
        )

        return img, (r, r), (dw, dh)

    def preprocess(self, image_path: Path) -> tuple[cv2.Mat, np.array, tuple[int, int]]:
        """预处理输入图像，返回调整后的图像和比例信息"""
        with open(image_path, "rb") as f:
            img_array = np.asarray(bytearray(f.read()), dtype=np.uint8)

        img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)

        h, w = img.shape[:2]

        # 将图像颜色空间从 BGR 转换为 RGB
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

        # 保持宽高比，进行 letterbox 填充, 使用模型要求的输入尺寸
        img, self.ratio, (self.dw, self.dh) = self.letterbox(img, new_shape=(640, 640))

        # 通过除以 255.0 来归一化图像数据
        img_normalized = np.array(img) / 255.0

        # 将图像的通道维度移到第一维
        img_transposed = np.transpose(img_normalized, (2, 0, 1))

        # 扩展图像数据的维度，以匹配模型输入的形状
        image_data = np.expand_dims(img_transposed, axis=0).astype(np.float32)

        return img, image_data, (w, h)

    def postprocess(
        self,
        outputs,
        img_path: Path,
        output_dir: Path,
    ):
        # 解析模型输出
        outputs = np.transpose(np.squeeze(outputs[0]))

        conf_id_list = []

        for i in range(outputs.shape[0]):

            classes_scores = outputs[i][4:]

            conf = np.amax(classes_scores)

            if conf < self.conf_thresh:
                continue

            class_id = np.argmax(classes_scores)

            conf_id_list.append((conf, class_id))

        # 按置信度降序排序
        conf_id_list.sort(key=lambda x: x[0], reverse=True)

        # 如果识别到的物体不是当前目录的药品名，则移动到对应目录
        medicine_name = img_path.stem.split("_")[0]
        if not conf_id_list or medicine_name not in CLASS_NAMES[conf_id_list[0][1]]:
            new_dir = output_dir / medicine_name
            new_dir.mkdir(exist_ok=True)
            img_path.rename(new_dir / img_path.name)

    @override
    def run(self):
        start = datetime.now()

        imgs = [
            f for f in self.img_dir.rglob("*") if f.suffix in self.SUPPORTED_FORMATS
        ]

        fail_imgs: list[Path] = []

        self.setProgressInfo.emit(0, len(imgs))

        for i, image_path in enumerate(imgs):
            self.setProgress.emit((i + 1) / len(imgs) * 100)
            self.setProgressInfo.emit(i + 1, len(imgs))

            try:
                original_img, preprocessed_img, original_size = self.preprocess(
                    image_path
                )
                outputs = self.session.run(
                    None, {self.session.get_inputs()[0].name: preprocessed_img}
                )
                self.postprocess(outputs, image_path, self.output_dir)
            except Exception as e:
                fail_imgs.append(image_path)
                self.logInfo.emit(f"推理失败: {str(e)}")
                continue

        # 打印推理失败的图片
        if fail_imgs:
            self.logInfo.emit("\n推理失败的图片:")
            for img in fail_imgs:
                self.logInfo.emit(str(img))

        # 计算还剩多少张图片
        remain_imgs = len(
            [f for f in self.img_dir.rglob("*") if f.suffix in self.SUPPORTED_FORMATS]
        )

        self.logInfo.emit(
            f"\n耗时: {datetime.now() - start}. 共有 {len(imgs)} 张图片, 识别后剩余 {remain_imgs} 张图片"
        )


class YoloInterface(GalleryInterface):
    def __init__(self, parent=None):
        super().__init__("通过 yolo 识别药品", parent=parent)

        self.view = QWidget(self)

        # 状态提示
        self.stateTooltip = None

        self.vBoxLayout = QVBoxLayout(self.view)
        self.hBoxLayout_img = QHBoxLayout()
        self.hBoxLayout_onnx = QHBoxLayout()
        self.hBoxLayout_output = QHBoxLayout()
        self.hBoxLayout_progress = QHBoxLayout()

        self.label_img = BodyLabel(text="图片所在文件夹: ")
        # 模型文件路径文本框
        self.lineEdit_img_path = DropableLineEditDir()
        self.lineEdit_img_path.setPlaceholderText("请选择或者拖入图片所在文件夹")
        self.lineEdit_img_path.textChanged.connect(
            lambda: cfg.set(cfg.yolo_img_path, self.lineEdit_img_path.text())
        )

        self.label_onnx = BodyLabel(text="onnx 模型文件: ")
        # Excel 文件所在文件夹的文本框
        self.lineEdit_onnx_path = DropableLineEditOnnx()
        self.lineEdit_onnx_path.setPlaceholderText("请选择或者拖入 onnx 模型文件")
        self.lineEdit_onnx_path.textChanged.connect(
            lambda: cfg.set(cfg.yolo_onnx_path, self.lineEdit_onnx_path.text())
        )

        self.label_output = BodyLabel(text="输出文件夹: ")
        # 输出文件夹的文本框
        self.lineEdit_output_path = DropableLineEditDir()
        self.lineEdit_output_path.setPlaceholderText("请选择或者拖入输出文件夹")
        self.lineEdit_output_path.textChanged.connect(
            lambda: cfg.set(cfg.yolo_output_path, self.lineEdit_output_path.text())
        )

        # 选择路径的按钮
        self.btn_select_img_path = PushButton(text="···")
        self.btn_select_img_path.clicked.connect(
            lambda: self.lineEdit_img_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 选择模型文件路径的按钮
        self.btn_select_onnx_path = PushButton(text="···")
        self.btn_select_onnx_path.clicked.connect(
            lambda: self.lineEdit_onnx_path.setText(
                QFileDialog.getOpenFileName(
                    self, "选择文件", filter="onnx 模型 (*.onnx)"
                )[0]
            )
        )

        # 选择输出路径的按钮
        self.btn_select_output_path = PushButton(text="···")
        self.btn_select_output_path.clicked.connect(
            lambda: self.lineEdit_img_path.setText(
                QFileDialog.getExistingDirectory(self, "选择文件夹")
            )
        )

        # 下载按钮
        self.btn_download = PushButton(text="识别")
        self.btn_download.clicked.connect(self.start)

        # 文本框 用于打印日志
        self.textEdit_log = TextEdit()
        self.textEdit_log.setPlaceholderText("此处是用来打印日志的")

        # 进度条
        self.progressBar = ProgressBar()

        # 进度提示标签
        self.label_progress = BodyLabel(text="0/0")

        self.hBoxLayout_img.addWidget(self.label_img)
        self.hBoxLayout_img.addWidget(self.lineEdit_img_path)
        self.hBoxLayout_img.addWidget(self.btn_select_img_path)

        self.hBoxLayout_onnx.addWidget(self.label_onnx)
        self.hBoxLayout_onnx.addWidget(self.lineEdit_onnx_path)
        self.hBoxLayout_onnx.addWidget(self.btn_select_onnx_path)

        self.hBoxLayout_output.addWidget(self.label_output)
        self.hBoxLayout_output.addWidget(self.lineEdit_output_path)
        self.hBoxLayout_output.addWidget(self.btn_select_output_path)

        self.hBoxLayout_progress.addWidget(self.progressBar)
        self.hBoxLayout_progress.addWidget(self.label_progress)

        self.vBoxLayout.addLayout(self.hBoxLayout_img)
        self.vBoxLayout.addLayout(self.hBoxLayout_onnx)
        self.vBoxLayout.addLayout(self.hBoxLayout_output)

        self.vBoxLayout.addWidget(self.btn_download)
        self.vBoxLayout.addWidget(self.textEdit_log)

        # 进度条布局
        self.vBoxLayout.addLayout(self.hBoxLayout_progress)

        self.__initWidget()

        # 从配置文件中读取路径
        self.lineEdit_img_path.setText(cfg.yolo_img_path.value)
        self.lineEdit_onnx_path.setText(cfg.yolo_onnx_path.value)
        self.lineEdit_output_path.setText(cfg.yolo_output_path.value)

        self.worker: Optional[YoloWorker] = None

    def __initWidget(self):
        self.view.setObjectName("")
        self.setObjectName("YoloInterface")

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
        self.lineEdit_onnx_path.setEnabled(True)
        self.lineEdit_output_path.setEnabled(True)

        self.btn_select_img_path.setEnabled(True)
        self.btn_select_onnx_path.setEnabled(True)
        self.btn_select_output_path.setEnabled(True)
        self.btn_download.setEnabled(True)

        if self.stateTooltip is not None:
            self.stateTooltip.hide()

        self.createSuccessInfoBar("完成", "图片识别完成 ✅")

    def start(self):
        self.textEdit_log.clear()

        # 检查是否选择了文件夹
        img_dir = self.lineEdit_img_path.text()
        if not img_dir:
            self.createErrorInfoBar("错误", "请选择图片所在文件夹")
            return

        # 检查是否选择了 onnx 文件
        onnx_path = self.lineEdit_onnx_path.text()
        if not onnx_path:
            self.createErrorInfoBar("错误", "请选择 onnx 模型文件")
            return

        # 检查是否选择了输出文件夹
        output_dir = self.lineEdit_output_path.text()
        if not output_dir:
            self.createErrorInfoBar("错误", "请选择输出文件夹")
            return

        img_dir = Path(self.lineEdit_img_path.text())
        onnx_path = Path(self.lineEdit_onnx_path.text())
        output_dir = Path(self.lineEdit_output_path.text())

        self.lineEdit_img_path.setEnabled(False)
        self.lineEdit_onnx_path.setEnabled(False)

        self.btn_select_img_path.setEnabled(False)
        self.btn_select_onnx_path.setEnabled(False)

        self.lineEdit_output_path.setEnabled(False)
        self.btn_select_output_path.setEnabled(False)

        self.btn_download.setEnabled(False)

        self.worker = YoloWorker(img_dir, onnx_path, output_dir)

        self.worker.logInfo.connect(self.logInfo)
        self.worker.finished.connect(self.finished)
        self.worker.setProgress.connect(self.setProgress)
        self.worker.setProgressInfo.connect(self.setProgressInfo)

        self.worker.start()
