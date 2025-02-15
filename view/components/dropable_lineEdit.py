from pathlib import Path
from typing import overload, override

from qfluentwidgets import LineEdit


class DropableLineEdit(LineEdit):
    """
    可拖拽的文件选择器
    """

    def __init__(self):
        super().__init__()
        self.acceptDrops()

    @override
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    @override
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                url = urls[0]
                self.setText(url.toLocalFile())
        else:
            event.ignore()


class DropableLineEditDir(DropableLineEdit):
    """
    可拖拽的文件夹选择器
    """

    def __init__(self):
        super().__init__()
        self.acceptDrops()

    @override
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    @override
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()

            if not urls:
                return

            url = urls[0]

            if not Path(url.toLocalFile()).is_dir():
                return

            self.setText(url.toLocalFile())

        else:
            event.ignore()


class DropableLineEditOnnx(LineEdit):
    """
    可拖拽的 ONNX 文件选择器
    """

    def __init__(self):
        super().__init__()
        self.acceptDrops()

    @override
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    @override
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                filename = urls[0].toLocalFile()

                if not filename.endswith(".onnx"):
                    return

                self.setText(filename)
        else:
            event.ignore()


class DropableLineEditExcel(LineEdit):
    """
    可拖拽的 Excel 文件选择器
    """

    def __init__(self):
        super().__init__()
        self.acceptDrops()

    @override
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    @override
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                filename = urls[0].toLocalFile()

                if not filename.endswith(".xlsx"):
                    return

                self.setText(filename)
        else:
            event.ignore()
