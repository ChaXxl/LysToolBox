# coding:utf-8
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QVBoxLayout, QWidget
from qfluentwidgets import ScrollArea, TitleLabel

from common.style_sheet import StyleSheet


class ToolBar(QWidget):

    def __init__(self, title, parent=None):
        super().__init__(parent=parent)
        self.titleLabel = TitleLabel(title, self)

        self.vBoxLayout = QVBoxLayout(self)

        self.__initWidget()

    def __initWidget(self):
        self.setFixedHeight(60)
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(22, 12, 22, 12)
        self.vBoxLayout.addWidget(self.titleLabel)
        self.vBoxLayout.setAlignment(Qt.AlignTop)


class GalleryInterface(ScrollArea):
    def __init__(self, title: str, parent=None):
        """
        title: str
            The title of gallery

        parent: QWidget
            parent widget
        """
        super().__init__(parent=parent)
        self.view = QWidget(self)
        self.toolBar = ToolBar(title, self)
        self.vBoxLayout = QVBoxLayout(self.view)

        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setViewportMargins(0, self.toolBar.height(), 0, 0)
        self.setWidget(self.view)
        self.setWidgetResizable(True)

        self.vBoxLayout.setSpacing(30)
        self.vBoxLayout.setAlignment(Qt.AlignTop)
        self.vBoxLayout.setContentsMargins(36, 20, 36, 36)

        self.view.setObjectName("view")
        StyleSheet.GALLERY_INTERFACE.apply(self)

    def scrollToCard(self, index: int):
        """scroll to example card"""
        w = self.vBoxLayout.itemAt(index).widget()
        self.verticalScrollBar().setValue(w.y())

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.toolBar.resize(self.width(), self.toolBar.height())
