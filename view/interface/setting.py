# coding:utf-8
import platform
import sys

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QLabel, QWidget
from qfluentwidgets import ExpandLayout
from qfluentwidgets import FluentIcon as FIF
from qfluentwidgets import (
    InfoBar,
    InfoBarPosition,
    OptionsSettingCard,
    PrimaryPushSettingCard,
    PushSettingCard,
    ScrollArea,
    SettingCardGroup,
    SwitchSettingCard,
    setTheme,
)

from common.config import AUTHOR, VERSION, YEAR, cfg
from common.style_sheet import StyleSheet

if platform.system() == "Windows":
    import winreg


class SettingInterface(ScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent=parent)

        self.parent = parent

        self.scrollWidget = QWidget()
        self.expandLayout = ExpandLayout(self.scrollWidget)

        self.settingLabel = QLabel("设置", self)

        # 开机自启动
        self.startAtStartGroup = SettingCardGroup("开机自启", self.scrollWidget)

        self.runAtStartCard = SwitchSettingCard(
            FIF.TRANSPARENT,
            "是否开机自启动",
            "开机就运行...",
            cfg.autoStart,
            self.startAtStartGroup,
        )
        self.runAtStartCard.checkedChanged.connect(self.toggle_auto_start)

        # 是否让窗口保持置顶
        self.isStaysOnTopGroup = SettingCardGroup("窗口置顶", self.scrollWidget)
        self.isStaysOnTopCard = SwitchSettingCard(
            FIF.PIN,
            "是否置顶",
            "保持在最前面",
            cfg.staysOnTop,
            self.isStaysOnTopGroup,
        )
        self.isStaysOnTopCard.checkedChanged.connect(self.toggle_stays_on_top)

        if cfg.staysOnTop.value:
            self.parent.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)

        self.personalGroup = SettingCardGroup("个性化", parent=self.scrollWidget)

        self.themeCard = OptionsSettingCard(
            cfg.themeMode,
            FIF.BRUSH,
            "应用主题",
            "调整你的应用外观",
            texts=["浅色", "深色", "跟随系统设置"],
            parent=self.personalGroup,
        )

        self.aboutGroup = SettingCardGroup("关于", parent=self.scrollWidget)

        self.aboutCard = PrimaryPushSettingCard(
            "检查更新",
            FIF.INFO,
            "关于",
            f"© 版权所有 {YEAR}, {AUTHOR}. 当前版本 {VERSION}",
            self.aboutGroup,
        )
        self.aboutCard.clicked.connect(self.createTopRightInfoBar)

        self.__initWidget()

    def __initWidget(self):
        self.resize(1000, 800)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setViewportMargins(0, 80, 0, 20)
        self.setWidget(self.scrollWidget)
        self.setWidgetResizable(True)
        self.setObjectName("settingInterface")

        # 初始化样式表
        self.scrollWidget.setObjectName("scrollWidget")
        self.settingLabel.setObjectName("settingLabel")
        StyleSheet.SETTING_INTERFACE.apply(self)

        # 初始化布局
        self.__initLayout()
        self.__connectSignalToSlot()

    def __initLayout(self):
        self.settingLabel.move(36, 30)

        self.startAtStartGroup.addSettingCard(self.runAtStartCard)
        self.isStaysOnTopGroup.addSettingCard(self.isStaysOnTopCard)
        self.personalGroup.addSettingCard(self.themeCard)
        self.aboutGroup.addSettingCard(self.aboutCard)

        # 把设置卡组添加到布局中
        self.expandLayout.setSpacing(28)
        self.expandLayout.setContentsMargins(36, 10, 36, 0)
        self.expandLayout.addWidget(self.startAtStartGroup)
        self.expandLayout.addWidget(self.isStaysOnTopGroup)
        self.expandLayout.addWidget(self.personalGroup)
        self.expandLayout.addWidget(self.aboutGroup)

    def __onDownloadFolderCardClicked(self):
        # 选择文件夹
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", "../")
        if not folder or cfg.get(cfg.downloadFolder) == folder:
            return

        cfg.set(cfg.downloadFolder, folder)

    def __connectSignalToSlot(self):
        self.themeCard.optionChanged.connect(lambda ci: setTheme(cfg.get(ci)))

    def createTopRightInfoBar(self):
        InfoBar.info(
            title="提示",
            content="此功能在开发中......",
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP_RIGHT,
            duration=2500,
            parent=self,
        )

    def toggle_auto_start(self):
        """
        切换开机自启动
        """
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_path = sys.executable

        if platform.system() != "Windows":
            return

        if self.runAtStartCard.isChecked():
            # 添加开机自启动
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS
            )
            winreg.SetValue(key, "WeChatRobot", winreg.REG_SZ, app_path)
            winreg.CloseKey(key)
        else:
            # 删除开机自启动
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS
            )
            winreg.DeleteValue(key, "WeChatRobot")
            winreg.CloseKey(key)

    def toggle_stays_on_top(self):
        """
        切换窗口置顶
        """
        self.parent.setWindowFlag(
            Qt.WindowType.WindowStaysOnTopHint, self.isStaysOnTopCard.isChecked()
        )

        # 重新显示窗口，确保设置生效
        self.parent.show()
