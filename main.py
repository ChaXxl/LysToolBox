import sys

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication
from qfluentwidgets import FluentIcon as FIF
from qfluentwidgets import FluentWindow, NavigationItemPosition

from common import resource
from view.interface.deleterow import DeleteRowInterface
from view.interface.fiximgsuffix import FixImageSuffixInterface
from view.interface.formatExcel import FormatExcelInterface
from view.interface.imagesdownload import ImagesDownloadInterface
from view.interface.imgformattrans import ImgFormatTransInterface
from view.interface.jdtbauto import JdTBbAutoInterface
from view.interface.savetodatabase import SaveToDatabaseInterface
from view.interface.searchval import SearchValInterface
from view.interface.setting import SettingInterface
from view.interface.statistics import StatisticsInterface
from view.interface.incrementaldatas import IncrementalDatasInterface
from view.interface.updatecert import UpdateCertInterface
from view.interface.writeexcel import WriteExcelInterface
from view.interface.yolo import YoloInterface
from view.interface.mergedexcelFiles import MergedExcelFilesInterface


class MainWindow(FluentWindow):
    def __init__(self):
        super().__init__()
        self.initWindow()

        # 创建子项目

        # 京东淘宝自动化
        self.jdtb_interface = JdTBbAutoInterface(self)

        # 图片下载
        self.imgd_interface = ImagesDownloadInterface(self)

        # 图片格式转换
        self.imgFormatTrans_interface = ImgFormatTransInterface(self)

        # 通过 yolo 识别药品
        self.yolo_interface = YoloInterface(self)

        # 删除行
        self.deleteRowInterface = DeleteRowInterface(self)

        # 从数据库查询资质写入 Excel
        self.writeExcelInterface = WriteExcelInterface(self)

        # 格式化
        self.formatExcelInterface = FormatExcelInterface(self)

        # 保存 Excel 内容到数据库
        self.saveToDatabaseInterface = SaveToDatabaseInterface(self)

        # 统计数据
        self.statisticsInterface = StatisticsInterface(self)

        # 统计新增加的数据
        self.incrementalDatasInterface = IncrementalDatasInterface(self)

        # 查找值
        self.searchValInterface = SearchValInterface(self)

        # 导出资质空白的行
        self.exportEmptyRowInterface = None

        # 京东资质查询
        self.jdCertInterface = None

        # 更新数据库的资质名称
        self.updateCertInterface = UpdateCertInterface(self)

        # 修正图片后缀名
        self.fixImageSuffixInterface = FixImageSuffixInterface(self)

        # 合并 Excel 文件
        self.mergedExcelFilesInterface = MergedExcelFilesInterface(self)

        # 设置
        self.settingInterface = SettingInterface(self)

        # 往导航栏添加项目
        self.initNavigation()

    def initWindow(self):
        desktop = QApplication.screens()[0].availableGeometry()
        w, h = desktop.width(), desktop.height()

        self.resize(w * 0.6, h * 0.7)
        self.setMinimumWidth(760)
        self.setWindowIcon(QIcon(":/images/logo.png"))
        self.setWindowTitle("乐药师药品排查工具箱")

        self.move(w // 2 - self.width() // 2, h // 2 - self.height() // 2)
        self.show()
        QApplication.processEvents()

    def initNavigation(self):

        self.addSubInterface(self.jdtb_interface, FIF.PEOPLE, "京东淘宝自动化")
        self.addSubInterface(self.imgd_interface, FIF.SEARCH, "图片下载")
        self.addSubInterface(
            self.imgFormatTrans_interface, FIF.EDUCATION, "图片格式转换"
        )
        self.addSubInterface(self.yolo_interface, FIF.CAMERA, "通过 yolo 识别药品")
        self.addSubInterface(self.deleteRowInterface, FIF.DELETE, "删除行")
        self.addSubInterface(
            self.writeExcelInterface, FIF.IMAGE_EXPORT, "从数据库查询资质写入 Excel"
        )
        self.addSubInterface(self.formatExcelInterface, FIF.CAR, "格式化")
        self.addSubInterface(
            self.saveToDatabaseInterface, FIF.DICTIONARY, "保存 Excel 内容到数据库"
        )
        self.addSubInterface(self.statisticsInterface, FIF.AIRPLANE, "统计数据")
        self.addSubInterface(
            self.incrementalDatasInterface, FIF.AIRPLANE, "统计新增加的数据"
        )
        self.addSubInterface(self.searchValInterface, FIF.SEARCH, "查找值")

        self.addSubInterface(
            self.updateCertInterface, FIF.IMAGE_EXPORT, "更新数据库的资质名称"
        )
        self.addSubInterface(
            self.fixImageSuffixInterface, FIF.SETTING, "修正图片后缀名"
        )

        self.addSubInterface(
            self.mergedExcelFilesInterface, FIF.IMAGE_EXPORT, "合并 Excel 文件"
        )

        self.addSubInterface(
            self.settingInterface, FIF.SETTING, "设置", NavigationItemPosition.BOTTOM
        )

        # 设置导航栏默认展开, 宽度为 200
        self.navigationInterface.setExpandWidth(160)
        self.navigationInterface.expand(useAni=False)


def main():
    app = QApplication(sys.argv)

    w = MainWindow()
    w.show()

    app.exec()


if __name__ == "__main__":
    main()
