# coding:utf-8
import sys
from enum import Enum

from PySide6.QtCore import QLocale
from qfluentwidgets import (BoolValidator, ConfigItem, ConfigSerializer,
                            FolderValidator, QConfig, RangeValidator, Theme,
                            qconfig)

from utils.validator import IPValidator


class Language(Enum):
    """Language enumeration"""

    CHINESE_SIMPLIFIED = QLocale(QLocale.Chinese, QLocale.China)
    CHINESE_TRADITIONAL = QLocale(QLocale.Chinese, QLocale.HongKong)
    ENGLISH = QLocale(QLocale.English)
    AUTO = QLocale()


class LanguageSerializer(ConfigSerializer):
    """Language serializer"""

    def serialize(self, language):
        return language.value.name() if language != Language.AUTO else "Auto"

    def deserialize(self, value: str):
        return Language(QLocale(value)) if value != "Auto" else Language.AUTO


def isWin11():
    return sys.platform == "win32" and sys.getwindowsversion().build >= 22000


class Config(QConfig):
    """Config of application"""

    # 通过 MitmProxy 代理搜索
    mitmProxySearch_host = ConfigItem("MitmProxySearch", "Host", "", "")
    mitmProxySearch_excel_path = ConfigItem("MitmProxySearch", "ExcelPath", "", "")
    mitmProxySearch_output_path = ConfigItem(
        "MitmProxySearch", "OutputPath", "", FolderValidator()
    )
    mitmProxySearch_keyword = ConfigItem("MitmProxySearch", "Keyword", "", "")

    # 自启动
    autoStart = ConfigItem("General", "AutoStart", False, BoolValidator())

    # 置顶
    staysOnTop = ConfigItem("General", "StaysOnTop", False, BoolValidator())

    # 京东淘宝自动化
    jdtb_keyword_path = ConfigItem("JdTB", "KeywordPath", "", "")
    jdtb_output_path = ConfigItem("JdTB", "OutputPath", "", FolderValidator())

    # 下载图片
    downloadImg_img_path = ConfigItem("downloadImg", "ImgPath", "", FolderValidator())

    # 图片格式转换
    imgFormatTrans_excel_path = ConfigItem(
        "ImgFormatTrans", "ExcelPath", "", FolderValidator()
    )

    # 通过 yolo 识别药品
    yolo_img_path = ConfigItem("Yolo", "ImagePath", "", "")
    yolo_onnx_path = ConfigItem("Yolo", "OnnxPath", "", "")
    yolo_output_path = ConfigItem("Yolo", "OutputPath", "", "")

    # 删除行
    deleteRow_excel_path = ConfigItem("DeleteRow", "ExcelPath", "", FolderValidator())

    # 从数据库查询资质写入 Excel
    writeExcel_host = ConfigItem("WriteExcel", "Host", "127.0.0.1", "")
    writeExcel_dbname = ConfigItem("WriteExcel", "DbName", "", "")
    writeExcel_user = ConfigItem("WriteExcel", "User", "", "")
    writeExcel_password = ConfigItem("WriteExcel", "Password", "", "")
    writeExcel_excel_path = ConfigItem("WriteExcel", "ExcelPath", "", "")

    # 格式化
    formatExcel_excel_path = ConfigItem("Format", "ExcelPath", "", FolderValidator())

    # 保存 Excel 内容到数据库
    saveToDb_host = ConfigItem("SaveToDb", "Host", "127.0.0.1", "")
    saveToDb_dbname = ConfigItem("SaveToDb", "DbName", "", "")
    saveToDb_user = ConfigItem("SaveToDb", "User", "", "")
    saveToDb_password = ConfigItem("SaveToDb", "Password", "", "")
    saveToDb_excel_path = ConfigItem("SaveToDb", "ExcelPath", "", "")

    # 统计数据
    statistics_excel_path = ConfigItem("Statistics", "ExcelPath", "", FolderValidator())

    # 统计新增加的数据
    incrementalDatas_excel_path1 = ConfigItem("IncrementalDatas", "ExcelPath1", "", "")
    incrementalDatas_excel_path2 = ConfigItem("IncrementalDatas", "ExcelPath2", "", "")
    incrementalDatas_output_path = ConfigItem(
        "IncrementalDatas", "OutputPath", "", FolderValidator()
    )

    # 查找值
    searchval_excel_path = ConfigItem("SearchVal", "ExcelPath", "", FolderValidator())

    # 更新数据库的资质名称
    updateCert_host = ConfigItem("UpdateCert", "Host", "127.0.0.1", "")
    updateCert_dbname = ConfigItem("UpdateCert", "DbName", "", "")
    updateCert_user = ConfigItem("UpdateCert", "User", "", "")
    updateCert_password = ConfigItem("UpdateCert", "Password", "", "")
    updateCert_excel_path = ConfigItem("UpdateCert", "ExcelPath", "", "")

    fiximgsuffix_excel_path = ConfigItem("FixImgSuffix", "ExcelPath", "", "")

    # 合并 Excel 文件
    mergedExcelFiles_excel_path = ConfigItem(
        "MergedExcelFiles", "ExcelPath", "", FolderValidator()
    )
    mergedExcelFiles_output_path = ConfigItem(
        "MergedExcelFiles", "OutputPath", "", FolderValidator()
    )

    # 复查数据
    recheck_excel_path = ConfigItem("ReCheck", "ExcelPath", "")
    recheck_output_path = ConfigItem("ReCheck", "OutputPath", "", FolderValidator())

    # 京东营业执照
    searchJdCert_excel_path = ConfigItem("SearchJdCert", "ExcelPath", "", "")

    # 导出资质空白的行
    exportEmptyRow_excel_path = ConfigItem("ExportEmptyRow", "ExcelPath", "", "")
    exportEmptyRow_output_path = ConfigItem(
        "ExportEmptyRow", "OutputPath", "", FolderValidator()
    )


YEAR = 2024
AUTHOR = "ChaChaL"
VERSION = "1.0.0"

cfg = Config()
cfg.themeMode.value = Theme.AUTO
if hasattr(sys, "_MEIPASS"):
    qconfig.load(f"{sys._MEIPASS}/data/config.json", cfg)
else:
    qconfig.load("./data/config.json", cfg)
