from pathlib import Path

import openpyxl
import pandas as pd
import shortuuid
from loguru import logger
from PySide6.QtCore import Signal


class Save:
    logInfo = Signal(str)

    def __init__(self, filename: Path):
        self.filename = filename

    def to_excel(self, datas: list, tag: str):
        df = None

        # 判断文件是否存在, 不存在则新建
        if not self.filename.exists():
            workBook = openpyxl.Workbook()  # 创建一个工作簿对象
        else:
            workBook = openpyxl.load_workbook(
                self.filename, keep_vba=True
            )  # 打开 Excel 表格格

            df = pd.read_excel(self.filename)

        sheet = workBook.active  # 选取第一个sheet

        if datas is None or len(datas) == 0:
            return

        max_row = sheet.max_row
        i = 1

        # 表头
        headers = [
            "uuid",
            "药店名称",
            "店铺主页",
            "资质名称",
            "营业执照图片",
            "药品名",
            "药品图片",
            "原价",
            "挂网价格",
            "平台",
            "排查日期",
        ]

        # 如果是第一次保存数据, 就添加表头
        if max_row == 1:
            sheet.append(headers)

        save_flag: bool = True

        for data in datas:
            # 重复数据不保存 - 根据药店名称、店铺主页、药品名、挂网价格、平台判断
            if df is not None:
                temp_df = df[
                    (df["药店名称"] == data[1])
                    & (df["店铺主页"] == data[2])
                    & (df["药品名"] == data[5])
                    &
                    # (df["药品图片"] == data[6]) &
                    (df["挂网价格"] == float(data[8]))
                    & (df["平台"] == data[9])
                ]

                if not temp_df.empty:
                    continue

            # 生成一个短 UUID
            short_uuid = shortuuid.uuid()
            data[0] = short_uuid

            sheet.append(data)
            i += 1

            save_flag = False

        if save_flag:
            msg = f"{self.filename.stem} {tag}-没有数据需要保存"
            self.logInfo.emit(msg)
            logger.info(msg)
            return

        workBook.save(self.filename)

        msg = f"\n\n{self.filename.stem} {tag}-保存了{i - 1}条, 数据总条数: {sheet.max_row - 1}\n\n"
        self.logInfo.emit(msg)
        logger.success(msg)
