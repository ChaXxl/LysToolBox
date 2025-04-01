from pathlib import Path
from typing import List, Any, Optional

import openpyxl
import polars as pl
import shortuuid
from loguru import logger
from PySide6.QtCore import Signal


class Save:
    logInfo = Signal(str)

    def __init__(self, filename: Path):
        self.filename = filename

    def to_excel(self, datas: List[List[Any]], tag: Optional[str] = None) -> None:
        """
        保存数据到Excel文件

        Args:
            datas: 要保存的数据列表
            tag: 标签-指明哪个平台
        """
        # 数据为空则直接返回
        if not datas:
            return

        headers = [
            "uuid",
            "药店名称",
            "店铺主页",
            "资质名称",
            "营业执照图片",
            "药品名",
            "药品ID",
            "药品图片",
            "原价",
            "挂网价格",
            "平台",
            "排查日期",
        ]

        new_data = pl.DataFrame(datas, schema=headers)
        existing_df: Optional[pl.DataFrame] = None

        #  如果文件存在, 读取数据并去重
        if self.filename.exists():
            try:
                existing_df = pl.read_excel(self.filename)
            except Exception as e:
                logger.error(f"读取Excel文件失败: {e}")
                self.logInfo.emit(f"读取Excel文件失败: {e}\n请检查文件格式或路径")
                return

            # 去重
            combined_df = pl.concat([existing_df, new_data], how="vertical")
            combined_df = combined_df.unique(
                subset=["药店名称", "店铺主页", "药品名", "挂网价格", "平台"]
            )

        else:
            combined_df = new_data

        # 生成 UUID（只有新增的数据需要）
        combined_df = combined_df.with_columns(
            combined_df["uuid"]
            .map_elements(
                lambda x: shortuuid.uuid() if x is None or x == "" else x,
                return_dtype=pl.Utf8,
            )
            .alias("uuid")
        )

        # 保存数据到Excel
        try:
            combined_df.write_excel(self.filename)
        except Exception as e:
            logger.error(f"保存数据到Excel失败: {e}")
            self.logInfo.emit(f"保存数据到Excel失败: {e}\n请检查文件格式或路径")
            return

        saved_count = (
            combined_df.shape[0] - existing_df.shape[0]
            if existing_df
            else combined_df.shape[0]
        )

        msg = f"\n\n{self.filename.stem} {tag}-保存了 {saved_count} 条, 数据总条数: {combined_df.shape[0]}\n\n"
        self.logInfo.emit(msg)
