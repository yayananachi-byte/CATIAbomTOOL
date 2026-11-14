# -*- coding: utf-8 -*-
"""吉利数据导出时对总成、分总成和零件名称的专用处理。"""


def part_name_without_version(name: str) -> str:
    """名称中的第一个 / 及其后内容属于版本信息，原始 BOM 只保留名称。"""
    return name.split("/", 1)[0].strip()
