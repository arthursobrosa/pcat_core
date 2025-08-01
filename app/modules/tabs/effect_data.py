from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Optional
from ..utils import load_values


def load_effect_sheet(workbook: Workbook) -> Optional[Worksheet]:
    tab_name = "EFEITO"

    if tab_name not in workbook.sheetnames:
        return None
        
    effect_tab = workbook[tab_name]

    length = _get_length(worksheet=effect_tab)
    tariff_type_info = _load_tariff_type_info(length=length)

    subgroup_info = _load_info(
        worksheet=effect_tab,
        start_index=34
    )

    ra0_info = _load_info(
        worksheet=effect_tab,
        start_index=35
    )

    ra1_info = _load_info(
        worksheet=effect_tab,
        start_index=36
    )

    output_workbook = Workbook()
    new_worksheet = output_workbook.active

    new_worksheet.append([
        "TIPO TARIFA",
        "SUBGRUPO",
        "RA0",
        "RA1"
    ])

    for row in zip(tariff_type_info, subgroup_info, ra0_info, ra1_info):
        new_worksheet.append(row)

    return new_worksheet


def _get_length(worksheet: Worksheet) -> int:
    values = load_values(
        from_worksheet=worksheet,
        index=35,
        direction="column",
        start=2
    )

    return len(values)


def _load_tariff_type_info(length: int) -> list[any]:
    tariff_types = ["TUSD", "TE", "TOTAL"]
    values = []

    for tariff_type in tariff_types:
        for _ in range(length):
            values.append(tariff_type)

    return values


def _load_info(worksheet: Worksheet, start_index: int) -> list[any]:
    all_values = []

    for index in range(3):
        values = load_values(
            from_worksheet=worksheet,
            index=start_index + (5 * index),
            direction="column",
            start=2
        )

        all_values += values

    return all_values