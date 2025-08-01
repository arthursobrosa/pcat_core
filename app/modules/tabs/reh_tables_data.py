from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Literal, Optional
from ..utils import load_values, get_rows_and_columns_from


def load_reh_tables_sheet(workbook: Workbook) -> Optional[Worksheet]:
    tab_name = "TABELAS REH"

    if tab_name not in workbook.sheetnames:
        return None

    reh_tables_tab = workbook[tab_name]

    subgroup_info = _get_info_from("SUBGRUPO", worksheet=reh_tables_tab, start_jump=3)
    modality_info = _get_info_from("MODALIDADE", worksheet=reh_tables_tab, start_jump=3)

    acessor_info = _get_extended_info_from(
        column_name="ACESSANTE",
        first_table=False,
        worksheet=reh_tables_tab
    )

    class_info = _get_extended_info_from(
        column_name="CLASSE",
        first_table=True,
        worksheet=reh_tables_tab
    )

    subclass_info = _get_extended_info_from(
        column_name="SUBCLASSE",
        first_table=True,
        worksheet=reh_tables_tab
    )

    post_info = _get_info_from("POSTO", worksheet=reh_tables_tab, start_jump=3)

    tusd_kw_ta_info = _get_tusd_or_te_info(
        tusd_or_te="TUSD",
        unit="R$/kW",
        type="TARIFAS DE APLICAÇÃO",
        worksheet=reh_tables_tab
    )

    tusd_mwh_ta_info = _get_tusd_or_te_info(
        tusd_or_te="TUSD",
        unit="R$/MWh",
        type="TARIFAS DE APLICAÇÃO",
        worksheet=reh_tables_tab
    )

    te_mwh_ta_info = _get_tusd_or_te_info(
        tusd_or_te="TE",
        unit="R$/MWh",
        type="TARIFAS DE APLICAÇÃO",
        worksheet=reh_tables_tab
    )

    tusd_kw_be_info = _get_tusd_or_te_info(
        tusd_or_te="TUSD",
        unit="R$/kW",
        type="BASE ECONÔMICA",
        worksheet=reh_tables_tab
    )

    tusd_mwh_be_info = _get_tusd_or_te_info(
        tusd_or_te="TUSD",
        unit="R$/MWh",
        type="BASE ECONÔMICA",
        worksheet=reh_tables_tab
    )

    te_mwh_be_info = _get_tusd_or_te_info(
        tusd_or_te="TE",
        unit="R$/MWh",
        type="BASE ECONÔMICA",
        worksheet=reh_tables_tab
    )

    output_workbook = Workbook()
    new_worksheet = output_workbook.active

    new_worksheet.append([
        "SUBGRUPO",
        "MODALIDADE",
        "ACESSANTE",
        "CLASSE",
        "SUBCLASSE",
        "Posto Tarifário",
        "TUSD Aplicação R$/kW",
        "TUSD Aplicação R$/MWh",
        "TE Aplicação R$/MWh",
        "TUSD BE R$/kW",
        "TUSD BE R$/MWh",
        "TE BE R$/MWh"
    ])

    for row in zip(subgroup_info, modality_info, acessor_info, class_info, subclass_info, post_info, tusd_kw_ta_info, tusd_mwh_ta_info, te_mwh_ta_info, tusd_kw_be_info, tusd_mwh_be_info, te_mwh_be_info):
        new_worksheet.append(row)

    return new_worksheet


def _get_info_from(column_name: str, worksheet: Worksheet, start_jump: int) -> list[any]:
    rows_and_columns = get_rows_and_columns_from(
        value=column_name,
        worksheet=worksheet
    )

    all_values = []

    for row_and_column in rows_and_columns:
        values = load_values(
            from_worksheet=worksheet,
            index=row_and_column[1] - 1,
            direction="column",
            start=row_and_column[0] + start_jump
        )

        all_values += values

    return all_values


def _get_extended_info_from(column_name: str, first_table: bool, worksheet: Worksheet) -> list[any]:
    length = _get_table_length(
        first_table=first_table,
        worksheet=worksheet
    )

    values = []

    if first_table:
        for _ in range(length):
             values.append("Não se aplica")

        values += _get_info_from(column_name, worksheet, start_jump=3)
    else:
        values += _get_info_from(column_name, worksheet, start_jump=3)

        for _ in range(length):
             values.append("Não se aplica")

    return values


def _get_table_length(first_table: bool, worksheet: Worksheet) -> int:
    rows_and_columns = get_rows_and_columns_from(
        value="SUBGRUPO",
        worksheet=worksheet
    )

    if len(rows_and_columns) > 0:
        index = 0 if first_table else 1
        row_and_column = rows_and_columns[index]
    else:
        row_and_column = (2, 1) if first_table else (2, 12)

    values = load_values(
        from_worksheet=worksheet,
        index=row_and_column[1] - 1,
        direction="column",
        start=row_and_column[0] + 3
    )

    return len(values)


def _get_tusd_or_te_info(
    tusd_or_te: Literal["TUSD", "TE"], 
    unit: Literal["R$/kW", "R$/MWh"], 
    type: Literal["TARIFAS DE APLICAÇÃO", "BASE ECONÔMICA"],
    worksheet: Worksheet
) -> list[any]:
    rows_and_columns = get_rows_and_columns_from(
        value=unit, 
        worksheet=worksheet
    )

    all_column_values = []

    for row_and_column in rows_and_columns:
        tusd_or_te_values = load_values(
            from_worksheet=worksheet,
            index=row_and_column[0] - 2,
            direction="row",
            start=row_and_column[1]
        )

        tusd_or_te_upper_value = tusd_or_te_values[0]

        if tusd_or_te_upper_value == tusd_or_te:
            tariff_or_base_values = load_values(
                from_worksheet=worksheet,
                index=row_and_column[0] - 3,
                direction="row",
                start=row_and_column[1]
            )

            tariff_or_base_upper_value = tariff_or_base_values[0]

            if tariff_or_base_upper_value == type:
                column_values = load_values(
                    from_worksheet=worksheet,
                    index=row_and_column[1] - 1,
                    direction="column",
                    start=row_and_column[0] + 1
                )

                all_column_values += column_values

    return all_column_values
