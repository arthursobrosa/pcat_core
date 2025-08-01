import os
from typing import Literal
import difflib
from .utils import get_suffix


def _get_similar_item(
    base_value: str, 
    items: list[str], 
    suffix: str = ".xlsx"
):
    contained_items = []

    for item in items:
        item_suffix = get_suffix(item)
        item_without_suffix = item.replace(item_suffix, "")

        if item_without_suffix in base_value:
            contained_items.append(item_without_suffix)

    if not contained_items:
        return None

    similar_item = max(
        contained_items, 
        key=lambda item: difflib.SequenceMatcher(None, item, base_value).ratio()
    )

    if not similar_item:
        return None
    
    return similar_item + suffix


def _get_date_string(base_value: str, similar_item: str):
    similar_item_suffix = get_suffix(similar_item)
    similar_item_without_suffix = similar_item.replace(similar_item_suffix, "")

    return base_value.removeprefix(similar_item_without_suffix)


def _get_file_with_date(file_path: str, date_string: str):
    suffix = get_suffix(file_path)
    file_path_without_suffix = file_path.replace(suffix, "")

    file_path_with_date = file_path_without_suffix + date_string
    file_path_with_date = file_path_with_date + suffix

    return file_path_with_date


def _replace_files_suffixes(agent: Literal["Concessionária", "Permissionária"]):
    base_path = os.path.join(os.path.dirname(__file__), "../../")
    base_path = os.path.abspath(base_path)

    distributors_path = os.path.join(base_path, f"{agent}s")
    distributors_years_path = os.path.join(base_path, f"{agent}s_anos")

    distributors = [
        name for name in os.listdir(distributors_path)
        if os.path.isdir(os.path.join(distributors_path, name))
    ]

    distributors.sort()

    for distributor in distributors:
        distributor_year_path = os.path.join(distributors_years_path, distributor)
        distributor_path = os.path.join(distributors_path, distributor)

        for type in ["Ajuste EER ANGRA III", "Liminar abrace", "Reajuste", "Revisão", "Revisão Extraordinária", "Tarifas Iniciais"]:
            year_type_path = os.path.join(distributor_year_path, type)
            type_path = os.path.join(distributor_path, type)

            year_file_names = [
                name for name in os.listdir(year_type_path)
                if (name.endswith(".xlsx") or name.endswith(".xlsm")) 
                and not name.startswith("~$")
            ]

            file_names = [
                name for name in os.listdir(type_path)
                if (name.endswith(".xlsx") or name.endswith(".xlsm")) 
                and not name.startswith("~$")
            ]

            file_paths = []

            for file_name in file_names:
                file_path = os.path.join(type_path, file_name)
                file_paths.append(file_path)

            for year_file_name in year_file_names:
                year_file_path = os.path.join(year_type_path, year_file_name)
                file_path = year_file_path.replace("_anos", "")
                file_path_suffix = get_suffix(file_path)
                file_path_without_suffix = file_path.replace(file_path_suffix, "")

                similar_path = _get_similar_item(
                    base_value=file_path_without_suffix,
                    items=file_paths,
                    suffix=file_path_suffix
                )

                if not similar_path:
                    continue

                date_string = _get_date_string(
                    base_value=file_path_without_suffix,
                    similar_item=similar_path
                )

                if not date_string:
                    continue

                similar_path_with_date = _get_file_with_date(
                    file_path=similar_path,
                    date_string=date_string
                )

                if os.path.exists(similar_path_with_date):
                    continue

                os.rename(similar_path, similar_path_with_date)


def replace_all_files_suffixes():
    _replace_files_suffixes("Concessionária")
    _replace_files_suffixes("Permissionária")


def _show_missing_files(agent: Literal["Concessionária", "Permissionária"]):
    base_path = os.path.join(os.path.dirname(__file__), "../../")
    base_path = os.path.abspath(base_path)

    distributors_path = os.path.join(base_path, f"{agent}s")
    distributors_years_path = os.path.join(base_path, f"{agent}s_anos")

    distributors = [
        name for name in os.listdir(distributors_path)
        if os.path.isdir(os.path.join(distributors_path, name))
    ]

    distributors.sort()

    for distributor in distributors:
        distributor_year_path = os.path.join(distributors_years_path, distributor)
        distributor_path = os.path.join(distributors_path, distributor)

        for type in ["Ajuste EER ANGRA III", "Liminar abrace", "Reajuste", "Revisão", "Revisão Extraordinária", "Tarifas Iniciais"]:
            year_type_path = os.path.join(distributor_year_path, type)
            type_path = os.path.join(distributor_path, type)

            year_file_names = [
                name for name in os.listdir(year_type_path)
                if (name.endswith(".xlsx") or name.endswith(".xlsm")) 
                and not name.startswith("~$")
            ]

            file_names = [
                name for name in os.listdir(type_path)
                if (name.endswith(".xlsx") or name.endswith(".xlsm")) 
                and not name.startswith("~$")
            ]

            missing_files = list(set(file_names) - set(year_file_names))

            if missing_files:
                print(f"missing files at {distributor} - {type}: {missing_files}")


def show_all_missing_files():
    _show_missing_files("Concessionária")
    _show_missing_files("Permissionária")