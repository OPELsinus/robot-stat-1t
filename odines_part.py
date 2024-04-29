import datetime
import os
import random
import re
import shutil
import time
import traceback
from contextlib import suppress
from math import floor
from time import sleep
import pandas as pd

import psycopg2
import xlrd
from mouseinfo import screenshot
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from pywinauto import keyboard
import xlwings as xw

from config import (
    logger,
    tg_token,
    chat_id,
    db_host,
    robot_name,
    db_port,
    db_name,
    db_user,
    db_pass,
    ip_address,
    saving_path,
    saving_path_1c,
    download_path,
    ecp_paths,
    main_excel_files,
    adb_db_password,
    adb_db_name,
    adb_db_username,
    adb_ip,
    adb_port,
    mapping_file,
    filled_files,
    reports_saving_path,
    main_executor,
)
from core import Odines
from tools.app import App
from tools.web import Web


def sql_create_table():
    conn = psycopg2.connect(
        host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass
    )
    table_create_query = f"""
        CREATE TABLE IF NOT EXISTS ROBOT.{robot_name.replace("-", "_")} (
            started_time timestamp,
            ended_time timestamp,
            store_name text UNIQUE,
            short_name text,
            executor_name text,
            status text,
            status_1c text,
            error_reason text,
            error_saved_path text,
            execution_time text,
            ecp_path text
            )
        """
    c = conn.cursor()
    c.execute(table_create_query)

    conn.commit()
    c.close()
    conn.close()


def sql_drop_table():
    conn = psycopg2.connect(
        host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass
    )
    table_create_query = f"""
        drop TABLE ROBOT.{robot_name.replace("-", "_")}"""
    c = conn.cursor()
    c.execute(table_create_query)

    conn.commit()
    c.close()
    conn.close()


def delete_by_id(id):
    conn = psycopg2.connect(
        host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass
    )
    table_create_query = f"""
                DELETE FROM ROBOT.{robot_name.replace("-", "_")} WHERE id = '{id}'
                """
    c = conn.cursor()
    c.execute(table_create_query)
    conn.commit()
    c.close()
    conn.close()


def get_all_data():
    conn = psycopg2.connect(
        host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass
    )
    table_create_query = f"""
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            order by started_time asc
            """
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())
    df1.columns = [
        "started_time",
        "ended_time",
        "full_name",
        "short_name",
        "executor_name",
        "status",
        "status_1c",
        "error_reason",
        "error_saved_path",
        "execution_time",
        "ecp_path",
    ]

    cur.close()
    conn.close()

    return df1


def get_data_by_name(store_name):
    conn = psycopg2.connect(
        host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass
    )
    table_create_query = f"""
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            where store_name = '{store_name}'
            order by started_time desc
            """
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())
    # df1.columns = ['started_time', 'ended_time', 'store_id', 'name', 'status', 'error_reason', 'error_saved_path', 'execution_time']

    cur.close()
    conn.close()

    return len(df1)


def get_data_to_execute():
    conn = psycopg2.connect(
        host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass
    )
    table_create_query = f"""
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            where (status_1c != 'success' and status_1c != 'error')
            and (executor_name is NULL or executor_name = '{ip_address}')
            order by RANDOM()
            """
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())

    with suppress(Exception):
        df1.columns = [
            "started_time",
            "ended_time",
            "full_name",
            "short_name",
            "executor_name",
            "status",
            "status_1c",
            "error_reason",
            "error_saved_path",
            "execution_time",
            "ecp_path",
        ]

    cur.close()
    conn.close()

    return df1


def insert_data_in_db(
    started_time: str,
    store_name: str,
    short_name: str,
    executor_name: str or None,
    status_: str,
    status_1c: str,
    error_reason: str,
    error_saved_path: str,
    execution_time: int,
    ecp_path_: str,
):

    conn = psycopg2.connect(
        host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass
    )

    print("Started inserting")
    # query_delete_id = f"""
    #         delete from ROBOT.{robot_name.replace("-", "_")}_2 where store_id = '{store_id}'
    #     """
    query_delete = f"""
        delete from ROBOT.{robot_name.replace("-", "_")} where store_name = '{store_name}'
    """
    query = f"""
        INSERT INTO ROBOT.{robot_name.replace("-", "_")} (started_time, ended_time, store_name, short_name, executor_name, status, status_1c, error_reason, error_saved_path, execution_time, ecp_path)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """
    # ended_time = '' if status_ != 'success' else datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f")
    ended_time = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f")
    values = (
        started_time,
        ended_time,
        store_name,
        short_name,
        executor_name,
        status_,
        status_1c,
        error_reason,
        error_saved_path,
        str(execution_time),
        ecp_path_,
    )

    # print(values)

    cursor = conn.cursor()

    cursor.execute(query_delete)
    # conn.autocommit = True
    try:
        cursor.execute(query_delete)
        # cursor.execute(query_delete_id)
    except Exception as e:
        print("GOVNO", e)
        pass
    if True:
        cursor.execute(query, values)
    # except Exception as e:
    #     conn.rollback()
    #     print(f"Error: {e}")

    conn.commit()

    cursor.close()
    conn.close()


def get_all_branches_with_codes():

    conn = psycopg2.connect(
        dbname=adb_db_name,
        host=adb_ip,
        port=adb_port,
        user=adb_db_username,
        password=adb_db_password,
    )

    cur = conn.cursor(name="1583_first_part")

    query = f"""
        select db.id_sale_object, ds.source_store_id, ds.store_name, ds.sale_obj_name 
        from dwh_data.dim_branches db
        left join dwh_data.dim_store ds on db.id_sale_object = ds.sale_source_obj_id
        where ds.store_name like '%Торговый%' and current_date between ds.datestart and ds.dateend
        group by db.id_sale_object, ds.source_store_id, ds.store_name, ds.sale_obj_name
        order by ds.source_store_id
    """

    cur.execute(query)

    print("Executed")

    df1 = pd.DataFrame(cur.fetchall())
    df1.columns = ["branch_id", "store_id", "store_name", "store_normal_name"]

    cur.close()
    conn.close()

    return df1


def get_all_branches():
    conn = psycopg2.connect(
        host=adb_ip,
        port=adb_port,
        database=adb_db_name,
        user=adb_db_username,
        password=adb_db_password,
    )
    table_create_query = f"""
                select name_1c_zup from dwh_data.dim_store
                where current_date between datestart and dateend
                group by name_1c_zup
                """
    cur = conn.cursor()
    cur.execute(table_create_query)

    return pd.DataFrame(cur.fetchall())


def replacements(ind, line):

    if ind == 0:
        return line.replace("С", "C")
    if ind == 1:
        return line.replace("ТОО ", "ТОО")
    if ind == 2:
        return line.replace("г. ", "г.")


def get_store_name(branch_1c: str):
    conn = psycopg2.connect(
        host=adb_ip,
        port=adb_port,
        database=adb_db_name,
        user=adb_db_username,
        password=adb_db_password,
    )
    table_create_query = f"""
            select distinct(store_name) from dwh_data.dim_store where store_name like '%Торговый%'and name_1c_zup = '{branch_1c.strip()}'
            and current_date between datestart and dateend
            """
    cur = conn.cursor()
    cur.execute(table_create_query)

    df_ = pd.DataFrame(cur.fetchall())

    return df_[df_.columns[0]].iloc[0]


def get_single_report(short_name_: str, name_up: str, name_down: str):

    # return ['success', os.path.join(saving_path_1c, f'{short_name_}.xlsx'), '']

    try:

        # app = Odines()
        app = Odines(
            base="zup_mcc",
            path=r"C:\Program Files\1cv8\8.3.16.1148\bin\1cv8.exe",
        )
        app.auth()
        # app.run()

        print("started navigating")

        app.open("Файл", "Открыть...")

        print("navigated")

        app1 = App("")
        app1.wait_element(
            {
                "title": "Открытие",
                "class_name": "#32770",
                "control_type": "Window",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            }
        )

        app1.find_element(
            {
                "title": "Имя файла:",
                "class_name": "Edit",
                "control_type": "Edit",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            }
        ).click()

        app1.find_element(
            {
                "title": "Имя файла:",
                "class_name": "Edit",
                "control_type": "Edit",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            }
        ).type_keys(
            r"\\172.16.8.87\d\.rpa\.agent\robot-stat-1t\РегламентированныйОтчетФорма1ТКвартальная_на тест.erf",
            app.keys.ENTER,
        )

        if app.wait_element(
            {
                "title": "1С:Предприятие",
                "class_name": "V8NewLocalFrameBaseWnd",
                "control_type": "Window",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            },
            timeout=1,
        ):
            app.find_element(
                {
                    "title": "Да",
                    "class_name": "",
                    "control_type": "Button",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                }
            ).click()

        app.parent_switch(
            {
                "title": "",
                "class_name": "",
                "control_type": "Pane",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 20,
            }
        )

        first_input = app.find_element(
            {
                "title": "",
                "class_name": "",
                "control_type": "Edit",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            }
        )
        try:
            second_input = app.find_element(
                {
                    "title": "",
                    "class_name": "",
                    "control_type": "Edit",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 1,
                }
            )
        except:
            second_input = app.find_element(
                {
                    "title": "",
                    "class_name": "",
                    "control_type": "Edit",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 2,
                }
            )

        first_input.click()

        sleep(0.1)

        keyboard.send_keys("%+r")

        app.parent_switch(app.root)

        # ? ---
        # app.parent_switch({"title": "", "class_name": "", "control_type": "Pane", "visible_only": True, "enabled_only": True, "found_index": 25})

        # arrow_right = app.find_element({"title": "", "class_name": "", "control_type": "Button",
        #                                 "visible_only": True, "enabled_only": True, "found_index": 3}, timeout=1)

        # all_branches = get_all_branches()

        print(name_down, "|||", short_name_)

        print(name_up, name_down, sep=" | ", end="")

        app.find_element(
            {
                "title": "",
                "class_name": "",
                "control_type": "Button",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 2,
            },
            timeout=1,
        ).click()

        first_input.click()
        first_input.type_keys("^a")
        first_input.type_keys("{BACKSPACE}")
        # * chasnge after anually
        first_input.type_keys('ТОО "Magnum Cash&Carry"', protect_first=True)
        # first_input.type_keys(name_up.strip(), protect_first=True)

        second_input.click()

        if app.wait_element(
            {
                "title": "1С:Предприятие",
                "class_name": "V8NewLocalFrameBaseWnd",
                "control_type": "Window",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            },
            timeout=1.5,
        ):
            app.find_element(
                {
                    "title": "Нет",
                    "class_name": "",
                    "control_type": "Button",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                }
            ).click()

            print(f" - BAD", end="")
            logger.info(f"{name_up} | {name_down} - BAD")

        else:
            print(f" - GOOD", end="")
            logger.info(f"{name_up} | {name_down} - GOOD")

        second_input.type_keys("^a")
        second_input.type_keys("{BACKSPACE}")
        second_input.type_keys(name_down.strip(), protect_first=True)

        second_input.type_keys("{ENTER}")
        if app.wait_element(
            {
                "title": "1С:Предприятие",
                "class_name": "V8NewLocalFrameBaseWnd",
                "control_type": "Window",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            },
            timeout=1.5,
        ):
            app.find_element(
                {
                    "title": "Нет",
                    "class_name": "",
                    "control_type": "Button",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                }
            ).click()

            print(f' - BAD",')
            logger.info(f"{name_up} | {name_down} - BAD")

        else:
            print(f' - GOOD",')
            logger.info(f"{name_up} | {name_down} - GOOD")
        print()

        # year_ = int(str(app.find_element({"title_re": ".* г.", "class_name": "", "control_type": "Text",
        #                                   "visible_only": True, "enabled_only": True, "found_index": 0}).element.element_info.rich_text).replace(' г.', ''))
        #
        # if int(datetime.date.today().year) < year_:
        #     app.find_element({"title": "", "class_name": "", "control_type": "Button",
        #                       "visible_only": True, "enabled_only": True, "found_index": 2}).click()
        # if int(datetime.date.today().year) > year_:
        #     app.find_element({"title": "", "class_name": "", "control_type": "Button",
        #                       "visible_only": True, "enabled_only": True, "found_index": 3}).click()

        # return ['sucess', '', '']

        if app.wait_element(
            {
                "class_name": "",
                "control_type": "ListItem",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            },
            timeout=1,
        ):
            els = app.find_elements(
                {
                    "class_name": "",
                    "control_type": "ListItem",
                    "visible_only": True,
                    "enabled_only": True,
                },
                timeout=10,
            )
            for el in els:
                el.click()
                print("CLICKED")
                break
        print("CLICKED OK")
        with suppress(Exception):
            app.find_element(
                {
                    "title": "ОК",
                    "class_name": "",
                    "control_type": "Button",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                },
                timeout=15,
            ).click()
        print("CLICKED FILL")
        app.find_element(
            {
                "title": "Заполнить",
                "class_name": "",
                "control_type": "Button",
                "visible_only": True,
                "enabled_only": True,
                "found_index": 0,
            }
        ).click()
        print("CLICKED AFTER FILL")
        checker = False

        for _ in range(3000):
            with suppress(Exception):
                app.find_element(
                    {
                        "title": "",
                        "class_name": "",
                        "control_type": "DataGrid",
                        "visible_only": True,
                        "enabled_only": True,
                        "found_index": 0,
                    },
                    timeout=10,
                ).click()
                checker = True
                break
            sleep(1)
        print("CHECKER", checker)

        if checker:
            print("OPENING")
            app.open("Файл", "Сохранить как...")
            print("OPENING1")

            app.wait_element(
                {
                    "title": "Сохранение",
                    "class_name": "#32770",
                    "control_type": "Window",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                }
            )

            print("OPENING2")
            app.find_element(
                {
                    "title": "Тип файла:",
                    "class_name": "AppControlHost",
                    "control_type": "ComboBox",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                }
            ).click()

            print("OPENING3")
            app.find_element(
                {
                    "title": "Лист Excel2007-... (*.xlsx)",
                    "class_name": "",
                    "control_type": "ListItem",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                }
            ).click()

            app.find_element(
                {
                    "title": "Имя файла:",
                    "class_name": "Edit",
                    "control_type": "Edit",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                }
            ).click()

            app.find_element(
                {
                    "title": "Имя файла:",
                    "class_name": "Edit",
                    "control_type": "Edit",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                }
            ).type_keys(os.path.join(saving_path_1c, short_name_), app.keys.ENTER)

            doc_already_exists = app.wait_element(
                {
                    "title": "Подтвердить сохранение в виде",
                    "class_name": "#32770",
                    "control_type": "Window",
                    "visible_only": True,
                    "enabled_only": True,
                    "found_index": 0,
                },
                timeout=2,
            )

            if doc_already_exists:
                app.find_element(
                    {
                        "title": "Да",
                        "class_name": "CCPushButton",
                        "control_type": "Button",
                        "visible_only": True,
                        "enabled_only": True,
                        "found_index": 0,
                    }
                ).click()
                time.sleep(0.3)
                if app.wait_element(
                    {
                        "title": "Да",
                        "class_name": "CCPushButton",
                        "control_type": "Button",
                        "visible_only": True,
                        "enabled_only": True,
                        "found_index": 0,
                    },
                    timeout=2,
                ):
                    app.find_element(
                        {
                            "title": "Да",
                            "class_name": "CCPushButton",
                            "control_type": "Button",
                            "visible_only": True,
                            "enabled_only": True,
                            "found_index": 0,
                        }
                    ).click()

            return ["success", os.path.join(saving_path_1c, f"{short_name_}.xlsx"), ""]

    except Exception as err:
        traceback.print_exc()
        return ["error", "", traceback.format_exc()]


def edit_main_excel_file(filepath__: str, branch_name: str, filepath: str, number: str):

    os.system("taskkill /im excel.exe /f")
    print("OPENING EXCEL", filepath__)
    main_excel = xw.Book(filepath__, corrupt_load=True)

    needed_sheet_name = None

    for sheet_name in main_excel.sheets:
        # print(number, sheet_name.name)
        if number.replace("№", "").replace(" ", "") == sheet_name.name.replace(
            "№", ""
        ).replace(" ", ""):
            needed_sheet_name = sheet_name
            break

    print("FOUND SHEET:", branch_name, needed_sheet_name, filepath__, sep=" | ")

    if needed_sheet_name is None:
        main_excel.close()
        return None
    else:
        needed_sheet_name = needed_sheet_name.name
        main_excel.close()

        return needed_sheet_name
        # return 'H'

    # * Comment ELSE in the beggining of filling excels

    main_sheet = main_excel.sheets[needed_sheet_name]

    quarter = None

    var_4 = "4"
    var_21 = "21"
    var_29 = "29"
    var_31 = "31"

    if main_sheet[f"A34"].value is None:
        var_4 = "3"
        var_21 = "20"
        var_29 = "28"
        var_31 = "30"

    for col_ in "BDFH":
        if (
            main_sheet[f"{col_}{var_4}"].value is None
            or main_sheet[f"{col_}{var_21}"].value is None
        ):
            quarter = col_
            break

    if quarter is None:
        main_excel.close()
        return None

    print("col:", quarter)
    print(quarter, var_4, var_21, var_29, var_31)
    print(
        main_sheet[f"{quarter}{var_4}"].value,
        main_sheet[f"{quarter}{var_21}"].value,
        type(main_sheet[f"{quarter}{var_4}"].value),
        type(main_sheet[f"{quarter}{var_21}"].value),
        sep=" | ",
    )

    print(
        f"{main_sheet[f'{quarter}{var_31}'].value} | main_sheet[f'{quarter}{var_31}'].value"
    )

    workers_end_of_period_main = int(main_sheet[f"{quarter}{var_31}"].value)

    excel_app = xw.App(visible=False)
    excel_app.books.open(filepath, corrupt_load=True)

    app = xw.apps.active
    # branch_excel = xlrd.open_workbook(filepath)
    # print(app.range('AJ155').value)

    values = {
        "4": "AJ92",
        "5": "AJ96",
        "6": "AJ98",
        "7": "AJ102",
        "8": "AJ104",
        "9": "AJ110",
        "12": "AJ116",
        "13": "AJ118",
        "14": "AJ120",
    }

    workers_hired: int = (
        int(app.range("AJ132").value) if app.range("AJ132").value is not None else 0
    )
    workers_end_of_period: int = (
        int(app.range("AJ155").value) if app.range("AJ155").value is not None else 0
    )
    workers_fired = int(
        sum([s for s in app.range("AJ138:AJ153").value if s is not None])
    )
    # print(workers_fired)
    # print(workers_hired, workers_end_of_period, workers_end_of_period_main, workers_fired)
    if workers_end_of_period != (
        workers_end_of_period_main + workers_hired - workers_fired
    ):
        # print(int(workers_end_of_period), int((workers_end_of_period_main + workers_hired - workers_fired)))
        workers_hired += int(workers_end_of_period) - int(
            (workers_end_of_period_main + workers_hired - workers_fired)
        )

    # print(workers_hired, workers_end_of_period, workers_end_of_period_main, workers_fired)

    for key_, val in values.items():
        key = int(key_)
        if main_sheet[f"A34"].value is None:
            key -= 1
        if key != 8 or key != 9:
            if app.range(val).value is not None:
                main_sheet[f"{quarter}{key}"].value = round(float(app.range(val).value))
            else:
                main_sheet[f"{quarter}{key}"].value = 0
        else:
            main_sheet[f"{quarter}{key}"].value = app.range(val).value

    main_sheet[f"{quarter}{var_21}"].value = workers_hired
    main_sheet[f"{quarter}{var_29}"].value = workers_fired

    main_excel.app.calculate()
    print(os.path.join(filled_files, filepath__))
    main_excel.save(os.path.join(filled_files, filepath__))
    main_excel.close()

    os.system("taskkill /im excel.exe /f")

    return quarter


def open_1c_zup():

    all_branches = get_data_to_execute()

    all_excels_ = []

    pattern = r"\d+"

    aa = []

    for ind in range(len(all_branches)):
        print(f"Started {all_branches['full_name'].iloc[ind]}", end=" ")
        try:
            short_name_ = get_store_name(all_branches["full_name"].iloc[ind])
        except:
            try:
                short_name_ = get_store_name(
                    all_branches["full_name"].iloc[ind].replace("С", "C")
                )
            except Exception as errorkin:
                traceback.print_exc()
                print(
                    f"Branch {all_branches['full_name'].iloc[ind]} is dead: {errorkin}"
                )
                logger.info(
                    f"{datetime.datetime.now()} | Branch {all_branches['full_name'].iloc[ind]} is dead: {errorkin}"
                )
                continue

        if short_name_ == "Торговый зал СТМ 5АСФ":
            short_name_ = "Торговый зал АСФ №1"

        print(" | ", short_name_)

        found_ = False
        for reports in os.listdir(
            r"\\172.16.8.87\d\.rpa\.agent\robot-stat-1t\Output\Выгрузка 1Т из 1С"
        ):
            if reports.replace(".xlsx", "") == short_name_:
                found_ = True
                break
        if found_:
            continue

        branches_to_execute = get_data_to_execute()

        numbers = re.findall(pattern, all_branches["full_name"].iloc[ind])
        # with suppress(Exception):
        #     if numbers[0] == '27':
        #         print(all_branches['full_name'].iloc[ind])
        if True:

            # * -----
            # checkus = True
            # for file in os.listdir(saving_path_1c):
            #     if short_name_ == file.replace('.xlsx', ''):
            #         # print(f'SKIPPING FILE: {file}')
            #         aa.append(short_name_)
            #         checkus = False
            #         break
            #
            # if checkus:
            #     continue
            # * -----

            start_time_ = time.time()

            insert_data_in_db(
                started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"),
                store_name=str(all_branches["full_name"].iloc[ind]),
                short_name=str(short_name_).replace("Торговый зал ", ""),
                executor_name=ip_address,
                status_="",
                status_1c="processing",
                error_reason="",
                error_saved_path="",
                execution_time=0,
                ecp_path_=os.path.join(ecp_paths, all_branches["full_name"].iloc[ind]),
            )
            if True:
                status_, filepath, error_ = get_single_report(
                    short_name_,
                    all_branches["full_name"].iloc[ind],
                    all_branches["full_name"].iloc[ind],
                )

                insert_data_in_db(
                    started_time=datetime.datetime.now().strftime(
                        "%d.%m.%Y %H:%M:%S.%f"
                    ),
                    store_name=all_branches["full_name"].iloc[ind],
                    short_name=short_name_.replace("Торговый зал ", ""),
                    executor_name=ip_address,
                    status_="",
                    status_1c=status_,
                    error_reason=error_,
                    error_saved_path="",
                    execution_time=round(time.time() - start_time_),
                    ecp_path_=os.path.join(
                        ecp_paths, all_branches["full_name"].iloc[ind]
                    ),
                )

                # filepath = ind
                all_excels_.append(
                    {
                        short_name_: [
                            filepath,
                            short_name_.replace("Торговый зал ", ""),
                            numbers[0],
                        ]
                    }
                )

            # except Exception as error__:
            #     traceback.print_exc()
            #     insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=all_branches['full_name'].iloc[ind], short_name=short_name_.replace('Торговый зал ', ''),
            #                       executor_name=ip_address, status_='', status_1c='failed 1C', error_reason=str(error__), error_saved_path='', execution_time=round(time.time() - start_time_), ecp_path_=os.path.join(ecp_paths, all_branches['full_name'].iloc[ind]))

    print(len(aa))
    print(aa)
    # sleep(10000)

    return all_excels_


def get_all_excels():

    all_branches = get_all_data()

    all_excels_ = []
    all_downloaded_excels = []
    pattern = r"\d+"

    for reports in os.listdir(
        r"\\172.16.8.87\d\.rpa\.agent\robot-stat-1t\Output\Выгрузка 1Т из 1С"
    ):
        all_downloaded_excels.append(reports.replace(".xlsx", ""))

    for ind in range(len(all_branches)):
        print(f"Started {all_branches['full_name'].iloc[ind]}", end=" ")
        try:
            short_name_ = get_store_name(all_branches["full_name"].iloc[ind])
        except:
            try:
                short_name_ = get_store_name(
                    all_branches["full_name"].iloc[ind].replace("С", "C")
                )
            except Exception as errorkin:
                print(
                    f"Branch {all_branches['full_name'].iloc[ind]} is dead: {errorkin}"
                )
                logger.info(
                    f"{datetime.datetime.now()} | Branch {all_branches['full_name'].iloc[ind]} is dead: {errorkin}"
                )
                continue

        if short_name_ == "Торговый зал СТМ 5АСФ":
            short_name_ = "Торговый зал АСФ №1"

        print(" | ", short_name_)

        found_ = True
        for reports in os.listdir(
            r"\\172.16.8.87\d\.rpa\.agent\robot-stat-1t\Output\Выгрузка 1Т из 1С"
        ):
            if reports.replace(".xlsx", "") == short_name_:
                found_ = False
                break
        if found_:
            continue
        numbers = re.findall(pattern, short_name_)
        filepath = os.path.join(saving_path_1c, f"{short_name_}.xlsx")
        print(numbers)
        all_excels_.append(
            {
                short_name_: [
                    filepath,
                    short_name_.replace("Торговый зал ", ""),
                    numbers[0],
                ]
            }
        )

    # sleep(10000)

    return all_excels_


def dispatcher():

    with suppress(Exception):
        sql_drop_table()

    sql_create_table()

    all_branches = pd.read_excel(mapping_file)

    all_branches = all_branches.drop_duplicates()

    for ind in range(len(all_branches)):

        print(f"Started {all_branches['Низ'].iloc[ind]}", end=" ")
        try:
            short_name_ = get_store_name(all_branches["Низ"].iloc[ind])
        except:
            try:
                short_name_ = get_store_name(
                    all_branches["Низ"].iloc[ind].replace("С", "C")
                )
            except:
                try:
                    short_name_ = get_store_name(all_branches["Низ"].iloc[ind].replace("Акмолинской области", "г. Нур-Султан"))
                except Exception as errorkin:
                    traceback.print_exc()
                    print(f"Branch {all_branches['Низ'].iloc[ind]} is dead: {errorkin}")
                    logger.info(
                        f"{datetime.datetime.now()} | Branch {all_branches['Низ'].iloc[ind]} is dead: {errorkin}"
                    )
                    continue
        print()

        insert_data_in_db(
            started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"),
            store_name=str(all_branches["Низ"].iloc[ind]),
            short_name=str(short_name_).replace("Торговый зал ", "").replace('№', '').replace(' ', ''),
            executor_name=None,
            status_="new",
            status_1c="new",
            error_reason="",
            error_saved_path="",
            execution_time=0,
            ecp_path_=os.path.join(ecp_paths, all_branches["Низ"].iloc[ind]),
        )


if __name__ == "__main__":

    try:

        if ip_address == main_executor:

            dispatcher()

        # branches = get_data_to_execute()
        # print(branches)
        #
        open_1c_zup()

    except:
        traceback.print_exc()
