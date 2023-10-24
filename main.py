import datetime
import os
import random
import re
import shutil
import time
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

from config import logger, tg_token, chat_id, db_host, robot_name, db_port, db_name, db_user, db_pass, ip_address, saving_path, saving_path_1c, download_path, ecp_paths, main_excel_file, adb_db_password, adb_db_name, adb_db_username, adb_ip, adb_port, mapping_file, filled_file
from core import Odines
from tools.app import App
from tools.web import Web


def sql_create_table():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
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
        '''
    c = conn.cursor()
    c.execute(table_create_query)

    conn.commit()
    c.close()
    conn.close()


def delete_by_id(id):
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
                DELETE FROM ROBOT.{robot_name.replace("-", "_")} WHERE id = '{id}'
                '''
    c = conn.cursor()
    c.execute(table_create_query)
    conn.commit()
    c.close()
    conn.close()


def get_all_data():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            order by started_time asc
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())
    df1.columns = ['started_time', 'ended_time', 'full_name', 'short_name', 'executor_name', 'status', 'status_1c', 'error_reason', 'error_saved_path', 'execution_time', 'ecp_path']

    cur.close()
    conn.close()

    return df1


def get_data_by_name(store_name):
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            where store_name = '{store_name}'
            order by started_time desc
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())
    # df1.columns = ['started_time', 'ended_time', 'store_id', 'name', 'status', 'error_reason', 'error_saved_path', 'execution_time']

    cur.close()
    conn.close()

    return len(df1)


def get_data_to_execute():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            where (status_1c != 'success' and status_1c != 'processing')
            and (executor_name is NULL or executor_name = '{ip_address}')
            order by started_time desc
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())

    with suppress(Exception):
        df1.columns = ['started_time', 'ended_time', 'full_name', 'short_name', 'executor_name', 'status', 'status_1c', 'error_reason', 'error_saved_path', 'execution_time', 'ecp_path']

    cur.close()
    conn.close()

    return df1


def insert_data_in_db(started_time: str, store_name: str, short_name: str, executor_name: str, status_: str, status_1c: str, error_reason: str, error_saved_path: str, execution_time: int, ecp_path_: str):

    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)

    print('Started inserting')
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
        ecp_path_
    )

    print(values)

    cursor = conn.cursor()

    cursor.execute(query_delete)
    # conn.autocommit = True
    try:
        cursor.execute(query_delete)
        # cursor.execute(query_delete_id)
    except Exception as e:
        print('GOVNO', e)
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

    conn = psycopg2.connect(dbname=adb_db_name, host=adb_ip, port=adb_port,
                            user=adb_db_username, password=adb_db_password)

    cur = conn.cursor(name='1583_first_part')

    query = f"""
        select db.id_sale_object, ds.source_store_id, ds.store_name, ds.sale_obj_name 
        from dwh_data.dim_branches db
        left join dwh_data.dim_store ds on db.id_sale_object = ds.sale_source_obj_id
        where ds.store_name like '%Торговый%' and current_date between ds.datestart and ds.dateend
        group by db.id_sale_object, ds.source_store_id, ds.store_name, ds.sale_obj_name
        order by ds.source_store_id
    """

    cur.execute(query)

    print('Executed')

    df1 = pd.DataFrame(cur.fetchall())
    df1.columns = ['branch_id', 'store_id', 'store_name', 'store_normal_name']

    cur.close()
    conn.close()

    return df1


def sign_ecp(ecp):
    logger.info('Started ECP')
    logger.info(f'KEY: {ecp}')
    app = App('')

    el = {"title": "Открыть файл", "class_name": "SunAwtDialog", "control_type": "Window",
          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None}

    if app.wait_element(el, timeout=30):

        keyboard.send_keys(ecp.replace('(', '{(}').replace(')', '{)}'), pause=0.01, with_spaces=True)
        sleep(0.05)
        keyboard.send_keys('{ENTER}')

        if app.wait_element({"title_re": "Формирование ЭЦП.*", "class_name": "SunAwtDialog", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None}, timeout=30):
            app.find_element({"title_re": "Формирование ЭЦП.*", "class_name": "SunAwtDialog", "control_type": "Window",
                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None}).type_keys('Aa123456')

            sleep(2)
            keyboard.send_keys('{ENTER}')
            sleep(3)
            keyboard.send_keys('{ENTER}')
            app = None
            logger.info('Finished ECP')

        else:
            logger.info('Quit mazafaka1')
            app = None
            return 'broke'
    else:
        logger.info('Quit mazafaka')
        app = None
        return 'broke'


def save_screenshot(store):
    scr = screenshot()
    save_path = os.path.join(saving_path, 'Ошибки 1Т')
    scr_path = str(os.path.join(os.path.join(saving_path, 'Ошибки 1Т'), str(store + '.png')))
    scr.save(scr_path)

    return scr_path


def wait_loading(web, xpath):
    print('Started loading')
    ind = 0
    element = ''
    while True:
        try:
            print(web.get_element_display('//*[@id="loadmask-1315"]'))
            if web.get_element_display('//*[@id="loadmask-1315"]') == '':
                element = ''
            if (element == '' and web.get_element_display('//*[@id="loadmask-1315"]') == 'none') or (ind >= 500):
                print('Loaded')
                sleep(0.5)
                break
        except:
            print('No loader')
            break
        ind += 1
        sleep(0.05)


def send_file_to_tg(tg_token, chat_id, param, param1):
    pass


def create_and_send_final_report():
    df = get_all_data()

    df.columns = ['Время начала', 'Время окончания', 'Название филиала', 'Короткое название', 'Статус', 'Причина ошибки', 'Пусть сохранения скриншота', 'Время исполнения (сек)', 'Факт1', 'Факт2', 'Факт3', 'Сайт1', 'Сайт2', 'Сайт3']

    df['Время исполнения (сек)'] = df['Время исполнения (сек)'].astype(float)
    df['Время исполнения (сек)'] = df['Время исполнения (сек)'].round()

    df.to_excel('result.xlsx', index=False)

    workbook = load_workbook('result.xlsx')
    sheet = workbook.active

    red_fill = PatternFill(start_color="FFA864", end_color="FFA864", fill_type="solid")
    green_fill = PatternFill(start_color="A6FF64", end_color="A6FF64", fill_type="solid")

    for cell in sheet['D']:
        if cell.value == 'failed':
            cell.fill = red_fill
        if cell.value == 'success':
            cell.fill = green_fill

    for col in 'ABCDGH':

        max_length = max(len(str(cell.value)) for cell in sheet[col])

        if col == 'A' or col == 'B':
            max_length -= 3
        if col == 'D':
            max_length += 5
        if col == 'A':
            max_length -= 3

        sheet.column_dimensions[col].width = max_length

    for col in 'ABCDGEFGH':
        for cell in sheet[col]:
            cell.alignment = Alignment(horizontal='center')

    workbook.save('result.xlsx')

    send_file_to_tg(tg_token, chat_id, 'Отправляем отчёт по заполнению', 'result.xlsx')


def wait_image_loaded(name):

    found = False
    while True:
        for file in os.listdir(download_path):
            if '.jpg' in file and 'crdownload' not in file:
                shutil.move(os.path.join(download_path, file), os.path.join(os.path.join(saving_path, 'Отчёты 1Т'), name + '.jpg'))
                print(file)
                found = True
                break
        if found:
            break


def save_and_send(web, save):

    print('Saving and Sending')
    if save:
        web.execute_script_click_xpath("//span[text() = 'Сохранить']")
        sleep(1)
        print('Clicked Save')
        if web.wait_element("//span[text() = 'Сохранить отчет и Удалить другие']", timeout=5):
            web.execute_script_click_xpath("//span[text() = 'Сохранить отчет и Удалить другие']")

    print('Clicking Send')
    errors_count = web.find_elements('//*[@id="statflc"]/ul/li/a')
    if len(errors_count) <= 1:
        print('ALL GOOD')
        # web.execute_script_click_xpath("//span[text() = 'Отправить']")
        # print('Clicked Send')
        # web.wait_element("//input[@value = 'Персональный компьютер']", timeout=30)
        # web.execute_script_click_xpath("//input[@value = 'Персональный компьютер']")
    else:
        print('GOVNO OSHIBKA VYLEZLA')


def start_single_branch(branch_name: str,  values_first_part: dict, values_second_part: dict):

    def pass_later():
        if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=1.5):
            web.execute_script_click_xpath("//span[contains(text(), 'Пройти позже')]")

    print('Started web')

    web = Web()
    web.run()
    web.get('https://cabinet.stat.gov.kz/')

    web.wait_element('//*[@id="idLogin"]')
    web.find_element('//*[@id="idLogin"]').click()

    web.wait_element('//*[@id="button-1077-btnEl"]')
    web.find_element('//*[@id="button-1077-btnEl"]').click()

    web.wait_element('//*[@id="lawAlertCheck"]')
    web.find_element('//*[@id="lawAlertCheck"]').click()

    sleep(0.5)
    web.find_element('//*[@id="loginButton"]').click()

    ecp_auth = ''
    ecp_sign = ''
    for files in os.listdir(os.path.join(ecp_paths, branch_name)):
        if 'AUTH' in files:
            ecp_auth = os.path.join(os.path.join(ecp_paths, branch_name), files)
        if 'GOST' in files:
            ecp_sign = os.path.join(os.path.join(ecp_paths, branch_name), files)

    sleep(1)
    sign_ecp(ecp_auth)

    logged_in = web.wait_element('//*[@id="idLogout"]/a', timeout=60)
    # sleep(1000)
    if logged_in:
        if web.find_element("//a[text() = 'Выйти']"):

            pass_later()

            if web.wait_element('//*[@id="dontAgreeId-inputEl"]', timeout=5):
                web.find_element('//*[@id="dontAgreeId-inputEl"]').click()
                sleep(0.3)
                web.find_element('//*[@id="saveId-btnIconEl"]').click()
                sleep(1)
                web.find_element('//*[@id="ext-gen1893"]').click()
                web.find_element('//*[@id="boundlist-1327-listEl"]/ul/li').click()

                sleep(1)
                web.find_element('//*[@id="button-1326-btnIconEl"]').click()
                print('Done lol')
                sign_ecp(ecp_sign)

                try:
                    pass_later()

                except:
                    pass

            web.wait_element('//*[@id="tab-1168-btnInnerEl"]')
            web.find_element('//*[@id="tab-1168-btnInnerEl"]').click()

            # sleep(0.7)

            web.wait_element('//*[@id="radio-1131-boxLabelEl"]')

            pass_later()

            # ? Check if 1Т exists

            pass_later()

            # * ------- Uncomment -------
            # for _ in range(5):
            #
            #     is_loaded = True if len(web.find_elements("//div[contains(@class, 'x-grid-row-expander')]", timeout=15)) >= 1 else False
            #
            #     if is_loaded:
            #         if web.wait_element("//div[contains(text(), '1-Т')]", timeout=3):
            #             web.find_element("//div[contains(text(), '1-Т')]").click()
            #
            #         else:
            #             saved_path = save_screenshot(branch_name)
            #             web.close()
            #             web.quit()
            #
            #             print('Return those shit')
            #             return ['failed', saved_path, 'Нет 1-Т']
            #
            #     else:
            #         web.refresh()
            #
            # if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=1.5):
            #     web.execute_script_click_xpath("//span[contains(text(), 'Пройти позже')]")
            #
            # sleep(0.5)
            #
            # web.find_element('//*[@id="createReportId-btnIconEl"]').click()
            #
            # sleep(1)
            #
            # # ? Switch to the second window
            print('kek1')
            sleep(7)
            print('switched')
            web.driver.switch_to.window(web.driver.window_handles[-1])

            web.find_element('/html/body/div[1]').click()
            web.wait_element('//*[@id="td_select_period_level_1"]/span')
            web.execute_script_click_js("#btn-opendata")
            sleep(0.3)

            if web.get_element_display('/html/body/div[7]') == 'block':

                web.find_element('/html/body/div[7]/div[11]/div/button[2]').click()

                saved_path = save_screenshot(branch_name)
                web.close()
                web.quit()

                print('Return that shit')
                return ['failed', saved_path, 'Выскочила ошиПочка']

            web.wait_element('//*[@id="sel_statcode_accord"]/div/p/b[1]', timeout=100)
            web.execute_script_click_js("body > div:nth-child(16) > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button:nth-child(1) > span")

            web.wait_element('//*[@id="sel_rep_accord"]/h3[1]/a')

            # ? Open new report to fill it

            print('Clicking1')
            # web.execute_script_click_js("body > div:nth-child(18) > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button:nth-child(1)")
            web.execute_script_click_xpath('/html/body/div[17]/div[11]/div/button[1]/span')

            # ? First page
            print('kek')
            web.wait_element("//a[contains(text(), 'Страница 1')]", timeout=10)
            web.find_element("//a[contains(text(), 'Страница 1')]").click()
            print()

            for row, values in values_first_part.items():

                if values[0] is not None:
                    web.find_element(f'(//*[@id="{row}"]/td[3])[1]').click()
                    web.find_element(f'(//*[@id="{row}_col_2"])').type_keys(values[0])

                if values[1] is not None:
                    web.find_element(f'(//*[@id="{row}"]/td[4])[1]').click()
                    web.find_element(f'(//*[@id="{row}_col_3"])').type_keys(values[1])

            keyboard.send_keys('{TAB}')
            # sleep(100)
            # ? Second page
            web.wait_element("//a[contains(text(), 'Страница 2')]", timeout=10)
            web.find_element("//a[contains(text(), 'Страница 2')]").click()

            web.find_element('//*[@id="rtime"]').select('1')
            sleep(1)
            print('-----')

            for row, values in values_second_part.items():

                if values[0] is not None:
                    web.find_element(f'(//*[@id="{row}"]/td[3])[2]').click()
                    web.find_element(f'(//*[@id="{row}_col_2"])').type_keys(values[0])

                if values[1] is not None:
                    web.find_element(f'(//*[@id="{row}"]/td[4])[2]').click()
                    web.find_element(f'(//*[@id="{row}_col_3"])').type_keys(values[1])

            keyboard.send_keys('{TAB}')
            # ? Last page
            web.find_element("//a[contains(text(), 'Данные исполнителя')]").click()
            web.execute_script(element_type="value", xpath="//*[@id='inpelem_3_0']", value='Естаева Акбота Канатовна')
            web.execute_script(element_type="value", xpath="//*[@id='inpelem_3_1']", value='7273391350')
            web.execute_script(element_type="value", xpath="//*[@id='inpelem_3_2']", value='7073882688')
            web.execute_script(element_type="value", xpath="//*[@id='inpelem_3_3']", value='Yestayeva@magnum.kz')

            sleep(3000)
            # save_and_send(web, save=True)
            sleep(3000)
            # sign_ecp(ecp_sign)
            # sleep(1000)

            # wait_image_loaded(branch_name)

            web.close()
            web.quit()

            print('Successed')
            return ['success', '', '']

            # return ['success', '', sites]

    else:

        saved_path = save_screenshot(branch_name)

        web.close()
        web.quit()

        print('Srok istek')
        return ['failed', saved_path, 'Срок ЭЦП истёк']


def get_all_branches():
    conn = psycopg2.connect(host=adb_ip, port=adb_port, database=adb_db_name, user=adb_db_username, password=adb_db_password)
    table_create_query = f'''
                select name_1c_zup from dwh_data.dim_store
                where current_date between datestart and dateend
                group by name_1c_zup
                '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    return pd.DataFrame(cur.fetchall())


def replacements(ind, line):

    if ind == 0:
        return line.replace('С', 'C')
    if ind == 1:
        return line.replace('ТОО ', 'ТОО')
    if ind == 2:
        return line.replace('г. ', 'г.')


def get_store_name(branch_1c):
    conn = psycopg2.connect(host=adb_ip, port=adb_port, database=adb_db_name, user=adb_db_username, password=adb_db_password)
    table_create_query = f'''
            select distinct(store_name) from dwh_data.dim_store where store_name like '%Торговый%'and name_1c_zup = '{branch_1c}'
            and current_date between datestart and dateend
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df_ = pd.DataFrame(cur.fetchall())

    return df_[df_.columns[0]].iloc[0]


def get_single_report(short_name_: str, name_up: str, name_down: str):

    try:

        app = Odines()
        app.run()

        print('started navigating')

        app.navigate("Файл", "Открыть...")

        print('navigated')

        app1 = App('')
        app1.wait_element({"title": "Открытие", "class_name": "#32770", "control_type": "Window",
                           "visible_only": True, "enabled_only": True, "found_index": 0})

        app1.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
                           "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        app1.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
                           "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(r'C:\Users\Abdykarim.D\Documents\РегламентированныйОтчетФорма1ТКвартальная_на тест.erf', app.keys.ENTER)

        app.parent_switch({"title": "", "class_name": "", "control_type": "Pane", "visible_only": True, "enabled_only": True, "found_index": 20})

        first_input = app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                        "visible_only": True, "enabled_only": True, "found_index": 0})

        second_input = app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                         "visible_only": True, "enabled_only": True, "found_index": 1})

        first_input.click()

        sleep(.1)

        keyboard.send_keys("%+r")

        app.parent_switch(app.root)

        # ? ---
        # app.parent_switch({"title": "", "class_name": "", "control_type": "Pane", "visible_only": True, "enabled_only": True, "found_index": 25})

        # arrow_right = app.find_element({"title": "", "class_name": "", "control_type": "Button",
        #                                 "visible_only": True, "enabled_only": True, "found_index": 3}, timeout=1)

        # all_branches = get_all_branches()

        print(name_down, '|||', short_name_)

        print(name_up, name_down, sep=' | ', end='')

        app.find_element({"title": "", "class_name": "", "control_type": "Button",
                          "visible_only": True, "enabled_only": True, "found_index": 2}, timeout=1).click()

        first_input.click()
        first_input.type_keys("^a")
        first_input.type_keys("{BACKSPACE}")
        first_input.type_keys(name_up.strip(), protect_first=True)

        second_input.click()

        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1.5):
            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            print(f' - BAD', end='')

        else:
            print(f' - GOOD', end='')

        second_input.type_keys("^a")
        second_input.type_keys("{BACKSPACE}")
        second_input.type_keys(name_down.strip(), protect_first=True)

        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1.5):
            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            print(f' - BAD",')

        else:
            print(f' - GOOD",')
        print()
        return ['sucess', '', '']
        # app.find_element({"title": "ОК", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        #
        # app.find_element({"title": "Заполнить", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        #
        # checker = False
        #
        # for _ in range(1000):
        #     with suppress(Exception):
        #         app.find_element({"title": "", "class_name": "", "control_type": "DataGrid", "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=10).click()
        #         checker = True
        #         break
        #     sleep(10)
        #
        # if checker:
        #     app.navigate("Файл", "Сохранить как...")
        #
        #     app.wait_element({"title": "Сохранение", "class_name": "#32770", "control_type": "Window",
        #                       "visible_only": True, "enabled_only": True, "found_index": 0})
        #
        #     app.find_element({"title": "Тип файла:", "class_name": "AppControlHost", "control_type": "ComboBox",
        #                       "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        #
        #     app.find_element({"title": "Лист Excel2007-... (*.xlsx)", "class_name": "", "control_type": "ListItem",
        #                       "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        #
        #     app.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
        #                       "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        #
        #     app.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
        #                       "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(os.path.join(saving_path_1c, short_name_), app.keys.ENTER)
        #
        #     return ['success', os.path.join(saving_path_1c, short_name_ + '.xlsx), '']

    except Exception as err:
        return ['error', '', 'err']


def edit_main_excel_file(number: str, filepath: str):

    os.system('taskkill /im excel.exe /f')

    main_excel = xw.Book(main_excel_file, corrupt_load=True)

    needed_sheet_name = None

    for sheet_name in main_excel.sheets:
        if number == sheet_name.name.split()[1]:
            needed_sheet_name = sheet_name
            break

    main_sheet = main_excel.sheets[needed_sheet_name]

    quarter = None

    for col in 'BDFH':
        if main_sheet[f'{col}4'].value is None and main_sheet[f'{col}21'].value is None:
            quarter = col
            break

    print('col:', quarter)

    workers_end_of_period_main = int(main_sheet[f'{col}31'].value)

    excel_app = xw.App(visible=False)
    excel_app.books.open(filepath, corrupt_load=True)

    app = xw.apps.active
    # branch_excel = xlrd.open_workbook(filepath)
    print(app.range('AJ155').value)

    values = {
        '4': 'AJ92',
        '5': 'AJ96',
        '6': 'AJ98',
        '7': 'AJ102',
        '8': 'AJ104',
        '9': 'AJ110',
        '12': 'AJ116',
        '13': 'AJ118',
        '14': 'AJ120'
    }

    workers_hired: int = int(app.range('AJ132').value)
    workers_end_of_period: int = int(app.range('AJ155').value)
    workers_fired = int(sum([s for s in app.range('AJ138:AJ153').value if s is not None]))
    print(workers_fired)
    print(workers_hired, workers_end_of_period, workers_end_of_period_main, workers_fired)
    if workers_end_of_period != (workers_end_of_period_main + workers_hired - workers_fired):
        print(int(workers_end_of_period), int((workers_end_of_period_main + workers_hired - workers_fired)))
        workers_hired += int(workers_end_of_period) - int((workers_end_of_period_main + workers_hired - workers_fired))

    print(workers_hired, workers_end_of_period, workers_end_of_period_main, workers_fired)

    for key, val in values.items():
        if key not in '89':
            main_sheet[f'{quarter}{key}'].value = round(app.range(val).value)
        else:
            main_sheet[f'{quarter}{key}'].value = app.range(val).value

    main_sheet[f'{quarter}21'].value = workers_hired
    main_sheet[f'{quarter}29'].value = workers_fired

    main_excel.app.calculate()
    main_excel.save(filled_file)
    main_excel.close()

    os.system('taskkill /im excel.exe /f')

    return quarter


def open_1c_zup():

    all_branches = pd.read_excel(mapping_file)
    c = 0

    all_excels_ = dict()

    print()

    pattern = r'\d+'

    # ! TODO
    # ! branches_to_execute

    for ind in range(len(all_branches)):

        try:
            short_name_ = get_store_name(all_branches['Низ'].iloc[ind])
        except:
            try:
                short_name_ = get_store_name(all_branches['Низ'].iloc[ind].replace('С', 'C'))
            except:
                print(f"Branch {all_branches['Низ'].iloc[ind]} is dead")
                continue

        branches_to_execute = get_data_to_execute()

        # if short_name_.replace('Торговый зал ', '') not in list(branches_to_execute['short_name']):
        #     print('skipped', short_name_)
        #     continue

        numbers = re.findall(pattern, all_branches['Низ'].iloc[ind])
        # with suppress(Exception):
        #     if numbers[0] == '27':
        #         print(all_branches['Низ'].iloc[ind])

        if 'алмат' in all_branches['Низ'].iloc[ind].lower() and int(numbers[0]) <= 39 and 'центр' not in all_branches['Низ'].iloc[ind].lower():
            print(numbers, '|||', all_branches['Низ'].iloc[ind])
            # continue
            c += 1

            start_time_ = time.time()

            insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=str(all_branches['Низ'].iloc[ind]), short_name=str(short_name_).replace('Торговый зал ', ''),
                              executor_name=str(ip_address), status_='', status_1c='processing', error_reason='', error_saved_path='', execution_time=0, ecp_path_=os.path.join(ecp_paths, all_branches['Низ'].iloc[ind]))
            try:
                status_, filepath, error_ = get_single_report(short_name_, all_branches['Верх'].iloc[ind], all_branches['Низ'].iloc[ind])

                insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=all_branches['Низ'].iloc[ind], short_name=short_name_.replace('Торговый зал ', ''),
                                  executor_name=ip_address, status_='', status_1c=status_, error_reason=error_, error_saved_path='', execution_time=round(time.time() - start_time_), ecp_path_=os.path.join(ecp_paths, all_branches['Низ'].iloc[ind]))

                # filepath = ind
                all_excels_.update({short_name_: [filepath, numbers[0]]})

            except Exception as error__:

                insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=all_branches['Низ'].iloc[ind], short_name=short_name_.replace('Торговый зал ', ''),
                                  executor_name=ip_address, status_='', status_1c='failed 1C', error_reason=str(error__), error_saved_path='', execution_time=round(time.time() - start_time_), ecp_path_=os.path.join(ecp_paths, all_branches['Низ'].iloc[ind]))

    return all_excels_


def get_data_to_fill(short_name_, col_):

    os.system('taskkill /im excel.exe /f')

    main_excel = xw.Book(filled_file, corrupt_load=True) # * os.path.join(saving_path, '1Т Заполненный файл.xlsx')

    needed_sheet_name = None
    print('col:', col_)
    number = short_name_.split('№')[1]

    for sheet_name in main_excel.sheets:
        if number == sheet_name.name.split()[1]:
            needed_sheet_name = sheet_name
            break

    main_sheet = main_excel.sheets[needed_sheet_name]
    print(needed_sheet_name)

    first_part_, second_part_ = dict(), dict()

    first_vals = {
        0: 1,
        1: 2,
        2: 3,
        3: 4,
        4: 5,
        5: 6,
        6: 9,
        7: 10,
        8: 11
    }

    second_vals = {
        0: 2,
        1: 10,
        2: 11,
        3: 12
    }

    for ind_, row in enumerate(['4', '5', '6', '7', '8', '9', '12', '13', '14']):

        if row in '89':

            first_part_.update({first_vals.get(ind_): [main_sheet.range(f'{col_}{row}').value, main_sheet.range(f'{chr(ord(col_) + 1)}{row}').value]})

        else:
            left = round(main_sheet.range(f'{col_}{row}').value) if main_sheet.range(f'{col_}{row}').value is not None else None
            right = round(main_sheet.range(f'{chr(ord(col_) + 1)}{row}').value) if main_sheet.range(f'{chr(ord(col_) + 1)}{row}').value is not None else None

            first_part_.update({first_vals.get(ind_): [left, right]})

    for ind_, row in enumerate(['21', '29', '30', '31']):

        left = round(main_sheet.range(f'{col_}{row}').value) if main_sheet.range(f'{col_}{row}').value is not None else None
        right = round(main_sheet.range(f'{chr(ord(col_) + 1)}{row}').value) if main_sheet.range(f'{chr(ord(col_) + 1)}{row}').value is not None else None

        second_part_.update({second_vals.get(ind_): [left, right]})

    main_excel.close()

    os.system('taskkill /im excel.exe /f')

    return needed_sheet_name, first_part_, second_part_


def dispatcher():

    df = pd.read_excel(main_excel_file)


if __name__ == '__main__':

    sql_create_table()

    dispatcher()

    all_excels = open_1c_zup()
    exit()
    # edit_main_excel_file(r'\\172.16.8.87\d\.rpa\.agent\robot-stat-1t\Output\Выгрузка 1Т из 1С\Торговый зал АФ №21.xlsx', '21')

    col = 'F' # None
    for key, val in all_excels.items():
        col = edit_main_excel_file(key, val[1])

    for branch, vals in all_excels.items():
        if '~' not in vals:
            if '21' not in branch:
                continue
            print(branch, vals)

            short_name, first_part, second_part = get_data_to_fill(branch, col)

            print('-------------')
            print(first_part, second_part, sep='\n')

            start_time = time.time()
            insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch, short_name=short_name,
                              executor_name=ip_address, status_='processing', status_1c='success', error_reason='', error_saved_path='', execution_time=0, ecp_path_=os.path.join(ecp_paths, branch))
            try:
                status, error_saved_path, error = start_single_branch(branch, first_part, second_part)

                insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch, short_name=short_name,
                                  executor_name=ip_address, status_=status, status_1c='success', error_reason=error, error_saved_path=error_saved_path, execution_time=round(time.time() - start_time), ecp_path_=os.path.join(ecp_paths, branch))

            except Exception as error:

                insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch, short_name=short_name,
                                  executor_name=ip_address, status_='failed with error', status_1c='success', error_reason=str(error), error_saved_path='', execution_time=round(time.time() - start_time), ecp_path_=os.path.join(ecp_paths, branch))




