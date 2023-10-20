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

from config import logger, tg_token, chat_id, db_host, robot_name, db_port, db_name, db_user, db_pass, ip_address, saving_path, saving_path_1c, download_path, ecp_paths, main_excel_file, adb_db_password, adb_db_name, adb_db_username, adb_ip, adb_port
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
    df1.columns = ['started_time', 'ended_time', 'full_name', 'executor_name', 'status', 'status_1c', 'error_reason', 'error_saved_path', 'execution_time', 'ecp_path']

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
            where (status != 'success' and status != 'processing')
            and (executor_name is NULL or executor_name = '{ip_address}')
            order by started_time desc
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())

    with suppress(Exception):
        df1.columns = ['started_time', 'ended_time', 'full_name', 'executor_name', 'status', 'status_1c', 'error_reason', 'error_saved_path', 'execution_time', 'ecp_path']

    cur.close()
    conn.close()

    return df1


def insert_data_in_db(started_time, store_name, executor_name, status_, status_1c, error_reason, error_saved_path, execution_time, ecp_path_):
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)

    print('Started inserting')
    # query_delete_id = f"""
    #         delete from ROBOT.{robot_name.replace("-", "_")}_2 where store_id = '{store_id}'
    #     """
    query_delete = f"""
        delete from ROBOT.{robot_name.replace("-", "_")} where store_name = '{store_name}'
    """
    query = f"""
        INSERT INTO ROBOT.{robot_name.replace("-", "_")} (started_time, ended_time, store_name, executor_name, status, status_1c, error_reason, error_saved_path, execution_time, ecp_path)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
    """
    # ended_time = '' if status_ != 'success' else datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f")
    ended_time = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f")
    values = (
        started_time,
        ended_time,
        store_name,
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
    try:
        cursor.execute(query, values)
    except Exception as e:
        conn.rollback()
        print(f"Error: {e}")

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

    df.columns = ['Время начала', 'Время окончания', 'Название филиала', 'Статус', 'Причина ошибки', 'Пусть сохранения скриншота', 'Время исполнения (сек)', 'Факт1', 'Факт2', 'Факт3', 'Сайт1', 'Сайт2', 'Сайт3']

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


def wait_image_loaded():

    found = False
    while True:
        for file in os.listdir(download_path):
            if '.jpg' in file and 'crdownload' not in file:
                shutil.move(os.path.join(download_path, file), os.path.join(os.path.join(saving_path, 'Отчёты 1Т'), branch + '.jpg'))
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


def start_single_branch(branch_name: str, store: str, values_first_part, values_second_part):

    print('Started web')

    web = Web()
    web.run()
    web.get('https://cabinet.stat.gov.kz/')
    logger.info('Check-1')
    web.wait_element('//*[@id="idLogin"]')
    web.find_element('//*[@id="idLogin"]').click()

    web.wait_element('//*[@id="button-1077-btnEl"]')
    web.find_element('//*[@id="button-1077-btnEl"]').click()

    web.wait_element('//*[@id="lawAlertCheck"]')
    web.find_element('//*[@id="lawAlertCheck"]').click()

    sleep(0.5)
    web.find_element('//*[@id="loginButton"]').click()

    logger.info('Check-2')
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

            if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=5):
                try:
                    web.find_element("//span[contains(text(), 'Пройти позже')]").click()
                except:
                    save_screenshot(store)
            logger.info('Check0')
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
                    web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=5)
                    web.find_element("//span[contains(text(), 'Пройти позже')]").click()

                except:
                    pass
            logger.info('Check1')
            web.wait_element('//*[@id="tab-1168-btnInnerEl"]')
            web.find_element('//*[@id="tab-1168-btnInnerEl"]').click()

            # sleep(0.7)

            web.wait_element('//*[@id="radio-1131-boxLabelEl"]')

            if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=1.5):
                web.execute_script_click_xpath("//span[contains(text(), 'Пройти позже')]")
            sleep(1)

            # ? Check if 1Т exists

            if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=1.5):
                web.execute_script_click_xpath("//span[contains(text(), 'Пройти позже')]")

            for _ in range(5):

                is_loaded = True if len(web.find_elements("//div[contains(@class, 'x-grid-row-expander')]", timeout=15)) >= 1 else False

                if is_loaded:
                    if web.wait_element("//div[contains(text(), '1-Т')]", timeout=3):
                        web.find_element("//div[contains(text(), '1-Т')]").click()

                    else:
                        saved_path = save_screenshot(store)
                        web.close()
                        web.quit()

                        print('Return those shit')
                        return ['failed', saved_path, 'Нет 1-Т']

                else:
                    web.refresh()

            if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=1.5):
                web.execute_script_click_xpath("//span[contains(text(), 'Пройти позже')]")
            # web.find_element('//*[@id="radio-1133-boxLabelEl"]').click()
            # wait_loading(web, '//*[@id="loadmask-1315"]')
            # web.refresh()

            sleep(0.5)

            web.find_element('//*[@id="createReportId-btnIconEl"]').click()

            sleep(1)

            # ? Switch to the second window
            web.driver.switch_to.window(web.driver.window_handles[-1])

            web.find_element('/html/body/div[1]').click()
            web.wait_element('//*[@id="td_select_period_level_1"]/span')
            web.execute_script_click_js("#btn-opendata")
            sleep(0.3)

            if web.get_element_display('/html/body/div[7]') == 'block':
                web.find_element('/html/body/div[7]/div[11]/div/button[2]').click()

                saved_path = save_screenshot(store)
                web.close()
                web.quit()

                print('Return that shit')
                return ['failed', saved_path, 'Выскочила ошиПочка']

            logger.info('Check3')
            web.wait_element('//*[@id="sel_statcode_accord"]/div/p/b[1]', timeout=100)
            web.execute_script_click_js("body > div:nth-child(16) > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button:nth-child(1) > span")
            # sleep(10900)
            web.wait_element('//*[@id="sel_rep_accord"]/h3[1]/a')
            logger.info('Check999')
            sites = []

            # ? Open new report to fill it

            print('Clicking1')
            # web.execute_script_click_js("body > div:nth-child(18) > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button:nth-child(1)")
            web.execute_script_click_xpath('/html/body/div[17]/div[11]/div/button[1]/span')

            # ? First page

            web.wait_element("//a[contains(text(), 'Страница 1')]", timeout=10)
            web.find_element("//a[contains(text(), 'Страница 1')]").click()
            print()
            id_ = 3



            keyboard.send_keys('{TAB}')
            # sleep(100)
            # ? Second page
            web.wait_element("//a[contains(text(), 'Страница 2')]", timeout=10)
            web.find_element("//a[contains(text(), 'Страница 2')]").click()

            web.find_element('//*[@id="rtime"]').select('2')
            sleep(1)
            print('-----')

            id_ = 3
            for i in range(len(second)):

                cur_key = list(second.keys())[i]

                if cur_key == 'Всего':
                    continue

                web.find_element(f"//table[@id='tb_p1_e0']//tr[{id_}]/td[@role='gridcell'][2]").click()
                web.wait_element(f"//table[@id='tb_p1_e0']//tr[{id_}]/td[@role='gridcell'][2]//input")
                web.find_element(f"//table[@id='tb_p1_e0']//tr[{id_}]/td[@role='gridcell'][2]//input").type_keys(cur_key, delay=1)
                sleep(1)
                keyboard.send_keys('{ENTER}')
                print(cur_key)

                for ind, val in enumerate(second.get(cur_key)):

                    if val == 0 and ind >= 2:
                        continue
                    else:
                        web.find_element(f"//table[@id='tb_p1_e0']//tr[{id_}]/td[@role='gridcell'][{ind + 4}]").click(double=True)
                        print(second.get(cur_key)[ind])
                        # print(f"//table[@id='tb_p1_e0']//tr[{id_}]/td[@role='gridcell'][{ind + 4}]//input")
                        web.wait_element(f"//table[@id='tb_p1_e0']//tr[{id_}]/td[@role='gridcell'][{ind + 4}]//input")
                        web.find_element(f"//table[@id='tb_p1_e0']//tr[{id_}]/td[@role='gridcell'][{ind + 4}]//input").type_keys(str(second.get(cur_key)[ind]), delay=1)

                id_ += 1

            keyboard.send_keys('{TAB}')
            # ? Last page
            web.find_element("//a[contains(text(), 'Данные исполнителя')]").click()
            web.execute_script(element_type="value", xpath="//*[@id='inpelem_2_0']", value='Нарымбаева Алия')
            web.execute_script(element_type="value", xpath="//*[@id='inpelem_2_1']", value='87717041897')
            web.execute_script(element_type="value", xpath="//*[@id='inpelem_2_2']", value='87717041897')
            web.execute_script(element_type="value", xpath="//*[@id='inpelem_2_3']", value='Narymbayeva@magnum.kz')

            save_and_send(web, save=True)
            sleep(3000)
            # sign_ecp(ecp_sign)
            # sleep(1000)

            # wait_image_loaded()

            web.close()
            web.quit()

            print('Successed')
            return ['success', '', '']

            # return ['success', '', sites]

    else:

        saved_path = save_screenshot(store)

        web.close()
        web.quit()

        print('Srok istek')
        return ['failed', saved_path, 'Срок ЭЦП истёк']


def get_all_branches():
    conn = psycopg2.connect(host='172.16.10.22', port=db_port, database='adb', user='rpa_robot', password='Qaz123123+')
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
    conn = psycopg2.connect(host='172.16.10.22', port=db_port, database='adb', user='rpa_robot', password='Qaz123123+')
    table_create_query = f'''
            select distinct(store_name) from dwh_data.dim_store where store_name like '%Торговый%'and name_1c_zup = '{branch_1c}'
            and current_date between datestart and dateend
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df_ = pd.DataFrame(cur.fetchall())

    return df_[df_.columns[0]].iloc[0]


def get_single_report(name_up: str, name_down: str):

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
        print()
        keyboard.send_keys("%+r")

        app.parent_switch(app.root)

        # ? ---
        # app.parent_switch({"title": "", "class_name": "", "control_type": "Pane", "visible_only": True, "enabled_only": True, "found_index": 25})

        # arrow_right = app.find_element({"title": "", "class_name": "", "control_type": "Button",
        #                                 "visible_only": True, "enabled_only": True, "found_index": 3}, timeout=1)
        print()
        # all_branches = get_all_branches()

        short_name = get_store_name(name_down)

        print(name_down, '|||', short_name)

        print(name_up, name_down, sep=' | ', end='')

        app.find_element({"title": "", "class_name": "", "control_type": "Button",
                          "visible_only": True, "enabled_only": True, "found_index": 2}, timeout=1).click()

        first_input.click()
        first_input.type_keys("^a")
        first_input.type_keys("{BACKSPACE}")
        first_input.type_keys(name_up)

        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1.5):
            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            print(f' - BAD', end='')

        else:
            print(f' - GOOD', end='')

        second_input.click()
        second_input.type_keys("^a")
        second_input.type_keys("{BACKSPACE}")
        second_input.type_keys(name_down)

        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1.5):
            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            print(f' - BAD",')

        else:
            print(f' - GOOD",')
        return ['', '']
        # app.find_element({"title": "ОК", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        #
        # app.find_element({"title": "Заполнить", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        #
        # checker = False
        #
        # for _ in range(100):
        #     with suppress(Exception):
        #         app.find_element({"title": "", "class_name": "", "control_type": "DataGrid", "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=10).click()
        #         checker = True
        #         break
        #     sleep(30)
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
        #                       "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(os.path.join(saving_path_1c, short_name), app.keys.ENTER)
        #
        #     return [os.path.join(saving_path_1c, short_name + '.xlsx), short_name]
    except:
        pass


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
    main_excel.save(os.path.join(saving_path, '1Т Заполненный файл.xlsx'))
    main_excel.close()

    os.system('taskkill /im excel.exe /f')

    return quarter


def open_1c_zup():

    all_branches = pd.read_excel('real_mopping.xlsx')
    c = 0

    all_excels = dict()

    print()

    for ind in range(len(all_branches)):

        pattern = r'\d+'

        numbers = re.findall(pattern, all_branches['Низ'].iloc[ind])
        # with suppress(Exception):
        #     if numbers[0] == '27':
        #         print(all_branches['Низ'].iloc[ind])

        if 'алмат' in all_branches['Низ'].iloc[ind].lower() and int(numbers[0]) <= 39 and 'центр' not in all_branches['Низ'].iloc[ind].lower():
            print(numbers, '|||', all_branches['Низ'].iloc[ind])
            c += 1

            # edit_main_excel_file(r'\\172.16.8.87\d\.rpa\.agent\robot-stat-1t\Output\Выгрузка 1Т из 1С\Торговый зал АФ №21.xlsx', '21')
            # break
            # continue

            # filepath, short_name = get_single_report(all_branches['Верх'].iloc[ind], all_branches['Низ'].iloc[ind])
            short_name = f'Торговый зал АФ №21'
            filepath = ind
            all_excels.update({short_name: [filepath, numbers[0]]})

    return all_excels


def get_data_to_fill(short_name_, col_):

    os.system('taskkill /im excel.exe /f')

    main_excel = xw.Book('temp.xlsx', corrupt_load=True) # * os.path.join(saving_path, '1Т Заполненный файл.xlsx')

    needed_sheet_name = None
    print('col:', col_)
    number = short_name_.split('№')[1]

    for sheet_name in main_excel.sheets:
        if number == sheet_name.name.split()[1]:
            needed_sheet_name = sheet_name
            break

    main_sheet = main_excel.sheets[needed_sheet_name]
    print(needed_sheet_name)
    for row in ['4', '5', '6', '7', '8', '9', '12', '13', '14', '20', '21', '29']:
        if row in '89':
            print(main_sheet.range(f'{col_}{row}').value, main_sheet.range(f'{chr(ord(col_) + 1)}{row}').value)
        else:
            print(round(main_sheet.range(f'{col_}{row}').value), round(main_sheet.range(f'{chr(ord(col_) + 1)}{row}').value))

    main_excel.close()


if __name__ == '__main__':

    sql_create_table()

    all_excels = open_1c_zup()

    # edit_main_excel_file(r'\\172.16.8.87\d\.rpa\.agent\robot-stat-1t\Output\Выгрузка 1Т из 1С\Торговый зал АФ №21.xlsx', '21')

    col = 'F' # None
    # for key, val in all_excels.items():
    #     col = edit_main_excel_file(key, val[1])

    for key, vals in all_excels.items():
        if '~' not in vals:
            if '21' not in key:
                continue
            print(key, vals)

            get_data_to_fill(key, col)

            # start_time = time.time()
            # insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch,
            #                   executor_name=ip_address, status_='processing', error_reason='', error_saved_path='', execution_time='', ecp_path_=os.path.join(ecp_paths, branch_))
            # try:
            #     status, error_saved_path, error = start_single_branch(os.path.join(ecp_paths, branch_), branch_, first, second)
            #
            #     insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch,
            #                       executor_name=ip_address, status_=status, error_reason=error, error_saved_path=error_saved_path, execution_time=round(time.time() - start_time), ecp_path_=os.path.join(ecp_paths, branch_))
            #
            # except Exception as error:
            #
            #     insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch,
            #                       executor_name=ip_address, status_='failed with error', error_reason=str(error), error_saved_path='', execution_time=round(time.time() - start_time), ecp_path_=os.path.join(ecp_paths, branch_))




