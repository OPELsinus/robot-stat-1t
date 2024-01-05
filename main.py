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

from config import logger, tg_token, chat_id, db_host, robot_name, db_port, db_name, db_user, db_pass, ip_address, saving_path, saving_path_1c, download_path, ecp_paths, main_excel_files, adb_db_password, adb_db_name, adb_db_username, adb_ip, adb_port, mapping_file, filled_files
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

    # print(values)

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


def save_and_send(web, save, ecp_sign):

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
        web.execute_script_click_xpath("//span[text() = 'Отправить']")
        print('Clicked Send')
        if web.wait_element("//input[@value = 'Персональный компьютер']", timeout=60):
            web.execute_script_click_xpath("//input[@value = 'Персональный компьютер']")
        else:
            web.find_element("//button[@class='btn-savesigned ui-button ui-widget ui-state-default ui-corner-all ui-button-text-icon-primary']/span[text() = 'Отправить']").click()
            web.wait_element("//input[@value = 'Персональный компьютер']", timeout=120)
            web.execute_script_click_xpath("//input[@value = 'Персональный компьютер']")
        print('Checkpoint on signing')
        sign_ecp(ecp_sign)

        if web.wait_element("//span[text() = 'Продолжить']", timeout=10):
            web.execute_script_click_xpath("//span[text() = 'Продолжить']")

        if web.wait_element("//h1[contains(text(), 'Whitelabel')]", timeout=5):
            for _ in range(10):
                try:
                    web.execute_script_click_xpath("//span[text() = 'Сохранить']")
                except:
                    sleep(60)
                    web.execute_script_click_xpath("//span[text() = 'Сохранить']")
                sleep(1)
                if web.wait_element("//span[text() = 'Сохранить отчет и Удалить другие']", timeout=30):
                    web.execute_script_click_xpath("//span[text() = 'Сохранить отчет и Удалить другие']")
                web.execute_script_click_xpath("//button[@class='btn-savesigned ui-button ui-widget ui-state-default ui-corner-all ui-button-text-icon-primary']/span[text() = 'Отправить']")

                if web.wait_element("//input[@value = 'Персональный компьютер']", timeout=60):
                    web.execute_script_click_xpath("//input[@value = 'Персональный компьютер']")
                else:
                    web.find_element("//button[@class='btn-savesigned ui-button ui-widget ui-state-default ui-corner-all ui-button-text-icon-primary']/span[text() = 'Отправить']").click()
                    web.wait_element("//input[@value = 'Персональный компьютер']", timeout=120)
                    web.execute_script_click_xpath("//input[@value = 'Персональный компьютер']")
                print('Checkpoint on signing')
                sign_ecp(ecp_sign)

                if web.wait_element("//span[text() = 'Продолжить']", timeout=10):
                    web.execute_script_click_xpath("//span[text() = 'Продолжить']")

                if not web.wait_element("//h1[contains(text(), 'Whitelabel')]", timeout=5):
                    break
    else:
        print('GOVNO OSHIBKA VYLEZLA')


def wait_loading_1t(web, store):

    for i in range(5):

        if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=1.5):
            web.execute_script_click_xpath("//span[contains(text(), 'Пройти позже')]")

        if web.wait_element("(//tr[@role='row'])[1]", timeout=5):

            if web.wait_element("//div[contains(text(), '1-Т (кварт')]", timeout=3):
                web.find_element("//div[contains(text(), '1-Т (кварт')]").click()

                return 0

            else:
                saved_path = save_screenshot(store)
                web.close()
                web.quit()

                print('Return those shit')
                return ['Нет 1-Т', saved_path]

        else:
            web.driver.refresh()

    saved_path = save_screenshot(store)
    web.close()
    web.quit()

    print('Return those shit')
    return ['Нет 1-инвест', saved_path]


def proverka_ecp(web):

    if web.wait_element('//*[@id="AgreeId_header_hd-textEl"]', timeout=.5):
        web.execute_script_click_xpath("//span[text() = 'Согласен']")


def start_single_branch(branch_name: str,  values_first_part: dict, values_second_part: dict):

    def pass_later():
        if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=1.5):
            web.execute_script_click_xpath("//span[contains(text(), 'Пройти позже')]")

    print('Started web')

    ecp_auth = ''
    ecp_sign = ''
    for file in os.listdir(os.path.join(ecp_paths, branch)):

        if 'AUTH' in file:
            ecp_auth = os.path.join(os.path.join(ecp_paths, branch), file)
        if 'GOST' in file:
            ecp_sign = os.path.join(os.path.join(ecp_paths, branch), file)
    print(ecp_auth, '|', ecp_sign)
    web = Web()
    web.run()
    web.get('https://cabinet.stat.gov.kz/')
    logger.info('Check-1')

    logger.info('refreshed')

    proverka_ecp(web=web)

    web.wait_element('//*[@id="idLogin"]')
    web.find_element('//*[@id="idLogin"]').click()

    proverka_ecp(web=web)

    # * --- deprecated (maybe useful in future)
    # web.wait_element('//*[@id="button-1077-btnEl"]')
    # web.find_element('//*[@id="button-1077-btnEl"]').click()
    # * ---
    # proverka_ecp(web=web)
    print()
    # web.wait_element('//*[@id="lawAlertCheck"]')
    # web.find_element('//*[@id="lawAlertCheck"]').click()
    web.execute_script_click_xpath("//input[@id='lawAlertCheck']")

    time.sleep(0.5)
    web.find_element('//*[@id="loginButton"]').click()

    logger.info('Check-2')

    time.sleep(1)

    # send_message_to_tg(tg_token, chat_id, f"Started ECP, {datetime.datetime.now()}")
    sign_ecp(ecp_auth)
    # send_message_to_tg(tg_token, chat_id, f"Finished ECP, {datetime.datetime.now()}")

    logged_in = web.wait_element('//*[@id="idLogout"]/a')

    store = branch.split('\\')[-1]
    # sleep(1000)
    if logged_in:
        if web.find_element("//a[text() = 'Выйти']"):

            if web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=5):
                try:
                    web.find_element("//span[contains(text(), 'Пройти позже')]").click()
                except:
                    save_screenshot(store)
                    # print('HUETA')
                    # sleep(200)
            logger.info('Check0')
            if web.wait_element('//*[@id="dontAgreeId-inputEl"]', timeout=5):
                web.find_element('//*[@id="dontAgreeId-inputEl"]').click()
                sleep(0.3)
                web.find_element('//*[@id="saveId-btnIconEl"]').click()
                sleep(1)

                # * --- Deprecated (maybe useful)
                # web.find_element('//*[@id="ext-gen1893"]').click()
                # web.find_element('//*[@id="boundlist-1327-listEl"]/ul/li').click()
                # * ---

                web.wait_element('//*[@id="keyCombo-inputEl"]')

                web.execute_script_click_xpath("//*[@id='keyCombo-inputEl']/../following-sibling::td//div")

                web.find_element("//li[contains(text(), 'Персональный компьютер')]").click()
                sleep(1.5)

                web.execute_script_click_xpath("//span[contains(text(), 'Продолжить')]")

                print('Done lol')
                sign_ecp(ecp_sign)
                print()
                try:
                    web.wait_element("//span[contains(text(), 'Пройти позже')]", timeout=5)
                    web.find_element("//span[contains(text(), 'Пройти позже')]").click()

                except:
                    pass

            # web.wait_element('//*[@id="radio-1131-boxLabelEl"]')

            pass_later()
            print('OTCHETY')
            web.wait_element("//span[contains(text(), 'Мои отчёты')]")
            web.execute_script_click_xpath("//span[contains(text(), 'Мои отчёты')]")

            # ? Check if 1Т exists

            pass_later()

            # * ------- Uncomment -------
            wait_loading_1t(web, store)
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

            sleep(0.5)

            web.find_element('//*[@id="createReportId-btnIconEl"]').click()

            sleep(1)

            # ? Switch to the second window
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

            # sleep(3000)
            save_and_send(web, save=True, ecp_sign=ecp_sign)
            # sleep(3000)
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

    # return ['success', os.path.join(saving_path_1c, f'{short_name_}.xlsx'), '']

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
                           "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(r'\\172.16.8.87\d\.rpa\.agent\robot-stat-1t\РегламентированныйОтчетФорма1ТКвартальная_на тест.erf', app.keys.ENTER)

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
        # return ['sucess', '', '']
        app.find_element({"title": "ОК", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        app.find_element({"title": "Заполнить", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        checker = False

        for _ in range(1000):
            with suppress(Exception):
                app.find_element({"title": "", "class_name": "", "control_type": "DataGrid", "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=10).click()
                checker = True
                break
            sleep(10)

        if checker:
            app.navigate("Файл", "Сохранить как...")

            app.wait_element({"title": "Сохранение", "class_name": "#32770", "control_type": "Window",
                              "visible_only": True, "enabled_only": True, "found_index": 0})

            app.find_element({"title": "Тип файла:", "class_name": "AppControlHost", "control_type": "ComboBox",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "Лист Excel2007-... (*.xlsx)", "class_name": "", "control_type": "ListItem",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(os.path.join(saving_path_1c, short_name_), app.keys.ENTER)

            return ['success', os.path.join(saving_path_1c, f'{short_name_}.xlsx'), '']

    except Exception as err:
        return ['error', '', str(err)]


def edit_main_excel_file(filepath__: str, branch_name: str, filepath: str, number: str):

    os.system('taskkill /im excel.exe /f')

    main_excel = xw.Book(filepath__, corrupt_load=True)

    needed_sheet_name = None

    for sheet_name in main_excel.sheets:
        print(number, sheet_name.name)
        if number == sheet_name.name.split()[1]:
            needed_sheet_name = sheet_name
            break

    print(needed_sheet_name)

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
    print(os.path.join(filled_files, filepath__))
    main_excel.save(os.path.join(filled_files, filepath__))
    main_excel.close()

    os.system('taskkill /im excel.exe /f')

    return quarter


def open_1c_zup():

    all_branches = pd.read_excel(mapping_file)
    c = 0

    all_excels_ = []

    print()

    pattern = r'\d+'

    # ! TODO
    # ! branches_to_execute

    c = 0

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

        if True:
        # if 'алмат' in all_branches['Низ'].iloc[ind].lower() and int(numbers[0]) <= 39 and 'центр' not in all_branches['Низ'].iloc[ind].lower():
            print(numbers, '|||', all_branches['Низ'].iloc[ind])
            if all_branches['Низ'].iloc[ind] == 'Алматинский филиал №1 ТОО "Magnum Cash&Carry"' or all_branches['Низ'].iloc[ind] == 'Алматинский филиал №10 ТОО "Magnum Cash&Carry"' or all_branches['Низ'].iloc[ind] == 'Алматинский филиал №11 ТОО "Magnum Cash&Carry"'\
               or all_branches['Низ'].iloc[ind] == 'Алматинский филиал №12 ТОО "Magnum Cash&Carry"' or all_branches['Низ'].iloc[ind] == 'Алматинский филиал №13 ТОО "Magnum Cash&Carry"' or all_branches['Низ'].iloc[ind] == 'Алматинский филиал №14 ТОО "Magnum Cash&Carry"'\
               or all_branches['Низ'].iloc[ind] == 'Алматинский филиал №15 ТОО "Magnum Cash&Carry"':
                continue

            # if all_branches['Низ'].iloc[ind] != 'Алматинский филиал №9 ТОО "Magnum Cash&Carry"':
            #     continue

            if 'Астан' in all_branches['Низ'].iloc[ind]:
                continue

            start_time_ = time.time()

            insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=str(all_branches['Низ'].iloc[ind]), short_name=str(short_name_).replace('Торговый зал ', ''),
                              executor_name=str(ip_address), status_='', status_1c='processing', error_reason='', error_saved_path='', execution_time=0, ecp_path_=os.path.join(ecp_paths, all_branches['Низ'].iloc[ind]))
            try:
                status_, filepath, error_ = get_single_report(short_name_, all_branches['Верх'].iloc[ind], all_branches['Низ'].iloc[ind])

                insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=all_branches['Низ'].iloc[ind], short_name=short_name_.replace('Торговый зал ', ''),
                                  executor_name=ip_address, status_='', status_1c=status_, error_reason=error_, error_saved_path='', execution_time=round(time.time() - start_time_), ecp_path_=os.path.join(ecp_paths, all_branches['Низ'].iloc[ind]))

                # filepath = ind
                all_excels_.append({short_name_: [filepath, numbers[0]]})

            except Exception as error__:

                insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=all_branches['Низ'].iloc[ind], short_name=short_name_.replace('Торговый зал ', ''),
                                  executor_name=ip_address, status_='', status_1c='failed 1C', error_reason=str(error__), error_saved_path='', execution_time=round(time.time() - start_time_), ecp_path_=os.path.join(ecp_paths, all_branches['Низ'].iloc[ind]))

    return all_excels_


def get_data_to_fill(filepath: str, short_name_: str, col_: str):

    os.system('taskkill /im excel.exe /f')

    main_excel = xw.Book(filepath, corrupt_load=True) # * os.path.join(saving_path, '1Т Заполненный файл.xlsx')

    needed_sheet_name = None
    print('col:', col_)
    print('short_name:', short_name_)
    number = short_name_.split('№')[1]

    for sheet_name in main_excel.sheets:
        # if number == sheet_name.name.split()[1]:
        if short_name_.replace('Торговый зал ', '').replace(' №', '') == sheet_name.name:
            needed_sheet_name = sheet_name
            break

    if needed_sheet_name is None:
        for sheet_name in main_excel.sheets:
            # if number == sheet_name.name.split()[1]:
            if short_name_.replace('Торговый зал ', '').replace('№', '') == sheet_name.name:
                needed_sheet_name = sheet_name
                break

    print()
    print(short_name_.replace('Торговый зал ', '').replace(' №', ''))
    print(needed_sheet_name.name)
    sheet_name = needed_sheet_name.name
    main_sheet = main_excel.sheets[needed_sheet_name]

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

    return sheet_name, first_part_, second_part_


def dispatcher():

    pass

    # df = pd.read_excel(main_excel_file)


if __name__ == '__main__':

    try:
        sql_create_table()

        dispatcher()

        all_excels = open_1c_zup()

        # all_excels = [{'Торговый зал АФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №1.xlsx', '1'], 'Торговый зал АФ №10': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №10.xlsx', '10'], 'Торговый зал АФ №11': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №11.xlsx', '11'], 'Торговый зал АФ №12': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №12.xlsx', '12'], 'Торговый зал АФ №14': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №14.xlsx', '14'], 'Торговый зал АФ №15': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №15.xlsx', '15'], 'Торговый зал АФ №16': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №16.xlsx', '16'], 'Торговый зал АФ №17': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №17.xlsx', '17'], 'Торговый зал АФ №18': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №18.xlsx', '18'], 'Торговый зал АФ №19': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №19.xlsx', '19'], 'Торговый зал АФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №2.xlsx', '2'], 'Торговый зал АФ №20': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №20.xlsx', '20'], 'Торговый зал АФ №21': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №21.xlsx', '21'], 'Торговый зал АФ №22': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №22.xlsx', '22'], 'Торговый зал АФ №23': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №23.xlsx', '23'], 'Торговый зал АФ №24': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №24.xlsx', '24'], 'Торговый зал АФ №25': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №25.xlsx', '25'], 'Торговый зал АФ №26': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №26.xlsx', '26'], 'Торговый зал АФ №28': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №28.xlsx', '28'], 'Торговый зал АФ №29': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №29.xlsx', '29'], 'Торговый зал АФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №3.xlsx', '3'], 'Торговый зал АФ №30': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №30.xlsx', '30'], 'Торговый зал АФ №31': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №31.xlsx', '31'], 'Торговый зал АФ №32': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №32.xlsx', '32'], 'Торговый зал АФ №33': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №33.xlsx', '33'], 'Торговый зал АФ №34': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №34.xlsx', '34'], 'Торговый зал АФ №35': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №35.xlsx', '35'], 'Торговый зал АФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №6.xlsx', '6'], 'Торговый зал АФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №7.xlsx', '7'], 'Торговый зал АФ №8': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №8.xlsx', '8'], 'Торговый зал АФ №9': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №9.xlsx', '9'], 'Торговый зал КФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №2.xlsx', '2'], 'Торговый зал ЕКФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ЕКФ №1.xlsx', '1'], 'Торговый зал КПФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КПФ №1.xlsx', '1'], 'Торговый зал ФКС №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ФКС №1.xlsx', '1'], 'Торговый зал КЗФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КЗФ №1.xlsx', '1'], 'Торговый зал ТФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТФ №1.xlsx', '1'], 'Торговый зал ППФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №1.xlsx', '1'], 'Торговый зал ТЗФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТЗФ №1.xlsx', '1'], 'Торговый зал УКФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал УКФ №1.xlsx', '1'], 'Торговый зал ППФ №10': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №10.xlsx', '10'], 'Торговый зал ШФ №10': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №10.xlsx', '10'], 'Торговый зал ППФ №11': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №11.xlsx', '11'], 'Торговый зал ППФ №12': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №12.xlsx', '12'], 'Торговый зал ШФ №12': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №12.xlsx', '12'], 'Торговый зал ППФ №13': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №13.xlsx', '13'], 'Торговый зал ШФ №13': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №13.xlsx', '13'], 'Торговый зал АФ №13': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №13.xlsx', '13'], 'Торговый зал ППФ №14': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №14.xlsx', '14'], 'Торговый зал ШФ №14': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №14.xlsx', '14'], 'Торговый зал ППФ №15': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №15.xlsx', '15'], 'Торговый зал ШФ №15': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №15.xlsx', '15'], 'Торговый зал АСФ №16': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №16.xlsx', '16'], 'Торговый зал ППФ №16': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №16.xlsx', '16'], 'Торговый зал ППФ №17': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №17.xlsx', '17'], 'Торговый зал ШФ №17': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №17.xlsx', '17'], 'Торговый зал АСФ №17': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №17.xlsx', '17'], 'Торговый зал АСФ №18': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №18.xlsx', '18'], 'Торговый зал ППФ №18': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №18.xlsx', '18'], 'Торговый зал ШФ №18': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №18.xlsx', '18'], 'Торговый зал АСФ №19': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №19.xlsx', '19'], 'Торговый зал ППФ №19': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №19.xlsx', '19'], 'Торговый зал ШФ №19': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №19.xlsx', '19'], 'Торговый зал ФКС №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ФКС №2.xlsx', '2'], 'Торговый зал КЗФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КЗФ №2.xlsx', '2'], 'Торговый зал ППФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №2.xlsx', '2'], 'Торговый зал ТКФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТКФ №2.xlsx', '2'], 'Торговый зал ТЗФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТЗФ №2.xlsx', '2'], 'Торговый зал ТФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТФ №2.xlsx', '2'], 'Торговый зал АСФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №2.xlsx', '2'], 'Торговый зал УКФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал УКФ №2.xlsx', '2'], 'Торговый зал ШФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №2.xlsx', '2'], 'Торговый зал АСФ №20': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №20.xlsx', '20'], 'Торговый зал ППФ №20': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №20.xlsx', '20'], 'Торговый зал ШФ №20': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №20.xlsx', '20'], 'Торговый зал АСФ №21': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №21.xlsx', '21'], 'Торговый зал ППФ №21': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №21.xlsx', '21'], 'Торговый зал ШФ №21': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №21.xlsx', '21'], 'Торговый зал АСФ №22': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №22.xlsx', '22'], 'Торговый зал ППФ №22': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №22.xlsx', '22'], 'Торговый зал ШФ №22': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №22.xlsx', '22'], 'Торговый зал АСФ №23': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №23.xlsx', '23'], 'Торговый зал ШФ №23': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №23.xlsx', '23'], 'Торговый зал ШФ №24': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №24.xlsx', '24'], 'Торговый зал АСФ №24': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №24.xlsx', '24'], 'Торговый зал АСФ №25': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №25.xlsx', '25'], 'Торговый зал ШФ №25': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №25.xlsx', '25'], 'Торговый зал АСФ №26': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №26.xlsx', '26'], 'Торговый зал ШФ №26': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №26.xlsx', '26'], 'Торговый зал АСФ №27': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №27.xlsx', '27'], 'Торговый зал ШФ №27': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №27.xlsx', '27'], 'Торговый зал ШФ №28': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №28.xlsx', '28'], 'Торговый зал АСФ №28': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №28.xlsx', '28'], 'Торговый зал АСФ №29': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №29.xlsx', '29'], 'Торговый зал ШФ №29': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №29.xlsx', '29'], 'Торговый зал ТФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТФ №3.xlsx', '3'], 'Торговый зал АСФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №3.xlsx', '3'], 'Торговый зал КФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №3.xlsx', '3'], 'Торговый зал ППФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №3.xlsx', '3'], 'Торговый зал ТЗФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТЗФ №3.xlsx', '3'], 'Торговый зал УКФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал УКФ №3.xlsx', '3'], 'Торговый зал ШФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №3.xlsx', '3'], 'Торговый зал АСФ №30': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №30.xlsx', '30'], 'Торговый зал ШФ №30': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №30.xlsx', '30'], 'Торговый зал АСФ №31': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №31.xlsx', '31'], 'Торговый зал АСФ №32': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №32.xlsx', '32'], 'Торговый зал ШФ №32': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №32.xlsx', '32'], 'Торговый зал АСФ №33': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №33.xlsx', '33'], 'Торговый зал ШФ №33': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №33.xlsx', '33'], 'Торговый зал АСФ №34': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №34.xlsx', '34'], 'Торговый зал ШФ №34': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №34.xlsx', '34'], 'Торговый зал АСФ №35': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №35.xlsx', '35'], 'Торговый зал_ОПТ ШФ №35': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал_ОПТ ШФ №35.xlsx', '35'], 'Торговый зал АСФ №36': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №36.xlsx', '36'], 'Торговый зал АФ №36': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №36.xlsx', '36'], 'Торговый зал АСФ №37': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №37.xlsx', '37'], 'Торговый зал АФ №37': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №37.xlsx', '37'], 'Торговый зал АСФ №38': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №38.xlsx', '38'], 'Торговый зал АФ №38': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №38.xlsx', '38'], 'Торговый зал АСФ №39': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №39.xlsx', '39'], 'Торговый зал АФ №39': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №39.xlsx', '39'], 'Торговый зал АСФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №4.xlsx', '4'], 'Торговый зал АФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №4.xlsx', '4'], 'Торговый зал КФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №4.xlsx', '4'], 'Торговый зал ППФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №4.xlsx', '4'], 'Торговый зал ТФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТФ №4.xlsx', '4'], 'Торговый зал ШФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №4.xlsx', '4'], 'Торговый зал АСФ №40': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №40.xlsx', '40'], 'Торговый зал АФ №40': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №40.xlsx', '40'], 'Торговый зал АСФ №41': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №41.xlsx', '41'], 'Торговый зал АФ №41': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №41.xlsx', '41'], 'Торговый зал АСФ №42': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №42.xlsx', '42'], 'Торговый зал АФ №42': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №42.xlsx', '42'], 'Торговый зал АФ №43': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №43.xlsx', '43'], 'Торговый зал АСФ №44': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №44.xlsx', '44'], 'Торговый зал АФ №44': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №44.xlsx', '44'], 'Торговый зал АСФ №45': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №45.xlsx', '45'], 'Торговый зал АФ №45': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №45.xlsx', '45'], 'Торговый зал АСФ №46': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №46.xlsx', '46'], 'Торговый зал АФ №46': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №46.xlsx', '46'], 'Торговый зал АСФ №47': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №47.xlsx', '47'], 'Торговый зал АФ №47': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №47.xlsx', '47'], 'Торговый зал АСФ №48': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №48.xlsx', '48'], 'Торговый зал АФ №48': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №48.xlsx', '48'], 'Торговый зал АФ №49': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №49.xlsx', '49'], 'Торговый зал АСФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №5.xlsx', '5'], 'Торговый зал КФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №5.xlsx', '5'], 'Торговый зал ППФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №5.xlsx', '5'], 'Торговый зал ШФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №5.xlsx', '5'], 'Торговый зал АСФ №50': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №50.xlsx', '50'], 'Торговый зал АФ №50': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №50.xlsx', '50'], 'Торговый зал АСФ №51': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №51.xlsx', '51'], 'Торговый зал АФ №51': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №51.xlsx', '51'], 'Торговый зал АСФ №52': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №52.xlsx', '52'], 'Торговый зал АФ №52': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №52.xlsx', '52'], 'Торговый зал АСФ №53': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №53.xlsx', '53'], 'Торговый зал АФ №53': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №53.xlsx', '53'], 'Торговый зал АСФ №54': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №54.xlsx', '54'], 'Торговый зал АФ №54': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №54.xlsx', '54'], 'Торговый зал АСФ №55': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №55.xlsx', '55'], 'Торговый_зал АФ №55 ОПТ': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый_зал АФ №55 ОПТ.xlsx', '55'], 'Торговый зал АСФ №56': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №56.xlsx', '56'], 'Торговый зал АФ №56': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №56.xlsx', '56'], 'Торговый зал АСФ №57': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №57.xlsx', '57'], 'Торговый зал АФ №57': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №57.xlsx', '57'], 'Торговый зал АСФ №58': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №58.xlsx', '58'], 'Торговый зал АФ №58': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №58.xlsx', '58'], 'Торговый зал АСФ №59': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №59.xlsx', '59'], 'Торговый зал АФ №59': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №59.xlsx', '59'], 'Торговый зал АСФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №6.xlsx', '6'], 'Торговый зал КФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №6.xlsx', '6'], 'Торговый зал ППФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №6.xlsx', '6'], 'Торговый зал ШФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №6.xlsx', '6'], 'Торговый зал АСФ №60': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №60.xlsx', '60'], 'Торговый зал АФ №60': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №60.xlsx', '60'], 'Торговый зал АСФ №61': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №61.xlsx', '61'], 'Торговый зал АФ №61': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №61.xlsx', '61'], 'Торговый зал АСФ №62': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №62.xlsx', '62'], 'Торговый зал АФ №62': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №62.xlsx', '62'], 'Торговый зал АСФ №63': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №63.xlsx', '63'], 'Торговый зал АФ №63': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №63.xlsx', '63'], 'Торговый зал АСФ №64': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №64.xlsx', '64'], 'Торговый зал АФ №64': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №64.xlsx', '64'], 'Торговый зал АСФ №65': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №65.xlsx', '65'], 'Торговый зал АФ №65': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №65.xlsx', '65'], 'Торговый зал АСФ №66': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №66.xlsx', '66'], 'Торговый зал АФ №66': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №66.xlsx', '66'], 'Торговый зал АСФ №67': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №67.xlsx', '67'], 'Торговый зал АФ №67': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №67.xlsx', '67'], 'Торговый зал АСФ №68': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №68.xlsx', '68'], 'Торговый зал АФ №68': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №68.xlsx', '68'], 'Торговый зал АСФ №69': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №69.xlsx', '69'], 'Торговый зал АФ №69': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №69.xlsx', '69'], 'Торговый зал КФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №7.xlsx', '7'], 'Торговый зал ППФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №7.xlsx', '7'], 'Торговый зал ШФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №7.xlsx', '7'], 'Торговый зал АФ №70': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №70.xlsx', '70'], 'Торговый зал АСФ №71': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №71.xlsx', '71'], 'Торговый зал АСФ №72': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №72.xlsx', '72'], 'Торговый зал АФ №72': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №72.xlsx', '72'], 'Торговый зал АСФ №73': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №73.xlsx', '73'], 'Торговый зал АФ №73': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №73.xlsx', '73'], 'Торговый зал АСФ №74': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №74.xlsx', '74'], 'Торговый зал АСФ №75': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №75.xlsx', '75'], 'Торговый зал АФ №75': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №75.xlsx', '75'], 'Торговый зал АСФ №76': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №76.xlsx', '76'], 'Торговый зал АФ №76': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №76.xlsx', '76'], 'Торговый зал АСФ №77': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №77.xlsx', '77'], 'Торговый зал АФ №77': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №77.xlsx', '77'], 'Торговый зал АФ №78': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №78.xlsx', '78'], 'Торговый зал АСФ №79': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №79.xlsx', '79'], 'Торговый зал АСФ №8': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №8.xlsx', '8'], 'Торговый зал ППФ №8': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №8.xlsx', '8'], 'Торговый зал ШФ №8': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №8.xlsx', '8'], 'Торговый зал АСФ №80': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №80.xlsx', '80'], 'Торговый зал АФ №80': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №80.xlsx', '80'], 'Торговый зал АСФ №81': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №81.xlsx', '81'], 'Торговый зал АФ №81': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №81.xlsx', '81'], 'Торговый зал АСФ №82': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №82.xlsx', '82'], 'Торговый зал АФ №82': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №82.xlsx', '82'], 'Торговый зал АСФ №83': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №83.xlsx', '83'], 'Торговый зал АФ №83': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №83.xlsx', '83'], 'Торговый зал АФ №84': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №84.xlsx', '84'], 'Торговый зал АФ №86': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №86.xlsx', '86'], 'Торговый зал ШФ №9': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №9.xlsx', '9'], 'Торговый зал ППФ №9': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №9.xlsx', '9'], 'Торговый зал СТМ 5АСФ': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал СТМ 5АСФ.xlsx', '1'], 'Торговый зал СТМ 1АФ': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал СТМ 1АФ.xlsx', '3'], 'Торговый зал АФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №5.xlsx', '5']}]
        print(all_excels)
        logger.info(all_excels)
        logger.warning(all_excels)
        for excel in all_excels:
            for key, val in excel.items():
                print(key, val)
            # print(excel)
            # print('---')
        exit()
        for filepath_ in os.listdir(main_excel_files):

            shutil.copy(os.path.join(main_excel_files, filepath_), os.path.join(filled_files, filepath_))

            # break

            if 'АФ1-АФ39' not in filepath_:
                continue

            print(all_excels)
            logger.info(all_excels)
            logger.warning(all_excels)

            # all_excels = [{'Торговый зал АФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №1.xlsx', '1'], 'Торговый зал АФ №10': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №10.xlsx', '10'], 'Торговый зал АФ №11': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №11.xlsx', '11']}]

            print("FINISHED!!!!")

            print(all_excels)
            print('----')
            for excel in all_excels:
                print(excel)
            # exit()
            col = None # 'F'
            for excel in all_excels:
                for branch, vals in excel.items():

                    col = edit_main_excel_file(os.path.join(filled_files, filepath_), branch, vals[0], vals[1])

                    if '~' not in vals:
                        # if '21' not in branch:
                        #     continue
                        print(branch, vals)

                        short_name, first_part, second_part = get_data_to_fill(os.path.join(filled_files, filepath_), branch, col)

                        print('-------------')
                        print(first_part, second_part, sep='\n')

                        start_time = time.time()
                        insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch, short_name=short_name,
                                          executor_name=ip_address, status_='processing', status_1c='success', error_reason='', error_saved_path='', execution_time=0, ecp_path_=os.path.join(ecp_paths, branch))
                        if True:
                            status, error_saved_path, error = start_single_branch(branch, first_part, second_part)

                            insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch, short_name=short_name,
                                              executor_name=ip_address, status_=status, status_1c='success', error_reason=error, error_saved_path=error_saved_path, execution_time=round(time.time() - start_time), ecp_path_=os.path.join(ecp_paths, branch))

                        # except Exception as error:
                        #
                        #     insert_data_in_db(started_time=datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"), store_name=branch, short_name=short_name,
                        #                       executor_name=ip_address, status_='failed with error', status_1c='success', error_reason=str(error), error_saved_path='', execution_time=round(time.time() - start_time), ecp_path_=os.path.join(ecp_paths, branch))

    except Exception as errorik:
        print(f'ERROR: {errorik}')
        logger.info(f'ERROR: {errorik}')
        logger.warning(f'ERROR: {errorik}')


