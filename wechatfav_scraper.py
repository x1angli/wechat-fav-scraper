import csv
import logging
import os
import shlex
import subprocess
import time
from datetime import datetime
from typing import Set

import psutil
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto.timings import TimeoutError
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError


# Config Consts
RECORDS_CSV_PATH = 'data/records.csv'
CHROME_EXE_CMD_LINE = r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 2"'
WECHAT_EXE_CMD_LINE = r'D:\app-portable-social\WeChat\WeChat.exe'


# Global Non-config Consts
LOGGER = logging.getLogger(__name__)
FIELD_NAMES = ['title_uia', 'title_html', 'channel', 'date', 'url', 'summary']

# Initial configuring
logging.basicConfig(level=logging.INFO)



def is_process_running(process_name):
    return any(proc.info['name'] == process_name for proc in psutil.process_iter(['name']))


def parse_item_text(item_text:str):
    title_uia, _, channel, date_str = item_text.rsplit(' ', 3) # will probably throw a value error
    title_uia = reduce_text(title_uia.lstrip("[链接]"))
    if "年" not in date_str:
        date_str = "2024年" + date_str
    date_str = datetime.strptime(date_str, "%Y年%m月%d日").strftime("%Y-%m-%d")
    return {'title_uia': title_uia, 'channel': channel, 'date': date_str}


def reduce_text(text:str, max_length=40):
    no_breaks = ' '.join(text.splitlines())
    cleaned = no_breaks.strip()
    if len(cleaned) > max_length:
        cleaned = cleaned[:max_length - 3] + '...'

    return cleaned
def connect_and_run_application(app_name: str, class_name: str, proc_name: str, exe_path:str, max_attempts:int = 2):
    app = Application(backend='uia')
    if not is_process_running(proc_name):
        LOGGER.info(f"{app_name} is not running, starting the application")
        if app_name == 'Chrome':
            cmd_args = shlex.split(exe_path)
            subprocess.Popen(cmd_args)
            time.sleep(2)
        else:
            app.start(exe_path)
        LOGGER.info(f"Started new {app_name} application. ")

    for attempt_i in range(1, max_attempts + 1):
        try:
            # app.connect(class_name=class_name, timeout=1)
            app.connect(title_re=f".*{app_name}.*")
            LOGGER.info(f"Connected to {app_name} application on attempt #{attempt_i}.")
            break
        except (ElementNotFoundError, TimeoutError) as _nfe:
            LOGGER.warning(f"Attempt #{attempt_i}: {app_name} application not found. ")
            if app_name == 'WeChat':
                LOGGER.warning(f"Probably not logged in. Please login your WeChat acct asap")
        except ElementAmbiguousError as _ae:
            LOGGER.warning(f"Attempt #{attempt_i}: multiple instances of {app_name} application have been found! Please close duplicate ones and just keep a single one. ")
        time.sleep(5)
    else:
        LOGGER.error(f"Failed to connect to {app_name} application after {max_attempts} attempts. Exiting the program.")
        raise Exception(f"Unable to interact with {app_name} application. ")

    main_wnd = app.window(class_name=class_name)
    return main_wnd


def main_core(csv_writer, existing_title_uias: Set):
    chrome_wnd = connect_and_run_application('Chrome', 'Chrome_WidgetWin_1', 'chrome.exe', CHROME_EXE_CMD_LINE, max_attempts=2 )
    LOGGER.info("Prepared Chrome.")

    wechat_wnd = connect_and_run_application('WeChat', 'WeChatMainWndForPC', 'WeChat.exe', WECHAT_EXE_CMD_LINE, max_attempts=10)
    wechat_wnd.set_focus()
    LOGGER.info("Activated WeChat.")

    # Locate and click the favorite button
    try:
        allfav_button = wechat_wnd.child_window(title="收藏", control_type="Button")
        if not allfav_button.exists(timeout=1):
            LOGGER.error("Button '收藏' not found or not visible.")
            return
        allfav_button.click_input()
    except Exception as e:
        LOGGER.error(f"An error occurred: {e}")
        return

    # Locate and iterate the "全部收藏" list
    allfav_list = wechat_wnd.child_window(title="全部收藏", control_type="List")
    if not allfav_list.exists(timeout=1):
        LOGGER.error("List '全部收藏' not found or not visible.")
        return

    # Scroll to the top
    allfav_list.click_input(coords=(40, 120))
    send_keys('{HOME}')
    time.sleep(0.5)

    idx = 0
    while True:
        # Refresh the list
        allfav_list_items = allfav_list.descendants(control_type="ListItem")

        # Identify the favorite item
        if len(allfav_list_items) <= 1: # coz item[0] is always '列表开始'
            break
        item = allfav_list_items[1]
        item_text = item.window_text()
        if item_text == "列表结束":
            break
        try:
            item_dict = parse_item_text(item_text)
        except ValueError as _ve:
            LOGGER.error(f"Unable to parse item_dict for item {idx}")
            LOGGER.error(item_text)
            continue

        if item_dict['title_uia'] in existing_title_uias:
            LOGGER.info(f"Skipping {item_dict['title_uia']}")
        else:
            # Click the favorite item to open a new page tab
            item_buttons = item.descendants(title="", control_type="Button")
            if not item_buttons or len(item_buttons) == 0:
                LOGGER.error(f"Button for item {item_dict['title_uia']} not found or not visible.")
                continue
            item_button = item_buttons[0]
            item_button.click_input()
            time.sleep(7)

            # get metadata such as url, title, summary, and persist metadata into db
            url = chrome_wnd.child_window(title="Address and search bar", control_type="Edit").get_value()
            item_dict['url'] = url

            title = chrome_wnd.window_text()
            title = title.rsplit(' - ', 3)[0]
            item_dict['title_html'] = title

            try:
                summary = chrome_wnd.child_window(control_type="Document").window_text()
            except ElementAmbiguousError as _ae:
                summary = chrome_wnd.child_window(control_type="Document", found_index=0).window_text()
            except ElementNotFoundError as _nfe:
                summary = ""
            summary = reduce_text(summary, 80)
            item_dict['summary'] = summary

            # Write entry to CSV
            LOGGER.info(f"{item_dict}")
            csv_writer.writerow(item_dict)
            existing_title_uias.add(item_dict['title_uia'])

            # Ctrl + W to close the tab
            chrome_wnd.set_focus()
            time.sleep(0.2)
            send_keys('^w')
            time.sleep(0.2)

        # Delete the favorite item
        wechat_wnd.set_focus()
        item.right_click_input()
        time.sleep(0.2)
        del_menu_item = wechat_wnd.child_window(title="删除", control_type="MenuItem")
        del_menu_item.click_input()
        time.sleep(0.2)
        del_conf_bttn = wechat_wnd.child_window(title="删除", control_type="Button")
        del_conf_bttn.click_input()

        time.sleep(1)
        idx += 1
    # end of while-loop
    LOGGER.info(f"Finished the ")

def main():
    # Check if file exists and read existing rows to avoid duplicates
    existing_title_uias = set()
    if os.path.exists(RECORDS_CSV_PATH):
        with open(RECORDS_CSV_PATH, mode='r', encoding='utf-8-sig', newline='') as file:
            reader = csv.DictReader(file)
            for row in reader:
                existing_title_uias.add(row['title_uia'])

    # Write the header row if necessary
    with open(RECORDS_CSV_PATH, mode='a', encoding='utf-8-sig', newline='') as file:
        csv_writer = csv.DictWriter(file, fieldnames=FIELD_NAMES, quoting=csv.QUOTE_ALL, quotechar='"')

        if file.tell() == 0:
            csv_writer.writeheader()

        main_core(csv_writer, existing_title_uias)


if __name__ == '__main__':
    main()
