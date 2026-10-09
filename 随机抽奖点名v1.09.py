import os
import random
import openpyxl
import PySimpleGUI as sg
import threading
import time

def run_lucky_draw():
    def doexcel(excel_path):
        user_list = []
        user_info = {}
        column_names = []
        try:
            w0=openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
            sheet=w0.worksheets[0]
            for idx, cell in enumerate(sheet[1]):
                column_names.append(cell.value)
            for row in sheet.iter_rows(min_row=2):
                candidate = row[3].value if len(row) >= 4 else None
                if candidate is None or str(candidate).strip() == "":
                    continue
                user_info[candidate] = [cell.value for cell in row]
                user_list.append(candidate)
            w0.close()
        except FileNotFoundError:
            sg.popup('无法找到 "users.xlsx" 文件。', title='文件错误', font=('Arial Black',13))
        except Exception as e:
            sg.popup(f"在读取Excel文件时发生错误: {str(e)}", title='文件错误', font=('Arial Black',13))
        return user_list, user_info, column_names

    layout=[
        [sg.Text('随机点名',key='show_stu',size=(20,1),font=('黑体',32),text_color='red',justification='center')],
        [sg.Text('', key='show_info', size=(60,1), font=('黑体',11), text_color='blue', justification='center')],
        [sg.Text('')],
        [sg.CB('不重复',key='NOCF',default=True, font=('黑体',12)),
        sg.Button('开始',key='butt_dm', font=('黑体',12)),
        sg.Button('确认',key='butt_confirm', font=('黑体',12), visible=False),
        sg.In('users.xlsx',key='file_url',visible = False),
        sg.FileBrowse(button_text='名单',key='stu_md',target='file_url',file_types=(("Excel 文件","*.xlsx"),), font=('黑体',12)),
        sg.Button('重置',key='butt_INIT', font=('黑体',12))],
    ]

    window=sg.Window('随机点名',layout,location=(0,0),keep_on_top=True, size=(450, 220))
    dmstu = None
    running = False
    user_list = []
    user_info = {}
    column_names = []
    loaded_path = None

    while True:
        event,values=window.read()

        if event is None:
            running = False
            break

        if event == "-RANDOM-":
            if running:
                dmstu = values[event]
                window["show_stu"].update(dmstu, text_color="red")
                window["show_info"].update("")

        elif event == "butt_INIT":
            running = False
            dmstu = None
            user_list = []
            user_info = {}
            column_names = []
            loaded_path = None
            window["show_stu"].update("随机点名", text_color="red")
            window["show_info"].update("")
            window["butt_confirm"].update(visible=False)
            window["butt_dm"].update(disabled=False)

        elif event == "butt_dm":
            if running:
                continue

            selected_path = values.get("file_url") or "users.xlsx"
            if selected_path != loaded_path:
                user_list, user_info, column_names = doexcel(selected_path)
                loaded_path = selected_path

            if not user_list:
                message = "名单已抽完，请点击重置" if user_info else "空名单"
                window["show_stu"].update(message, text_color="red")
                continue

            running = True
            dmstu = None
            window["butt_dm"].update(disabled=True)
            window["butt_INIT"].update(disabled=True)
            window["butt_confirm"].update(visible=True)

            def update_randomly():
                while running:
                    try:
                        window.write_event_value("-RANDOM-", random.choice(user_list))
                    except Exception:
                        break
                    time.sleep(0.1)

            threading.Thread(target=update_randomly, daemon=True).start()

        elif event == "butt_confirm":
            running = False
            window["butt_dm"].update(disabled=False)
            window["butt_INIT"].update(disabled=False)
            window["butt_confirm"].update(visible=False)
            if dmstu is None:
                continue

            window["show_stu"].update(dmstu, text_color="blue")
            info_list = user_info[dmstu]
            info_text = " ".join(
                f"{name}: {info}"
                for name, info in zip(column_names, info_list)
                if name is not None and info is not None
            )
            window["show_info"].update(info_text)
            if values.get("NOCF", True) and dmstu in user_list:
                user_list.remove(dmstu)

    window.close()

run_lucky_draw()
