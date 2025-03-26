import win32com.client
import subprocess
from time import sleep
import pythoncom


class SapFunctions():

    def __init__(self, server_name, username, password):
        self.server_name = server_name
        self.username = username
        self.password = password

    @log_error
    def login(self):

        pythoncom.CoInitialize()

        path = "C:\\Program Files\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe"
        subprocess.Popen(path)

        sleep(10)

        sapgui = win32com.client.GetObject('SAPGUI').GetScriptingEngine
        connection = sapgui.OpenConnection(self.server_name, True)

        self.session = connection.Children(0)
        self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.username
        self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
        self.session.findById("wnd[0]").sendVKey(0)

        return True

    def sessions_active(self):  #VERIFICA O TOTAL DE JANELAS ATIVAS
        sapgui = win32com.client.GetObject('SAPGUI').GetScriptingEngine
        sessions_list = []
        
        for session in range(sapgui.Children.Count):
            connection = sapgui.Children(session)
            sessions_list.append(f'Session: {connection.Name} | Server: {connection.Description}')
        
        return sessions_list
    
    def detect_any_window(self):
        sleep(1)
        try:
            for i in range(1, 5):
                window_id = f"wnd[{i}]"
                self.session.findById(window_id)
                return True
        except:
            return False
        
    #DIVERSAS FUNÇÕES PARA REALIZAR DENTRO DO SAP LOGO ABAIXO (SEMPRE SENDO ATUALIZADO)

    def tab_close(self):
        self.session.findById("wnd[0]").close()
        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        pythoncom.CoUninitialize()

    def write_text(self, button_path, text):
        self.session.findById(button_path).text = text

    def send_key(self, button_path, key):
        self.session.findById(button_path).sendVKey(int(key))

    def set_focus(self, button_path):
        self.session.findById(button_path).setFocus()
    
    def press(self, button_path):
        self.session.findById(button_path).press()

    def selected_rows(self, button_path, line):
        self.session.findById(button_path).selectedRows = int(line)
    
    def press_tool_bar_button(self, button_path, button_name):
        self.session.findById(button_path).pressToolbarButton(button_name)

    def caret_position(self, button_path, position_number):
        self.session.findById(button_path).caretPosition = int(position_number)
    
    def select(self, button_path):
        self.session.findById(button_path).select()
    
    def double_click(self, button_path):
        self.session.findById(button_path).doubleClick()
    
    def close_information_window(self, button_path):
        self.session.findById(button_path).close()
    
    def text_extract(self, button_path):
        text = self.session.findById(button_path).text
        return text
    
    def select_box_button(self, button_path, true_or_false_option):
        self.session.findById(button_path).selected = true_or_false_option

    def modify_check_box(self, button_path, box_position, box_name, true_or_false_option):
        self.session.findById(button_path).modifyCheckbox(int(box_position), box_name, true_or_false_option)
    
    def double_click_item(self, button_path, item, item_postion):
        self.session.findById(button_path).doubleClickItem(item, item_postion)
    
    def select_item(self, button_path, item, item_postion):
        self.session.findById(button_path).selectItem(item, item_postion)
    
    def ensure_visible_horizontal_item(self, button_path, item, item_postion):
        self.session.findById(button_path).ensureVisibleHorizontalItem(item, item_postion)

    def suspended_list_key(self, button_path, key):
        self.session.findById(button_path).key = f"{key}"
    
    def current_cell_column(self, button_path, selected_cell):
        self.session.findById(button_path).currentCellColumn = selected_cell

    def select_column(self, button_path, selected_cell):
        self.session.findById(button_path).selectColumn(selected_cell)

    def current_cell_row(self, button_path, selected_row):
        self.session.findById(button_path).currentCellRow = int(selected_row)
    
    def click_current_cell(self, button_path):
        self.session.findById(button_path).clickCurrentCell()

    def double_click_current_cell(self, button_path):
        self.session.findById(button_path).doubleClickCurrentCell()
    
    def modify_cell(self, button_path, cell_position, cell_name, text):
        self.session.findById(button_path).modifyCell(int(cell_position), cell_name, text)

    def scroll_bar_position(self, button_path, value):
        self.session.findById(button_path).verticalScrollbar.position = value

    def press_button(self, button_path, button_name):
        self.session.findById(button_path).pressButton(button_name)
    
    def get_cell_value(self, button_path, cell, field_name):
        cell_value = self.session.findById(button_path).GetCellValue(int(cell), field_name)
        return cell_value
    
    def context_menu_select_item(self, button_path, item):
        self.session.findById(button_path).contextMenu()
        self.session.findById(button_path).selectContextMenuItem(item)

    def get_absolute_row(self, button_path, line, true_or_false):
        self.session.findById(button_path).getAbsoluteRow(line).selected = true_or_false

    def set_current_cell(self, button_path, cell, field_name):
        self.session.findById(button_path).setCurrentCell(int(cell), field_name)
    
    def select_node(self, button_path, node):
        self.session.findById(button_path).selectedNode = f'{node}'

