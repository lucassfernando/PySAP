import win32com.client
import subprocess
from time import sleep
import pythoncom


class SapFunctions():

    def __init__(self, server_name, username, password):
        self.server_name = server_name
        self.username = username
        self.password = password

    def login(self):

        pythoncom.CoInitialize()

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)

        sleep(10)

        try:

            sapgui = win32com.client.GetObject('SAPGUI').GetScriptingEngine
            connection = sapgui.OpenConnection(self.server_name, True)

            self.session = connection.Children(0)
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.username
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
            self.session.findById("wnd[0]").sendVKey(0)

            return True
        
        except:
            
            return False

    def sessions_active(self):  #CHECK THE TOTAL OF ACTIVE WINDOWS
        sapgui = win32com.client.GetObject('SAPGUI').GetScriptingEngine
        sessions_list = []
        
        for session in range(sapgui.Children.Count):
            connection = sapgui.Children(session)
            sessions_list.append(f'Session: {connection.Name} | Server: {connection.Description}')
        
        return sessions_list
            
    #VARIOUS FUNCTIONS TO PERFORM WITHIN SAP LOGO BELOW (ALWAYS BEING UPDATED)

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
    
    def click_current_cell(self, button_path):
        self.session.findById(button_path).clickCurrentCell()
    
    def modify_cell(self, button_path, cell_position, cell_name, text):
        self.session.findById(button_path).modifyCell(int(cell_position), cell_name, text)

    def scroll_bar_position(self, button_path, value):
        self.session.findById(button_path).verticalScrollbar.position = value
