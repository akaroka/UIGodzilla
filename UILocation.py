import comtypes.client
import comtypes.gen.UIAutomationClient
from _ctypes import COMError


def element_match(_element, **kwargs):
    try:
        element_type = _element.CurrentControlType
    except COMError:
        return False
    if element_type != kwargs.pop('Type', 'default_type'):
        return False
    try:
        automation_id = _element.CurrentAutomationId
        if automation_id == kwargs.pop('AutomationId', 'default_automation_id'):
            return True
    except COMError:
        pass
    try:
        element_name = _element.CurrentName
        if element_name == kwargs.pop('Name', 'default_name'):
            return True
        if kwargs.pop('Name', 'default_name') in element_name:
            return True
    except COMError:
        pass
    try:
        element_class_name = _element.CurrentClassName
        if element_class_name == kwargs.pop('ClassName', 'default_class_name'):
            return True
    except COMError:
        pass
    return False


class UIScan:

    def __init__(self):
        self.UIA_dll = comtypes.client.GetModule('UIAutomationCore.dll')
        self.ui_automation_client = comtypes.gen.UIAutomationClient
        self.iuia = comtypes.CoCreateInstance(
                self.ui_automation_client.CUIAutomation().IPersist_GetClassID(),
                interface=self.ui_automation_client.IUIAutomation,
                clsctx=comtypes.CLSCTX_INPROC_SERVER
            )

    def find_all(self, root_element, **kwargs):
        this_generation_list = [root_element]
        next_generation_list = []
        result_list = []
        while True:
            while len(this_generation_list):
                _element = this_generation_list.pop()
                if element_match(_element, **kwargs):
                    result_list.append(_element)
                else:
                    _children = _element.FindAll(self.UIA_dll.TreeScope_Children, self.iuia.CreateTrueCondition())
                    _child_num = _children.Length
                    if _child_num == 0:
                        continue
                    else:
                        for _child_id in range(_child_num):
                            next_generation_list.append(_children.GetElement(_child_id))
            this_generation_list, next_generation_list = next_generation_list, []
            if not this_generation_list:
                break
        return result_list


if __name__ == '__main__':
    from time import time, sleep
    import subprocess
    subprocess.Popen("notepad.exe")
    sleep(3)
    example_scan = UIScan()
    windows_root = example_scan.iuia.GetRootElement()
    start_find_notepad_time = time()
    aa_list = example_scan.find_all(root_element=windows_root, Type=50032, ClassName='Notepad')
    print("查找记事本窗口耗时：%s秒" % str(round(time() - start_find_notepad_time, 2)))
    start_find_menu_bar_time = time()
    note_pad_list = example_scan.find_all(root_element=aa_list[0], Type=50010, AutomationId='MenuBar')
    print("查找menubar控件耗时：%s秒" % str(round(time() - start_find_menu_bar_time, 2)))

    for i in note_pad_list:
        print("menu_bar:%s" % i.CurrentName)













