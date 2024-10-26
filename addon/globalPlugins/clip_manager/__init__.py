import sys
import os
sys.path.append(os.path.dirname(__file__))

from scriptHandler import script
import globalPluginHandler, ui, api, gui, config, NVDAObjects
import pyperclip
import ctypes
import ctypes.wintypes
import wx
from gui import guiHelper, mainFrame, NVDASettingsDialog
from gui.settingsDialogs import SettingsPanel as BaseSettingsPanel
import pyautogui
import json
import time

WM_CLIPBOARDUPDATE = 0x031D
WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.wintypes.HWND, ctypes.c_uint, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM)

HCURSOR = ctypes.c_void_p
HICON = ctypes.c_void_p
HBRUSH = ctypes.c_void_p

class WNDCLASS(ctypes.Structure):
    _fields_ = [
        ("style", ctypes.c_uint),
        ("lpfnWndProc", ctypes.c_void_p), 
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", ctypes.wintypes.HINSTANCE),
        ("hIcon", HICON),
        ("hCursor", HCURSOR),
        ("hbrBackground", HBRUSH),
        ("lpszMenuName", ctypes.wintypes.LPCWSTR),
        ("lpszClassName", ctypes.wintypes.LPCWSTR),
    ]


class FullTextFrame(wx.Frame):
    """Frame to display the full text of a clipboard item."""
    def __init__(self, parent, text):
        wx.Frame.__init__(self, parent, title=_("Full Clipboard Text"), size=(400, 300))
        self.panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.text_ctrl = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.text_ctrl.SetValue(text)
        sizer.Add(self.text_ctrl, 1, wx.ALL | wx.EXPAND, 10)

        close_btn = wx.Button(self.panel, wx.ID_CLOSE, _("Close"))
        close_btn.Bind(wx.EVT_BUTTON, self.onClose)
        sizer.Add(close_btn, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 10)

        self.panel.SetSizer(sizer)
        self.text_ctrl.SetFocus()

    def onClose(self, event):
        self.Destroy()



class ClipboardHistoryFrame(wx.Frame): 
    def __init__(self, parent, clipboard_history, obj, clear_history, plugin_instance):
        wx.Frame.__init__(self, parent, title=_("Clipboard History"), size=(300, 200))
        self.clipboard_history = clipboard_history
        self.obj=obj
        self.clear_history = clear_history
        self.plugin_instance = plugin_instance

        self.panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.listbox = wx.ListBox(self.panel, choices=[item['text'] if isinstance(item, dict) and 'text' in item else item for item in self.clipboard_history], style=wx.LB_SINGLE)
        sizer.Add(self.listbox, 1, wx.ALL | wx.EXPAND, 10)

        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.pasteBtn = wx.Button(self.panel, wx.ID_OK, _("Paste"))
        self.closeBtn = wx.Button(self.panel, wx.ID_CANCEL, _("Close"))
        buttonSizer.Add(self.pasteBtn, 0, wx.ALL, 5)
        buttonSizer.Add(self.closeBtn, 0, wx.ALL, 5)
        sizer.Add(buttonSizer, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 10) 

        self.panel.SetSizer(sizer)
        self.Fit()

        self.pasteBtn.Bind(wx.EVT_BUTTON, self.onPaste)
        self.closeBtn.Bind(wx.EVT_BUTTON, self.onClose)
        self.listbox.Bind(wx.EVT_LISTBOX_DCLICK, self.onPaste)
        self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
        self.listbox.Bind(wx.EVT_CONTEXT_MENU, self.on_context_menu)

        self.listbox.SetFocus()


    def on_context_menu(self, event):
        """Displays the context menu."""
        selection = self.listbox.GetSelection()
        if selection == wx.NOT_FOUND:
            return 

        menu = wx.Menu()
        delete_item = menu.Append(wx.ID_ANY, _("Delete"))
        if selection < len(self.clipboard_history):
            item_data = self.clipboard_history[selection]
            if isinstance(item_data, dict) and 'pinned' in item_data:
                pin_item = menu.Append(wx.ID_ANY, _("Unpin") if item_data['pinned'] else _("Pin"))
            else:
                pin_item = menu.Append(wx.ID_ANY, _("Pin"))
            text = item_data.get('text', '') if isinstance(item_data, dict) else item_data
            if len(text) > config.conf['clipManager']['displayChars']:
                view_full_text_item = menu.Append(wx.ID_ANY, _("View Full Text"))
                self.Bind(wx.EVT_MENU, lambda evt: self.viewFullText(text), view_full_text_item)

        else:
            pin_item = menu.Append(wx.ID_ANY, _("Pin"))
        clear_all_item = menu.Append(wx.ID_ANY, _("Clear All"))

        self.Bind(wx.EVT_MENU, lambda evt: self.deleteItem(selection), delete_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.togglePinItem(selection), pin_item)
        self.Bind(wx.EVT_MENU, self.clearAllItems, clear_all_item)

        self.PopupMenu(menu)
    
    def onPaste(self, event):
        selection = self.listbox.GetSelection()
        if selection != wx.NOT_FOUND:
            try:
                selected_text = self.clipboard_history[selection]

                if self.obj.role == 8:
                    wx.CallAfter(pyautogui.hotkey, 'ctrl', 'v')
                    wx.CallAfter(ui.message, _("Pasted!"))
                else:
                    ui.message(_("Cannot paste: Focus is not on an editable text control."))
            except Exception as e:
                ui.message(_("Error pasting: {}").format(str(e)))
        pyperclip.copy(selected_text)
        self.Close()

    def on_delete(self, event):
        selection = self.listbox.GetSelection()
        if selection != wx.NOT_FOUND:
            del self.clipboard_history[selection]
            self.listbox.Delete(selection)      

    def onClose(self, event):
        self.Destroy()

    def onCharHook(self, event):
        keycode = event.GetKeyCode()

        if keycode == wx.WXK_ESCAPE:
            self.onClose(event) 
        elif keycode == wx.WXK_RETURN:
            self.onPaste(event)
        elif keycode == wx.WXK_DELETE:  # Bind delete key
            self.deleteItem(self.listbox.GetSelection())
        else:
            event.Skip()

    def viewFullText(self, text):
        """Opens a new frame to display the full text."""
        full_text_frame = FullTextFrame(self, text)
        full_text_frame.Show()

    def clearAllItems(self, event):
        """Clears all unpinned items from the listbox and clipboard history."""
        self.clipboard_history[:] = [item for item in self.clipboard_history if isinstance(item, dict) and 'pinned' in item and item['pinned']]
        self.listbox.Clear()
        self.listbox.AppendItems([item['text'] if isinstance(item, dict) and 'text' in item else item for item in self.clipboard_history])
        self.plugin_instance.saveHistory()

    def deleteItem(self, selection):
        """Deletes the selected item from the listbox and clipboard history."""
        if selection != wx.NOT_FOUND:
            del self.clipboard_history[selection]
            self.listbox.Delete(selection)      
            self.listbox.SetSelection(0)
        self.plugin_instance.saveHistory()

    def togglePinItem(self, selection):
        """Toggles the pin status of the selected item."""
        if selection != wx.NOT_FOUND:
            if isinstance(self.clipboard_history[selection], dict):
                self.clipboard_history[selection]['pinned'] = not self.clipboard_history[selection]['pinned']
            else:
                self.clipboard_history[selection] = {'text': self.clipboard_history[selection], 'pinned': True}
        self.plugin_instance.saveHistory()


class Settings(BaseSettingsPanel):
    title = _("Clipboard History Settings")

    def makeSettings(self, settingsSizer):
        sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

        sHelper.addItem(wx.StaticText(self, label=_("History size:")))
        self.historySizeSpin = sHelper.addItem(wx.SpinCtrl(self, -1, "", size=(60, -1), style=wx.SP_ARROW_KEYS, min=1, max=100000000))
        self.historySizeSpin.SetValue(config.conf['clipManager'].get('historySize', 500))
        self.historySizeSpin.Bind(wx.EVT_SPINCTRL, self.onSpinCtrlChanged)

        self.saveHistoryCheck = sHelper.addItem(wx.CheckBox(self, -1, _("Save history on exit")))
        self.saveHistoryCheck.SetValue(config.conf['clipManager'].get('saveHistory', True))
        self.saveHistoryCheck.Bind(wx.EVT_CHECKBOX, self.onCheckChanged)

        sHelper.addItem(wx.StaticText(self, label=_("Save threshold:")))
        self.saveThresholdSpin = sHelper.addItem(wx.SpinCtrl(self, -1, "", size=(60, -1), style=wx.SP_ARROW_KEYS, min=1, max=100))
        self.saveThresholdSpin.SetValue(config.conf['clipManager'].get('saveThreshold', 5))
        self.saveThresholdSpin.Bind(wx.EVT_SPINCTRL, self.onSaveThresholdChanged)

        sHelper.addItem(wx.StaticText(self, label=_("Display characters per item:")))
        self.displayCharsSpin = sHelper.addItem(wx.SpinCtrl(self, -1, "", size=(60, -1), style=wx.SP_ARROW_KEYS, min=1, max=50000))
        self.displayCharsSpin.SetValue(config.conf['clipManager'].get('displayChars', 2000))
        self.displayCharsSpin.Bind(wx.EVT_SPINCTRL, self.onDisplayCharsChanged)

    def onSpinCtrlChanged(self, event):
        config.conf['clipManager']['historySize'] = self.historySizeSpin.GetValue()
    
    def onCheckChanged(self, event):
        config.conf['clipManager']['saveHistory'] = self.saveHistoryCheck.GetValue()

    def onSaveThresholdChanged(self, event):
        config.conf['clipManager']['saveThreshold'] = self.saveThresholdSpin.GetValue()

    def onDisplayCharsChanged(self, event):
        config.conf['clipManager']['displayChars'] = self.displayCharsSpin.GetValue()

    def onSave(self):
        config.conf.save()

clipManagerSection = "clipManager"
clipManagerSetting = {
    "historySize": "integer(default=500)",
    "saveHistory": "boolean(default=True)",
    "saveThreshold": "integer(default=5)",
    "displayChars": "integer(default=2000)",
}
config.conf.spec[clipManagerSection] = clipManagerSetting


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    clipboard_history = []
    dlg = None
    history_file = os.path.join(os.path.dirname(__file__), "history.json")
    copy_count = 0 # Counter for copy attempts

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.loadHistory()

        # Create hidden window
        self.hinst = ctypes.windll.kernel32.GetModuleHandleW(0)
        self.wndclass = WNDCLASS()
        self.wndclass.lpszClassName = "MyWindowClass"
        self.wndclass.hInstance = self.hinst
        self.wndclass.lpfnWndProc = ctypes.cast(WNDPROC(self.wndProc), ctypes.c_void_p)
        self.regClass = ctypes.windll.user32.RegisterClassW(ctypes.byref(self.wndclass))

        self.hwnd = ctypes.windll.user32.CreateWindowExW(
            0,
            self.wndclass.lpszClassName,
            "Clipboard Listener",
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            self.hinst,
            0,
        )

        ctypes.windll.user32.AddClipboardFormatListener(self.hwnd)
        NVDASettingsDialog.categoryClasses.append(Settings)

    def terminate(self):
        self.saveHistory()
        NVDASettingsDialog.categoryClasses.remove(Settings)
        ctypes.windll.user32.RemoveClipboardFormatListener(self.hwnd)
        ctypes.windll.user32.DestroyWindow(self.hwnd)
        ctypes.windll.user32.UnregisterClassW(self.wndclass.lpszClassName, self.hinst)

    def wndProc(self, hwnd, msg, wparam, lparam):
        if msg == WM_CLIPBOARDUPDATE:
            self.onClipboardUpdate()
        return ctypes.windll.user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def clearClipboardHistory(self):
        """Clears the clipboard history list."""
        self.clipboard_history = [item for item in self.clipboard_history if isinstance(item, dict) and 'pinned' in item and item['pinned']]

    @script(
        description=_("Shows clipboard history"),
        gesture="kb:NVDA+v",
        category=_("clip_manager")
    )
    def script_showClipboardHistory(self, gesture):
        if not self.clipboard_history:
            ui.message(_("Clipboard history is empty."))
            return

        obj=api.getFocusObject()
        mainFrame.prePopup()
        self.dlg = ClipboardHistoryFrame(mainFrame, self.clipboard_history, obj, self.clearClipboardHistory, self)
        self.dlg.CentreOnScreen()

        # Update Listbox with truncated text
        self.dlg.listbox.Clear()
        for item in self.clipboard_history:
            text = item.get('text', '') if isinstance(item, dict) else item
            truncated_text = text[:config.conf['clipManager']['displayChars']]
            if len(text) > config.conf['clipManager']['displayChars']:
                truncated_text += " (See more)"
            self.dlg.listbox.Append(truncated_text)

        self.dlg.Show(True)
        mainFrame.postPopup()

    def onClipboardUpdate(self):
        try:
            new_text = pyperclip.paste()

            if new_text not in [item['text'] if isinstance(item, dict) and 'text' in item else item for item in self.clipboard_history]:
                self.clipboard_history.insert(0, new_text)
                if len(self.clipboard_history) > config.conf['clipManager']['historySize']:
                    self.clipboard_history.pop()
                self.copy_count += 1
                if self.copy_count >= config.conf['clipManager']['saveThreshold']:
                    self.saveHistory()
                    self.copy_count = 0

        except Exception as e:
            ui.message("Failed to get clipboard text")

    def loadHistory(self):
        """Loads clipboard history from file if saveHistory is enabled."""
        if config.conf['clipManager'].get("saveHistory", True) and os.path.exists(self.history_file):
            try:
                with open(self.history_file, "r") as f:
                    self.clipboard_history = json.load(f)
            except:
                ui.message(_("Error loading clipboard history!."))
    
    def saveHistory(self):
        """Saves clipboard history to file if saveHistory is enabled."""
        if config.conf['clipManager'].get("saveHistory", True):
            try:
                with open(self.history_file, "w", encoding="utf-8") as f:
                    json.dump(self.clipboard_history, f, ensure_ascii=False)
            except:
                ui.message(_("Error saving clipboard history!"))
