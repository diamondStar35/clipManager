import addonHandler
addonHandler.initTranslation()

import sys
import os
sys.path.append(os.path.dirname(__file__))

from scriptHandler import script
import globalPluginHandler, ui, api, gui, config, NVDAObjects, wx
import pyperclip
import ctypes
import ctypes.wintypes
import wx
from gui import guiHelper, mainFrame, NVDASettingsDialog
from gui.settingsDialogs import SettingsPanel as BaseSettingsPanel
import pyautogui
import json
import time
from .search import SearchDialog

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
        # Translators: Title for the window displaying the full text of a clipboard item.
        wx.Frame.__init__(self, parent, title=_("Full Clipboard Text"), size=(400, 300))
        self.panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.text_ctrl = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL)
        self.text_ctrl.SetValue(text)
        sizer.Add(self.text_ctrl, 1, wx.ALL | wx.EXPAND, 10)

        # Translators: Label for the close button in the Full Clipboard Text window.
        close_btn = wx.Button(self.panel, wx.ID_CLOSE, _("Close"))
        close_btn.Bind(wx.EVT_BUTTON, self.onClose)
        sizer.Add(close_btn, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 10)

        self.panel.SetSizer(sizer)
        self.text_ctrl.SetFocus()

    def onClose(self, event):
        self.Destroy()

class AddDialog(wx.Dialog):
    def __init__(self, parent, onSave):
        # Translators: Title for the add dialog.
        super(AddDialog, self).__init__(parent, title=_("Add New Item"), size=(400, 300))
        self.onSave = onSave

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Translators: Label for the text entry area in the add dialog.
        text_label = wx.StaticText(panel, label=_("Enter text:"))
        sizer.Add(text_label, 0, wx.ALL, 5)

        self.text_ctrl = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.HSCROLL)
        sizer.Add(self.text_ctrl, 1, wx.EXPAND | wx.ALL, 10)

        # Translators: Label for the pinned checkbox in the add dialog.
        self.pinned_checkbox = wx.CheckBox(panel, label=_("Pinned"))
        sizer.Add(self.pinned_checkbox, 0, wx.ALL, 10)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        # Translators: Label for the save button.
        save_btn = wx.Button(panel, label=_("Save"))
        save_btn.Bind(wx.EVT_BUTTON, self.on_save)
        button_sizer.Add(save_btn, 0, wx.ALL, 5)

        # Translators: Label for the cancel button in the add dialog.
        cancel_btn = wx.Button(panel, wx.ID_CANCEL)
        button_sizer.Add(cancel_btn, 0, wx.ALL, 5)

        sizer.Add(button_sizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        panel.SetSizer(sizer)

    def on_save(self, event):
        text = self.text_ctrl.GetValue()
        pinned = self.pinned_checkbox.GetValue()
        self.onSave(text, pinned)
        self.Destroy()

class EditDialog(wx.Dialog):
    def __init__(self, parent, title, text, onSave):
        super(EditDialog, self).__init__(parent, title=title, size=(400, 300))
        self.text = text
        self.onSave = onSave

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.text_ctrl = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.HSCROLL)
        self.text_ctrl.SetValue(self.text)
        sizer.Add(self.text_ctrl, 1, wx.EXPAND | wx.ALL, 10)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        # Translators: Label for the save button in the edit dialog.
        save_btn = wx.Button(panel, label=_("Save"))
        save_btn.Bind(wx.EVT_BUTTON, self.on_save)
        button_sizer.Add(save_btn, 0, wx.ALL, 5)

        # Translators: Label for the cancel button in the edit dialog.
        cancel_btn = wx.Button(panel, wx.ID_CANCEL)
        button_sizer.Add(cancel_btn, 0, wx.ALL, 5)

        sizer.Add(button_sizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        panel.SetSizer(sizer)

    def on_save(self, event):
        self.text = self.text_ctrl.GetValue()
        self.onSave(self.text)
        self.Destroy()


class ClipboardHistoryFrame(wx.Frame): 
    def __init__(self, parent, clipboard_history, obj, clear_history, plugin_instance):
        # Translators: Title for the window displaying the clipboard history.
        wx.Frame.__init__(self, parent, title=_("Clipboard History"), size=(300, 200))
        self.clipboard_history = clipboard_history
        self.obj=obj
        self.clear_history = clear_history
        self.plugin_instance = plugin_instance
        self.search_results = []
        self.current_result_index = -1

        self.panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.listbox = wx.ListBox(self.panel, choices=[item['text'] if isinstance(item, dict) and 'text' in item else item for item in self.clipboard_history], style=wx.LB_SINGLE)
        sizer.Add(self.listbox, 1, wx.ALL | wx.EXPAND, 10)

        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Translators: Label for the paste button in the clipboard history window.
        self.pasteBtn = wx.Button(self.panel, wx.ID_OK, _("Paste"))
        # Translators: Label for the close button in the clipboard history window.
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
        self.updateListBox()
        self.listbox.SetFocus()


    def on_context_menu(self, event):
        """Displays the context menu."""
        selection = self.listbox.GetSelection()
        if selection == wx.NOT_FOUND:
            return 

        menu = wx.Menu()
        # Translators: Label for the edit item in the context menu.
        edit_item = menu.Append(wx.ID_ANY, _("Edit"))
        # Translators: Label for the delete item in the context menu.
        delete_item = menu.Append(wx.ID_ANY, _("Delete"))
        if selection < len(self.clipboard_history):
            item_data = self.clipboard_history[selection]
            if isinstance(item_data, dict) and 'pinned' in item_data:
                # Translators: Label for the pin/unpin item in the context menu.
                pin_item = menu.Append(
                    wx.ID_ANY, _("Unpin") if item_data["pinned"] else _("Pin")
                )
            else:
                # Translators: Label for the pin item in the context menu.
                pin_item = menu.Append(wx.ID_ANY, _("Pin"))
            text = item_data.get('text', '') if isinstance(item_data, dict) else item_data
            if len(text) > config.conf['clipManager']['displayChars']:
                # Translators: Label for the 'view full text' item in the context menu.
                view_full_text_item = menu.Append(wx.ID_ANY, _("View Full Text"))
                self.Bind(
                    wx.EVT_MENU, lambda evt: self.viewFullText(text), view_full_text_item
                )

        else:
            # Translators: Label for the pin item in the context menu.
            pin_item = menu.Append(wx.ID_ANY, _("Pin"))
        # Translators: Label for the add item in the context menu.
        add_item = menu.Append(wx.ID_ANY, _("Add"))
        # Translators: Label for the clear all items in the context menu.
        clear_all_item = menu.Append(wx.ID_ANY, _("Clear All"))

        backup_menu = wx.Menu()
        # Translators: Label for the 'import history' item in the context menu.
        import_item = backup_menu.Append(wx.ID_ANY, _("Import history"))
        # Translators: Label for the 'export history' item in the context menu.
        export_item = backup_menu.Append(wx.ID_ANY, _("Export history"))

        self.Bind(wx.EVT_MENU, self.onImportHistory, import_item)
        self.Bind(wx.EVT_MENU, self.onExportHistory, export_item)
        menu.AppendSubMenu(backup_menu, _("Backup"))

        self.Bind(wx.EVT_MENU, lambda evt: self.onEdit(selection, text), edit_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.deleteItem(selection), delete_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.togglePinItem(selection), pin_item)
        self.Bind(wx.EVT_MENU, self.onAdd, add_item)
        self.Bind(wx.EVT_MENU, self.clearAllItems, clear_all_item)

        self.PopupMenu(menu)
    
    def onPaste(self, event):
        selection = self.listbox.GetSelection()
        if selection != wx.NOT_FOUND:
            try:
                selected_item = self.clipboard_history[selection]
                selected_text = selected_item.get("text", "") if isinstance(selected_item, dict) else selected_item

                if self.obj.role == 8:
                    wx.CallAfter(pyautogui.hotkey, "ctrl", "v")
                    # Translators: Message indicating that the selected item has been pasted.
                    wx.CallAfter(ui.message, _("Pasted!"))
                else:
                    # Translators: Message indicating that pasting is not possible because the focus is not on an editable text control.
                    ui.message(
                        _("Cannot paste: Focus is not on an editable text control.")
                    )
                if config.conf['clipManager']['moveToTop']:
                    self.clipboard_history.pop(selection)
                    self.clipboard_history.insert(0, selected_item)
                    self.updateListBox()
            except Exception as e:
                # Translators: Message indicating that an error occurred while pasting.
                ui.message(_("Error pasting: {}").format(str(e)))
        pyperclip.copy(selected_text)
        self.Close()

    def onEdit(self, selection, text):
        """Opens the edit dialog."""
        def on_save(new_text):
            # Update the item in clipboard history
            if isinstance(self.clipboard_history[selection], dict):
                self.clipboard_history[selection]['text'] = new_text
            else:
                self.clipboard_history[selection] = new_text
            self.updateListBox()
            self.plugin_instance.saveHistory()

        # Translators: Title for the edit dialog.
        edit_dialog = EditDialog(self, _("Edit Text"), text, on_save)
        edit_dialog.ShowModal()

    def on_delete(self, event):
        selection = self.listbox.GetSelection()
        if selection != wx.NOT_FOUND:
            del self.clipboard_history[selection]
            self.listbox.Delete(selection)      

    def show_search_dialog(self):
        """Opens the search dialog."""
        search_dialog = SearchDialog(self, self.clipboard_history)
        search_dialog.ShowModal()

    def onClose(self, event):
        self.Destroy()

    def onCharHook(self, event):
        keycode = event.GetKeyCode()
        modifiers = event.GetModifiers()
        selection = self.listbox.GetSelection()

        if keycode == wx.WXK_ESCAPE:
            self.onClose(event) 
        elif keycode == wx.WXK_RETURN:
            self.onPaste(event)
        elif keycode == wx.WXK_DELETE:  # Bind delete key
            self.deleteItem(self.listbox.GetSelection())
        elif keycode == wx.WXK_UP and modifiers == wx.MOD_SHIFT:
            self.moveItemUp(selection)
        elif keycode == wx.WXK_DOWN and modifiers == wx.MOD_SHIFT:
            self.moveItemDown(selection)
        elif keycode == wx.WXK_F2 and modifiers == wx.MOD_NONE:
            if selection != wx.NOT_FOUND:
                item_data = self.clipboard_history[selection]
                text = item_data.get('text', '') if isinstance(item_data, dict) else item_data
                if isinstance(text, str):
                    self.onEdit(selection, text)
                else:
                    # Translators: Message indicating that the selected item cannot be edited because it is not a text.
                    ui.message(_("This item cannot be edited because it is not a text."))
        elif keycode == ord('N') and modifiers == wx.MOD_CONTROL:
            self.onAdd(None)
        elif keycode == ord('F') and modifiers == wx.MOD_CONTROL:
            self.show_search_dialog()
        elif keycode == wx.WXK_F3:
            if self.search_results:
                if modifiers == wx.MOD_SHIFT:
                    self.show_previous_result()
                else:
                    self.show_next_result()
            else:
                self.show_search_dialog()
        else:
            event.Skip()

    def set_search_results(self, results, current_index):
        self.search_results = results
        self.current_result_index = current_index

    def show_next_result(self):
        if self.search_results:
            current_selection_index = self.listbox.GetSelection()
            next_result_index = -1

            # Find the next result index
            for i in range(len(self.search_results)):
                if self.search_results[i] > current_selection_index:
                    next_result_index = i
                    break

            if next_result_index == -1:
                # If no next result is found, and current selection is not the last result
                if current_selection_index != self.search_results[-1]:
                    next_result_index = 0
                else:
                    # Translators: Message indicating that there are no more results.
                    wx.MessageBox(_("No more results."), _("Search"), wx.OK | wx.ICON_INFORMATION, self)
                    return

            self.current_result_index = next_result_index
            self.listbox.SetSelection(self.search_results[self.current_result_index])
            self.listbox.SetFocus()
        else:
            self.show_search_dialog()

    def show_previous_result(self):
        if self.search_results:
            current_selection_index = self.listbox.GetSelection()
            previous_result_index = -1

            # Find the previous result index
            for i in range(len(self.search_results) - 1, -1, -1):
                if self.search_results[i] < current_selection_index:
                    previous_result_index = i
                    break

            if previous_result_index == -1:
                # If no previous result is found, and current selection is not the first result
                if current_selection_index != self.search_results[0]:
                    previous_result_index = len(self.search_results) - 1
                else:
                    # Translators: Message indicating that there are no previous results.
                    wx.MessageBox(_("No previous results."), _("Search"), wx.OK | wx.ICON_INFORMATION, self)
                    return

            self.current_result_index = previous_result_index
            self.listbox.SetSelection(self.search_results[self.current_result_index])
            self.listbox.SetFocus()
        else:
            self.show_search_dialog()

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
            if config.conf['clipManager']['showWarning']:
                # Translators: Confirmation message before deleting a clipboard item.
                if gui.messageBox(_("Are you sure you want to delete this item?"),
                                  # Translators: Title for the delete confirmation dialog.
                                  _("Confirm Deletion"),
                                  wx.YES_NO | wx.ICON_WARNING) == wx.NO:
                    return  # Cancel deletion

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

    def onAdd(self, event):
        """Opens the add dialog."""
        def on_save(text, pinned):
            # Add the new item to the clipboard history
            if pinned:
                new_item = {'text': text, 'pinned': True}
            else:
                new_item = {'text': text, 'pinned': False}
            self.clipboard_history.insert(0, new_item)

            # Update the listbox
            self.updateListBox()
            self.plugin_instance.saveHistory()

        add_dialog = AddDialog(self, on_save)
        add_dialog.ShowModal()

    def onImportHistory(self, event):
        """Handles the import history action."""
        wildcard = "JSON files (*.json)|*.json|" \
                   "All files (*.*)|*.*"
        # Translators: Title for the import history file dialog.
        dlg = wx.FileDialog(self, _("Choose a history file to import"),
                            wildcard=wildcard,
                            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return

        path = dlg.GetPath()
        dlg.Destroy()

        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    new_history = json.load(f)
                # Validate the imported history
                for item in new_history:
                    if not isinstance(item, dict) or 'text' not in item or 'pinned' not in item:
                        raise ValueError("Invalid history format")

                # Translators: Confirmation message before overwriting current history.
                if gui.messageBox(_("Importing this history will overwrite your current history. Are you sure you would like to continue?"),
                                  # Translators: Title for the import history confirmation dialog.
                                  _("Warning"),
                                  wx.YES_NO | wx.ICON_WARNING) == wx.YES:
                    self.clipboard_history[:] = new_history
                    self.updateListBox()
                    self.plugin_instance.saveHistory()
                    # Translators: Success message after importing history.
                    gui.messageBox(_("History imported successfully."), _("Success"), wx.OK | wx.ICON_INFORMATION)
            except (ValueError, json.JSONDecodeError):
                # Translators: Error message when the selected file is not a valid history file.
                ui.message(_("Invalid history file format."))
            except Exception as e:
                # Translators: Error message when importing history fails.
                ui.message(_("Error importing history: {}").format(e))

    def onExportHistory(self, event):
        """Handles the export history action."""
        wildcard = "JSON files (*.json)|*.json|" \
                   "All files (*.*)|*.*"
        # Translators: Title for the export history file dialog.
        dlg = wx.FileDialog(self, _("Save history as"),
                            wildcard=wildcard,
                            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return

        path = dlg.GetPath()
        dlg.Destroy()

        if path:
            if not path.lower().endswith(".json"):
                path += ".json"
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(self.clipboard_history, f, ensure_ascii=False, indent=4)
                # Translators: Success message after exporting history.
                gui.messageBox(_("Your clipboard History has been exported successfully."), _("Export Success"), wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                # Translators: Error message when exporting history fails.
                ui.message(_("Error exporting history: {}").format(e))

    def moveItemUp(self, selection):
        """Moves the selected item one position up in the list."""
        if selection > 0:
            item = self.clipboard_history.pop(selection)
            self.clipboard_history.insert(selection - 1, item)
            self.updateListBox()
            self.listbox.SetSelection(selection - 1)
            self.plugin_instance.saveHistory()

    def moveItemDown(self, selection):
        """Moves the selected item one position down in the list."""
        if selection < len(self.clipboard_history) - 1:
            item = self.clipboard_history.pop(selection)
            self.clipboard_history.insert(selection + 1, item)
            self.updateListBox()
            self.listbox.SetSelection(selection + 1)
            self.plugin_instance.saveHistory()

    def updateListBox(self):
        """Refreshes the listbox content."""
        self.listbox.Clear()
        for item in self.clipboard_history:
            text = item.get('text', '') if isinstance(item, dict) else item
            truncated_text = text[:config.conf['clipManager']['displayChars']]
            if len(text) > config.conf['clipManager']['displayChars']:
                truncated_text += " (See more)"
            self.listbox.Append(truncated_text)

    def onClose(self, event):
        if self.search_results:
            self.search_results = []  # Clear results on close
            self.current_result_index = -1
        self.Destroy()


class Settings(BaseSettingsPanel):
    # Translators: Title for the settings panel.
    title = _("Clipboard History Settings")

    def makeSettings(self, settingsSizer):
        sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

        # Translators: Label for the history size setting.
        sHelper.addItem(wx.StaticText(self, label=_("History size:")))
        self.historySizeSpin = sHelper.addItem(
            wx.SpinCtrl(
                self, -1, "", size=(60, -1), style=wx.SP_ARROW_KEYS, min=1, max=100000000
            )
        )
        self.historySizeSpin.SetValue(
            config.conf["clipManager"].get("historySize", 500)
        )
        self.historySizeSpin.Bind(wx.EVT_SPINCTRL, self.onSpinCtrlChanged)

        # Translators: Label for the 'save history on exit' checkbox.
        self.saveHistoryCheck = sHelper.addItem(
            wx.CheckBox(self, -1, _("Save history on exit"))
        )
        self.saveHistoryCheck.SetValue(
            config.conf["clipManager"].get("saveHistory", True)
        )
        self.saveHistoryCheck.Bind(wx.EVT_CHECKBOX, self.onCheckChanged)

        # Translators: Label for the save threshold setting.
        sHelper.addItem(wx.StaticText(self, label=_("Save threshold:")))
        self.saveThresholdSpin = sHelper.addItem(
            wx.SpinCtrl(self, -1, "", size=(60, -1), style=wx.SP_ARROW_KEYS, min=1, max=100)
        )
        self.saveThresholdSpin.SetValue(
            config.conf["clipManager"].get("saveThreshold", 5)
        )
        self.saveThresholdSpin.Bind(wx.EVT_SPINCTRL, self.onSaveThresholdChanged)

        # Translators: Label for the 'display characters per item' setting.
        sHelper.addItem(wx.StaticText(self, label=_("Display characters per item:")))
        self.displayCharsSpin = sHelper.addItem(
            wx.SpinCtrl(self, -1, "", size=(60, -1), style=wx.SP_ARROW_KEYS, min=1, max=50000)
        )
        self.displayCharsSpin.SetValue(
            config.conf["clipManager"].get("displayChars", 2000)
        )
        self.displayCharsSpin.Bind(wx.EVT_SPINCTRL, self.onDisplayCharsChanged)

        # Translators: Label for the 'show warning on delete' checkbox.
        self.showWarningCheck = sHelper.addItem(
            wx.CheckBox(self, -1, _("Show a warning dialog when deleting items"))
        )
        self.showWarningCheck.SetValue(
            config.conf["clipManager"].get("showWarning", True)
        )
        self.showWarningCheck.Bind(wx.EVT_CHECKBOX, self.onWarningCheckChanged)

        # Translators: Label for the 'move to top when pasting' checkbox.
        self.moveToTopCheck = sHelper.addItem(wx.CheckBox(self, -1, _("Move the item to the top of the list when pasting")))
        self.moveToTopCheck.SetValue(config.conf["clipManager"].get("moveToTop", True))
        self.moveToTopCheck.Bind(wx.EVT_CHECKBOX, self.onMoveToTopChanged)

        # Translators: Label for the 'move duplicate to top when copying' checkbox.
        self.moveDuplicateToTopCheck = sHelper.addItem(
            wx.CheckBox(self, -1, _("Move duplicate items to the top when copying"))
        )
        self.moveDuplicateToTopCheck.SetValue(
            config.conf["clipManager"].get("moveDuplicateToTop", False)
        )
        self.moveDuplicateToTopCheck.Bind(wx.EVT_CHECKBOX, self.onMoveDuplicateToTopChanged)

    def onSpinCtrlChanged(self, event):
        config.conf['clipManager']['historySize'] = self.historySizeSpin.GetValue()
    
    def onCheckChanged(self, event):
        config.conf['clipManager']['saveHistory'] = self.saveHistoryCheck.GetValue()

    def onSaveThresholdChanged(self, event):
        config.conf['clipManager']['saveThreshold'] = self.saveThresholdSpin.GetValue()

    def onDisplayCharsChanged(self, event):
        config.conf['clipManager']['displayChars'] = self.displayCharsSpin.GetValue()

    def onWarningCheckChanged(self, event):
        config.conf['clipManager']['showWarning'] = self.showWarningCheck.GetValue()

    def onMoveToTopChanged(self, event):
        config.conf['clipManager']['moveToTop'] = self.moveToTopCheck.GetValue()

    def onMoveDuplicateToTopChanged(self, event):
        config.conf['clipManager']['moveDuplicateToTop'] = self.moveDuplicateToTopCheck.GetValue()

    def onSave(self):
        config.conf.save()

clipManagerSection = "clipManager"
clipManagerSetting = {
    "historySize": "integer(default=500)",
    "saveHistory": "boolean(default=True)",
    "saveThreshold": "integer(default=5)",
    "displayChars": "integer(default=2000)",
    "showWarning": "boolean(default=True)",
    "moveToTop": "boolean(default=True)",
    "separatePinnedList": "boolean(default=False)",
    "moveDuplicateToTop": "boolean(default=False)",
}
config.conf.spec[clipManagerSection] = clipManagerSetting


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    clipboard_history = []
    dlg = None
    history_file = os.path.join(config.getUserDefaultConfigPath(), "clipManager", "history.json")
    copy_count = 0 # Counter for copy attempts
    menu_item = None

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Create the "clip_manager" directory if it doesn't exist
        clip_manager_dir = os.path.dirname(self.history_file)
        if not os.path.exists(clip_manager_dir):
            os.makedirs(clip_manager_dir)
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
        self.tools_menu = gui.mainFrame.sysTrayIcon.toolsMenu
        # Translators: Menu item text.  '&' makes 'h' the access key.
        self.menu_item = self.tools_menu.Append(
            wx.ID_ANY, _("&Clipboard history")
        )
        self.tools_menu.Bind(wx.EVT_MENU, self.showClipboardHistory_menu, self.menu_item)


    def terminate(self):
        self.saveHistory()
        NVDASettingsDialog.categoryClasses.remove(Settings)
        ctypes.windll.user32.RemoveClipboardFormatListener(self.hwnd)
        ctypes.windll.user32.DestroyWindow(self.hwnd)
        ctypes.windll.user32.UnregisterClassW(self.wndclass.lpszClassName, self.hinst)
        if self.menu_item:
            self.tools_menu.Delete(self.menu_item.GetId())
            self.menu_item = None

    def wndProc(self, hwnd, msg, wparam, lparam):
        if msg == WM_CLIPBOARDUPDATE:
            self.onClipboardUpdate()
        return ctypes.windll.user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def clearClipboardHistory(self):
        """Clears the clipboard history list."""
        self.clipboard_history = [item for item in self.clipboard_history if isinstance(item, dict) and 'pinned' in item and item['pinned']]

    @script(
        # Translators: Description of the script that shows the clipboard history.
        description=_("Shows clipboard history"),
        gesture="kb:NVDA+v",
        category=_("clip_manager")
    )
    def script_showClipboardHistory(self, gesture):
        if not self.clipboard_history:
            # Translators: Message shown when the clipboard history is empty.
            ui.message(_("Clipboard history is empty."))
            return

        obj=api.getFocusObject()
        mainFrame.prePopup()
        self.dlg = ClipboardHistoryFrame(mainFrame, self.clipboard_history, obj, self.clearClipboardHistory, self)
        self.dlg.CentreOnScreen()
        self.dlg.Show(True)
        mainFrame.postPopup()

    def onClipboardUpdate(self):
        try:
            new_text = pyperclip.paste()
            # Efficiently check if the new text is already present, considering both string and dictionary formats
            move_to_top = config.conf['clipManager']['moveDuplicateToTop']
            existing_index = -1
            for i, item in enumerate(self.clipboard_history):
                if (isinstance(item, dict) and item.get('text') == new_text) or (isinstance(item, str) and item == new_text):
                    existing_index = i
                    break

            if new_text and existing_index != -1 and move_to_top:
                # Move existing item to top
                existing_item = self.clipboard_history.pop(existing_index)
                if isinstance(existing_item, str):
                    existing_item = {'text': existing_item, 'pinned': False}
                self.clipboard_history.insert(0, existing_item)
            elif new_text and not any(item == new_text or (isinstance(item, dict) and item.get('text') == new_text) for item in self.clipboard_history):
                # Add new items as dictionaries for consistent handling of pinned status
                self.clipboard_history.insert(0, {'text': new_text, 'pinned': False})
                if len(self.clipboard_history) > config.conf['clipManager']['historySize']:
                    # Remove unpinned items from the end when exceeding history size
                    self.clipboard_history = [item for item in self.clipboard_history if isinstance(item, dict) and item.get('pinned')] + self.clipboard_history[:config.conf['clipManager']['historySize']]
                self.copy_count += 1
                if self.copy_count >= config.conf['clipManager']['saveThreshold']:
                    self.saveHistory()
                    self.copy_count = 0

        except Exception as e:
            # Translators: Message shown when getting clipboard text fails.
            ui.message(_("Failed to get clipboard text: {}").format(e))

    def loadHistory(self):
        """Loads clipboard history from file if saveHistory is enabled."""
        if config.conf['clipManager'].get("saveHistory", True) and os.path.exists(self.history_file):
            try:
                with open(self.history_file, "r", encoding="utf-8") as f:
                    loaded_history = json.load(f)

                # Ensure all loaded items are in dict format
                self.clipboard_history = []
                for item in loaded_history:
                    if isinstance(item, dict):
                        self.clipboard_history.append(item)
                    elif isinstance(item, str):
                        self.clipboard_history.append({'text': item, 'pinned': False})

            except json.JSONDecodeError:
                # Translators: Error message when loading history fails due to invalid JSON format.
                ui.message(_("Error loading clipboard history: Invalid JSON format."))
            except Exception as e:
                # Translators: Error message when loading history fails for any other reason.
                ui.message(_("Error loading clipboard history: {}").format(e))

    def saveHistory(self):
        """Saves clipboard history to file if saveHistory is enabled."""
        if config.conf['clipManager'].get("saveHistory", True):
            try:
                with open(self.history_file, "w", encoding="utf-8") as f:
                    json.dump(self.clipboard_history, f, ensure_ascii=False)
            except:
                # Translators: Error message when saving history fails.
                ui.message(_("Error saving clipboard history!"))

    def showClipboardHistory_menu(self, event): # Add this
        """Shows the clipboard history dialog."""
        if not self.clipboard_history:
            # Translators: Message shown when the clipboard history is empty.
            ui.message(_("Clipboard history is empty."))
            return

        obj = api.getFocusObject()
        mainFrame.prePopup()
        self.dlg = ClipboardHistoryFrame(mainFrame, self.clipboard_history, obj, self.clearClipboardHistory, self)
        self.dlg.CentreOnScreen()
        self.dlg.Show(True)
        mainFrame.postPopup()