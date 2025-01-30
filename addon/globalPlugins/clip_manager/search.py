import re
import wx
import ui

class SearchDialog(wx.Dialog):
    def __init__(self, parent, history):
        # Translators: Title for the search dialog
        super(SearchDialog, self).__init__(parent, title=_("Find"), size=(400, 200))
        self.history = history
        self.results = []
        self.current_result_index = -1
        self.parent = parent

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Translators: Label for the search term text control
        search_label = wx.StaticText(panel, label=_("Search term:"))
        sizer.Add(search_label, 0, wx.ALL, 5)

        self.search_ctrl = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER)
        self.search_ctrl.Bind(wx.EVT_TEXT_ENTER, self.on_search)
        sizer.Add(self.search_ctrl, 0, wx.EXPAND | wx.ALL, 5)

        # Translators: Label for the 'Match case' checkbox
        self.match_case_checkbox = wx.CheckBox(panel, label=_("Match case"))
        sizer.Add(self.match_case_checkbox, 0, wx.ALL, 5)

        # Translators: Label for the 'Regular expressions' checkbox
        self.regex_checkbox = wx.CheckBox(panel, label=_("Regular expressions"))
        sizer.Add(self.regex_checkbox, 0, wx.ALL, 5)

        # Buttons
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        # Translators: Label for the 'OK' button in the search dialog
        ok_button = wx.Button(panel, wx.ID_OK, label=_("OK"))
        ok_button.Bind(wx.EVT_BUTTON, self.on_search)
        button_sizer.Add(ok_button, 0, wx.ALL, 5)
        # Translators: Label for the 'Cancel' button in the search dialog
        cancel_button = wx.Button(panel, wx.ID_CANCEL, label=_("Cancel"))
        button_sizer.Add(cancel_button, 0, wx.ALL, 5)
        sizer.Add(button_sizer, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        panel.SetSizer(sizer)

    def on_search(self, event):
        search_term = self.search_ctrl.GetValue()
        match_case = self.match_case_checkbox.GetValue()
        use_regex = self.regex_checkbox.GetValue()

        self.results = []
        self.current_result_index = -1
        if not search_term:
            return

        try:
            if use_regex:
                flags = 0 if match_case else re.IGNORECASE
                pattern = re.compile(search_term, flags)
                for i, item in enumerate(self.history):
                    text = item.get("text", "") if isinstance(item, dict) else item
                    if pattern.search(text):
                        self.results.append(i)
            else:
                search_term_lower = search_term.lower()
                for i, item in enumerate(self.history):
                    text = item.get("text", "") if isinstance(item, dict) else item
                    if (match_case and search_term in text) or (
                        not match_case and search_term_lower in text.lower()
                    ):
                        self.results.append(i)

            if self.results:
                self.current_result_index = 0
                self.move_focus_to_result()
                self.parent.set_search_results(self.results, self.current_result_index)
                self.EndModal(wx.ID_OK)
            else:
                # Translators: Message indicating that no matching results were found.
                wx.MessageBox(_("No matching results found."), _("Search"), wx.OK | wx.ICON_INFORMATION, self)

        except re.error:
            # Translators: Message indicating that the regular expression is invalid.
            wx.MessageBox(_("Invalid regular expression."), _("Error"), wx.OK | wx.ICON_ERROR, self)

    def move_focus_to_result(self):
        if 0 <= self.current_result_index < len(self.results):
            index = self.results[self.current_result_index]
            if self.parent:
                self.parent.listbox.SetSelection(index)
                self.parent.listbox.SetFocus()