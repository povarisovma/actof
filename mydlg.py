import wx
import settings


ID_MD_CHOSDIRLOC = 105
ID_MD_CHOSDIR = 106
ID_MD_CHOSTMPL = 107
ID_MD_PATHDIRACTLOC = 111
ID_MD_PATHDIRACT = 112


class MyDlg(wx.Dialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        #Обьявление главного сайзера:
        self.mainsizer = wx.BoxSizer(wx.VERTICAL)

        #Блок 1, виджеты для указания пути к папке локальных актов--------------------------------------------------
        self.mainsizer.Add(wx.StaticText(self, wx.ID_ANY, label="Путь к папке с локальными актами:"),
                           flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.folderactssizer = wx.BoxSizer(wx.HORIZONTAL)
        self.mainsizer.Add(self.folderactssizer, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.tc_actloc_path = wx.TextCtrl(self, id=ID_MD_PATHDIRACTLOC, value=settings.get_local_acts_path_folder())
        self.folderactssizer.Add(self.tc_actloc_path, proportion=1)
        self.folderactssizer.Add(wx.Button(self, id=ID_MD_CHOSDIRLOC, label='...'), flag=wx.EXPAND | wx.LEFT, border=10)

        # Блок 2, виджеты для указания пути к папке общих актов------------------------------------------------------
        self.mainsizer.Add(wx.StaticText(self, wx.ID_ANY, label="Путь к папке с общими актами:"),
                           flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.folderactssizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.mainsizer.Add(self.folderactssizer2, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.tc_act_path = wx.TextCtrl(self, id=ID_MD_PATHDIRACT, value=settings.get_general_acts_path_folder())
        self.folderactssizer2.Add(self.tc_act_path, proportion=1)
        self.folderactssizer2.Add(wx.Button(self, id=ID_MD_CHOSDIR, label='...'), flag=wx.EXPAND | wx.LEFT, border=10)

        #Добавление главного сайзера с виджетами в окно
        self.SetSizer(self.mainsizer)

        #Назначение кнопок и функций
        self.Bind(wx.EVT_BUTTON, self.choosediractsloc, id=ID_MD_CHOSDIRLOC)
        self.Bind(wx.EVT_BUTTON, self.choosediracts, id=ID_MD_CHOSDIR)

    def choosediractsloc(self, event):
        dlg = wx.DirDialog(self, message="Выберите папку расположения локальных актов", defaultPath=self.tc_actloc_path.GetValue())
        res = dlg.ShowModal()
        if res == wx.ID_OK:
            print(dlg.GetPath())
            self.tc_actloc_path.SetValue(dlg.GetPath())
            settings.set_local_acts_path_folder_in_settings(dlg.GetPath())


    def choosediracts(self, event):
        dlg1 = wx.DirDialog(self, message="Выберите папку расположения общих актов", defaultPath=self.tc_act_path.GetValue())
        res = dlg1.ShowModal()
        if res == wx.ID_OK:
            print(dlg1.GetPath())
            self.tc_act_path.SetValue(dlg1.GetPath())
            settings.set_general_acts_path_folder_in_settings(dlg1.GetPath())