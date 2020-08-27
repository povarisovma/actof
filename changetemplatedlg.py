import wx


class ChangeTemplDlg(wx.Dialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        #Объявление главного сайзера и создание панели:
        self.mainsizer = wx.BoxSizer(wx.VERTICAL)
        self.panel = wx.Panel(self)
        


        #Подключение главного сайзера к панели:
        self.panel.SetSizer(self.mainsizer)