import wx


APP_EXIT = 1
VIEW_STATUS = 2
VIEW_RGB = 3
VIEW_SRGB = 4
IT_MIN = 5
IT_MAX = 6
ID_BTN = 25


class AppContextMenu(wx.Menu):
    def __init__(self, parent):
        self.parent = parent
        super().__init__()

        self.Append(IT_MIN, 'Минимизировать')
        self.Append(IT_MAX, 'Распахнуть')
        self.Bind(wx.EVT_MENU, self.onMinimize, id=IT_MIN)
        self.Bind(wx.EVT_MENU, self.onMaximize, id=IT_MAX)

    def onMinimize(self, event):
            self.parent.Iconize()

    def onMaximize(self, event):
            self.parent.Maximize()


class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title)

        menuber = wx.MenuBar()
        fileMenu = wx.Menu()
        expMenu = wx.Menu()

        expMenu.Append(wx.ID_ANY, 'Экспорт изображения')
        expMenu.Append(wx.ID_ANY, 'Экспорт видео')
        expMenu.Append(wx.ID_ANY, 'Экспорт данных')
        # item = wx.MenuItem(fileMenu, wx.ID_EXIT, 'Выход', 'Выход из приложения')
        # fileMenu.Append(item)
        fileMenu.Append(wx.ID_NEW, 'Новый\tCtrl+N')
        fileMenu.Append(wx.ID_OPEN, 'Открыть\tCtrl+O')
        fileMenu.Append(wx.ID_SAVE, 'Сохранить\tCtrl+S')
        fileMenu.AppendSubMenu(expMenu, '&Экспорт')
        fileMenu.AppendSeparator()
        fileMenu.Append(APP_EXIT, 'Выход\tCtrl+Q', 'Выход из приложения')

        viewMenu = wx.Menu()
        viewMenu.Append(VIEW_STATUS, 'Статусная строка', kind=wx.ITEM_CHECK)
        viewMenu.Append(VIEW_RGB, 'Тип RGB', 'Тип RGB', kind=wx.ITEM_RADIO)
        viewMenu.Append(VIEW_SRGB, 'Тип sRGB', 'Тип sRGB', kind=wx.ITEM_RADIO)

        menuber.Append(fileMenu, '&File')
        menuber.Append(viewMenu, '&Вид')
        self.SetMenuBar(menuber)

        self.Bind(wx.EVT_MENU, self.onQuit, id=APP_EXIT)
        self.Bind(wx.EVT_MENU, self.onStatus, id=VIEW_STATUS)
        self.Bind(wx.EVT_MENU, self.onImageType, id=VIEW_RGB)
        self.Bind(wx.EVT_MENU, self.onImageType, id=VIEW_SRGB)

        self.ctx = AppContextMenu(self)
        self.Bind(wx.EVT_RIGHT_DOWN, self.onRigntDown)

        toolbar = self.CreateToolBar()
        toolbar.SetBackgroundColour('#FFAFEF')
        toolbar.AddTool(10, 'Выход', wx.Bitmap('gphoto.png'), shortHelp='Exit', kind=wx.ITEM_NORMAL)
        toolbar.AddSeparator()
        toolbar.AddTool(wx.ID_UNDO, "", wx.Bitmap('gnomebaker.png'))
        toolbar.AddTool(wx.ID_REDO, "", wx.Bitmap('pitivi.png'))
        toolbar.AddSeparator()
        toolbar.AddCheckTool(wx.ID_ANY, "", wx.Bitmap('bittorrent.png'))
        toolbar.AddSeparator()
        toolbar.AddRadioTool(wx.ID_ANY, "", wx.Bitmap('bittorrent.png'))
        toolbar.AddRadioTool(wx.ID_ANY, "", wx.Bitmap('bittorrent.png'))
        toolbar.EnableTool(wx.ID_REDO, False)
        toolbar.Realize()

        self.Bind(wx.EVT_TOOL, self.onQuit, id=10)

        panel = wx.Panel(self)
        font = wx.SystemSettings.GetFont(wx.SYS_DEFAULT_GUI_FONT)
        font.SetPointSize(12)
        panel.SetFont(font)
        vbox = wx.BoxSizer(wx.VERTICAL)

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        st1 = wx.StaticText(panel, label='Путь к файлу:')
        tc = wx.TextCtrl(panel)
        hbox1.Add(st1, flag=wx.RIGHT, border=8)
        hbox1.Add(tc, proportion=1)
        vbox.Add(hbox1, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP, border=10)

        st2 = wx.StaticText(panel, label='Содержимое файла')
        vbox.Add(st2, flag=wx.EXPAND|wx.ALL, border=10)

        tc2 = wx.TextCtrl(panel, style=wx.TE_MULTILINE)
        vbox.Add(tc2, proportion=1, flag=wx.LEFT|wx.RIGHT|wx.BOTTOM|wx.EXPAND, border=10)

        btnOk = wx.Button(panel, label='Да', size=(70, 30))
        # btnCn = wx.Button(panel, label='Отмена', size=(70, 30), id=ID_BTN)
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        hbox2.Add(btnOk, flag=wx.LEFT, border=10)
        hbox2.Add(wx.Button(panel, id=ID_BTN, label='Отмена', size=(70, 30)))
        vbox.Add(hbox2, flag=wx.ALIGN_RIGHT|wx.BOTTOM|wx.RIGHT, border=10)

        panel.SetSizer(vbox)


    def onRigntDown(self, event):
        self.PopupMenu(self.ctx, event.GetPosition())

    def onQuit(self, event):
        self.Close()

    def onStatus(self, event):
        if event.IsChecked():
            print('Показать статусную строку')
        else:
            print('Скрыть статусную строку')

    def onImageType(self, event):
        if event.GetId() == VIEW_RGB:
            print('Режим RGB')
        elif event.GetId() == VIEW_SRGB:
            print('Режим sRGB')


app = wx.App()
frame = MyFrame(None, title='Hello World!').Show()


app.MainLoop()
