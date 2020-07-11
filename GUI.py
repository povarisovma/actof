import wx
import docxfilemaker
import win32com.client
import re
import os
import getlistfiles
import wx.lib.mixins.listctrl
from ObjectListView import ObjectListView, ColumnDefn

ID_BTN_CRARCT = 15
ID_BTN_SENDACT = 25
ID_BTN_DELACT = 26
ID_BTN_REFLACT = 27
ID_LC_ACTLIST = 35
ID_MB_EXIT = 41
ID_MB_OPENDOCX = 42
ID_MB_SETTINGS = 43
ID_MB_OPENFOLDERLOCAL = 51
ID_MB_OPENFOLDERREPO = 52
ID_MB_ABOUT = 61

ID_MD_CHOSDIRLOC = 105
ID_MD_CHOSDIR = 106
ID_MD_CHOSTMPL = 107
ID_MD_PATHDIRACTLOC = 111
ID_MD_PATHDIRACT = 112


class MyDlg(wx.Dialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.SetSize(500, 500)

        #Обьявление главного сайзера:
        self.mainsizer = wx.BoxSizer(wx.VERTICAL)

        #Блок 1, виджеты для указания пути к папке локальных актов--------------------------------------------------
        #Добавление статического текста "Путь к папке с локальными актами" в сайзер:
        self.mainsizer.Add(wx.StaticText(self, wx.ID_ANY, label="Путь к папке с локальными актами:"),
                           flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        #Создание горизонтального сайзера для добавления поля ввода и кнопки для выбора папки
        # а также добавление его в главный сайзер:
        self.folderactssizer = wx.BoxSizer(wx.HORIZONTAL)
        self.mainsizer.Add(self.folderactssizer, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        #Создание текстового поля для ввода пути, а также его добавление в горизонтальный сайзер:
        self.tc_actloc_path = wx.TextCtrl(self, id=ID_MD_PATHDIRACTLOC, value="C:\\")
        self.folderactssizer.Add(self.tc_actloc_path, proportion=1)
        #Создание кнопки для открытия диалогового окна выбора папки, и добавление его в горизонтальный сайзер:
        self.folderactssizer.Add(wx.Button(self, id=ID_MD_CHOSDIRLOC, label='...'), flag=wx.EXPAND | wx.LEFT, border=10)

        # Блок 2, виджеты для указания пути к папке общих актов------------------------------------------------------
        # Добавление статического текста "Путь к папке с общими актами" в сайзер:
        self.mainsizer.Add(wx.StaticText(self, wx.ID_ANY, label="Путь к папке с общими актами:"),
                           flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        # Создание горизонтального сайзера для добавления поля ввода и кнопки для выбора папки
        # а также добавление его в главный сайзер:
        self.folderactssizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.mainsizer.Add(self.folderactssizer2, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        # Создание текстового поля для ввода пути, а также его добавление в горизонтальный сайзер:
        self.tc_act_path = wx.TextCtrl(self, id=ID_MD_PATHDIRACT, value="C:\\")
        self.folderactssizer2.Add(self.tc_act_path, proportion=1)
        # Создание кнопки для открытия диалогового окна выбора папки, и добавление его в горизонтальный сайзер:
        self.folderactssizer2.Add(wx.Button(self, id=ID_MD_CHOSDIR, label='...'), flag=wx.EXPAND | wx.LEFT, border=10)
        #TODO придумать и реализовать указание шаблона, пока не понятно как лучше это отразить и вообще надо ли?
        #Блок 3, виджеты для указания пути к шаблону docx
        self.mainsizer.Add(wx.StaticText(self, wx.ID_ANY, label="Путь к docx шаблону:"),
                           flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)

        self.folderactssizer3 = wx.BoxSizer(wx.HORIZONTAL)
        self.mainsizer.Add(self.folderactssizer3, flag=wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, border=10)
        self.folderactssizer3.Add(wx.TextCtrl(self, wx.ID_ANY, value="C:\\"), proportion=1)
        self.folderactssizer3.Add(wx.Button(self, id=ID_MD_CHOSTMPL, label='...'), flag=wx.EXPAND | wx.LEFT, border=10)


        #Добавление главного сайзера с виджетами в окно
        self.SetSizer(self.mainsizer)

        #Назначение кнопок и функций
        self.Bind(wx.EVT_BUTTON, self.choosediractsloc, id=ID_MD_CHOSDIRLOC)
        self.Bind(wx.EVT_BUTTON, self.choosediracts, id=ID_MD_CHOSDIR)
        self.Bind(wx.EVT_BUTTON, self.choosetempldocx, id=ID_MD_CHOSTMPL)

    def choosediractsloc(self, event):
        print("press choose directory loc acts", event.GetId())
        dlg = wx.DirDialog(self, message="Выберете папку расположения локальных актов", defaultPath=self.tc_actloc_path.GetValue())
        res = dlg.ShowModal()
        if res == wx.ID_OK:
            print(dlg.GetPath())
            self.tc_actloc_path.SetValue(dlg.GetPath())


    def choosediracts(self, event):
        print("press choose directory acts", event.GetId())
        dlg1 = wx.DirDialog(self, message="Hello", defaultPath=self.tc_act_path.GetValue())
        res = dlg1.ShowModal()
        if res == wx.ID_OK:
            print(dlg1.GetPath())
            self.tc_act_path.SetValue(dlg1.GetPath())

    def choosetempldocx(self, event):
        print("press choose docx template", event.GetId())


class MyFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title='ActOf', size=wx.Size(1037, 605),
                          style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        #Создание меню
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        refMenu = wx.Menu()
        ftMenu = wx.Menu()

        fileMenu.Append(ID_MB_SETTINGS, "Настройки", "Открытие окна настроек")
        fileMenu.AppendSeparator()
        fileMenu.Append(ID_MB_EXIT, "Выход", "Выход из приложения")

        ftMenu.Append(ID_MB_OPENFOLDERLOCAL, "Открыть папку локальных актов")
        ftMenu.Append(ID_MB_OPENFOLDERREPO, "Открыть папку актов")
        ftMenu.Append(ID_MB_OPENDOCX, "Открыть шаблон docx")

        refMenu.Append(ID_MB_ABOUT, "О программе", "Описание программы")

        menubar.Append(fileMenu, "Файл")
        menubar.Append(ftMenu, "Навигация")
        menubar.Append(refMenu, "Справка")
        self.SetMenuBar(menubar)

        self.Bind(wx.EVT_MENU, self.onSettings, id=ID_MB_SETTINGS)
        self.Bind(wx.EVT_MENU, self.onQuit, id=ID_MB_EXIT)
        self.Bind(wx.EVT_MENU, self.openFolderLocalActs, id=ID_MB_OPENFOLDERLOCAL)
        self.Bind(wx.EVT_MENU, self.openFolderActs, id=ID_MB_OPENFOLDERREPO)
        self.Bind(wx.EVT_MENU, self.openDocxTemplate, id=ID_MB_OPENDOCX)
        self.Bind(wx.EVT_MENU, self.about, id=ID_MB_ABOUT)


        #Объявление сайзеров------------------------------------------------------------------------
        #Главный сайзер программы:
        self.mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        #Центральная часть(CA = CREATE ACTS): сайзер для создания актов:
        self.centrSCreateActs = wx.BoxSizer(wx.VERTICAL)
        #Правая часть: Сайзеры для отправки актов(горизонтальный для добавления кнопок):
        self.rightSSendActs = wx.BoxSizer(wx.VERTICAL)
        self.rightSSendActsTopBTN = wx.BoxSizer(wx.HORIZONTAL)

        #Центральная часть программы, поле для ввода текста и кнопка создать акт:
        #Основное окно, создание кнопки "Создать Акт" и поля для текстового ввода:
        self.BTNCreateActCS = wx.Button(self, id=ID_BTN_CRARCT, label=u"Создать Акт")
        self.TCTextInputCS = wx.TextCtrl(self,
                                         wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_MULTILINE)
        # Добавление кнопки "Создать Акт" и поле текстового ввода в центральный сайзер:
        self.centrSCreateActs.Add(self.BTNCreateActCS, 0, wx.TOP | wx.LEFT | wx.RIGHT | wx.EXPAND, 5)
        self.centrSCreateActs.Add(self.TCTextInputCS, 1, wx.ALL | wx.EXPAND, 5)
        #Добавление центрального сайзера в главный сайзер:
        self.mainSizer.Add(self.centrSCreateActs, proportion=1, flag=wx.EXPAND, border=5)

        #Правая часть программы, список созданных актов и кнопки для работы с ними:
        #Создание и добавление кнопок в сайзер:
        self.send_actBTN = wx.Button(self, id=ID_BTN_SENDACT, label="Отправить Акт")
        self.del_actBTN = wx.Button(self, id=ID_BTN_DELACT, label="Удалить файлы")
        self.refresh_lactsBTN = wx.Button(self, id=ID_BTN_REFLACT, label="Обновить список")
        self.rightSSendActsTopBTN.Add(self.send_actBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)
        self.rightSSendActsTopBTN.Add(self.del_actBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)
        self.rightSSendActsTopBTN.Add(self.refresh_lactsBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)

        self.rightSSendActs.Add(self.rightSSendActsTopBTN, proportion=0, flag=wx.EXPAND)

        #Создание списка актов
        self.OLVlocal_acts = ObjectListView(self, wx.ID_ANY, style=wx.LC_REPORT | wx.SUNKEN_BORDER)
        #Создание столбцов
        title = ColumnDefn("Title", "left", 220, "title", isSpaceFilling=False)
        creating = ColumnDefn("Date Creating", "left", 130, "creating",  stringConverter="%d-%m-%Y %H:%M:%S",
                              isSpaceFilling=False)
        modifine = ColumnDefn("Date Modifine", "left", 130, "modifine",  stringConverter="%d-%m-%Y %H:%M:%S",
                              isSpaceFilling=False)
        self.OLVlocal_acts.oddRowsBackColor = wx.WHITE
        self.OLVlocal_acts.SetColumns([title, creating, modifine])
        #Добавление в список актов из папки locals_act
        self.OLVlocal_acts.SetObjects(getlistfiles.getDictFilesParam())
        #Добавление списка актов в сайзер
        self.rightSSendActs.Add(self.OLVlocal_acts, proportion=1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)

        #Добавление правого сайзера в главный сайзер.
        self.mainSizer.Add(self.rightSSendActs, proportion=0, flag=wx.EXPAND)

        #Назначение главного сайзера
        self.SetSizer(self.mainSizer)

        #Назначение событий
        self.Bind(wx.EVT_BUTTON, self.createActOn, id=ID_BTN_CRARCT)
        self.Bind(wx.EVT_BUTTON, self.sendActOn, id=ID_BTN_SENDACT)
        self.Bind(wx.EVT_BUTTON, self.del_acts_action, id=ID_BTN_DELACT)
        self.Bind(wx.EVT_BUTTON, self.refresh_list_acts, id=ID_BTN_REFLACT)

    def onSettings(self, event):
        print("open settings")
        with MyDlg(self, title="Настройки") as dlg:
            res = dlg.ShowModal()

        print(res, wx.ID_CANCEL)

    def onQuit(self, event):
        self.Close()

    def openFolderLocalActs(self, event):
        print("open folder local acts")

    def openFolderActs(self, event):
        print("open folder acts")

    def openDocxTemplate(self, event):
        print("open docx template")

    def about(self, event):
        print("open about")

    def del_acts_action(self, event):
        if event.GetId() == ID_BTN_DELACT:
            selection = self.OLVlocal_acts.GetSelectedObjects()
            if selection:
                dlg = wx.MessageBox('Удалить выбранные файлы?', 'Подтверждение', wx.YES_NO | wx.NO_DEFAULT, self)
                if dlg == wx.YES:
                    for i in range(len(selection)):
                        pathdocx = os.path.abspath('local_acts') + '\\' + selection[i]['title'].rstrip('pdf') + 'docx'
                        pathpdf = os.path.abspath('local_acts') + '\\' + selection[i]['title']
                        os.remove(pathdocx)
                        os.remove(pathpdf)
                        self.refresh_list_acts(event)


    def refresh_list_acts(self, event):
        self.OLVlocal_acts.SetObjects(getlistfiles.getDictFilesParam())

    def createActOn(self, event):
        if event.GetId() == ID_BTN_CRARCT and self.TCTextInputCS.GetNumberOfLines() > 0:
            txtlst = list(map(lambda x: x.strip(), self.TCTextInputCS.GetValue().split('\n')))
            docxfilemaker.createdocxnpdffiles(txtlst)
            self.refresh_list_acts(event)

    def sendActOn(self, event):
        if event.GetId() == ID_BTN_SENDACT:
            selection = self.OLVlocal_acts.GetSelectedObjects()
            if selection:
                theme = 'АЗС ' + re.sub('\\D', '', selection[0]['title'].split('_')[1]) + ' ССО '
                themeset = set()
                bodiez = 'Доброго времени суток.<br />'
                if len(selection) == 1:
                    bodiez += 'Высылаю акт '
                if len(selection) > 1:
                    bodiez += 'Высылаю акты '
                for i in range(len(selection)):
                    themeset.add(re.sub('\\D', '', selection[i]['title'].split('_')[2]))
                    bodiez += re.sub('\\D', '', selection[i]['title'].split('_')[0]) + ', '
                for z in themeset:
                    theme += z + ' '
                bodiez = bodiez.rstrip(', ') + '.'
                print('Создание письма')
                app = win32com.client.Dispatch("Outlook.Application")
                mess = app.CreateItem(0)
                mess.Subject = theme
                mess.GetInspector()
                index = mess.HTMLbody.find('>', mess.HTMLbody.find('<body'))
                mess.HTMLbody = mess.HTMLbody[:index + 1] + bodiez + mess.HTMLbody[index + 1:]
                for i in range(len(selection)):
                    path = os.path.abspath('local_acts') + '\\' + selection[i]['title']
                    mess.Attachments.Add(path)
                print('Отправка письма')
                mess.Display(True)
                print('Письмо отправлено')


def main():
    app = wx.App()
    frame = MyFrame(None).Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
