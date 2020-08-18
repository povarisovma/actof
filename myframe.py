import wx
import fileProcessing

import win32com.client
import subprocess
import re
import os
import settings
import mydlg
import aboutdlg
from ObjectListView import ObjectListView, ColumnDefn

ID_BTN_REFALL = 16
ID_BTN_CRARCT = 15
ID_BTN_CLEAR = 17
ID_BTN_SENDACT = 25
ID_BTN_DELACT = 26
ID_BTN_REFLACT = 27
ID_BTN_OPENDOCX = 28
ID_BTN_RECOPYACT = 29
ID_BTN_COPYPDFBUFF = 30
ID_LC_ACTLIST = 35
ID_MB_EXIT = 41
ID_MB_OPENDOCX = 42
ID_MB_SETTINGS = 43
ID_MB_OPENFOLDERLOCAL = 51
ID_MB_OPENFOLDERREPO = 52
ID_MB_ABOUT = 61
ID_BTN_ACTNUM = 71
ID_TC_ACTNUM = 72
ID_BTN_AZSNUM = 73
ID_TC_AZSNUM = 74
ID_BTN_SSONUM = 75
ID_TC_SSONUM = 76
ID_BTN_DATEINP = 77
ID_TC_DATEINP = 78


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

        #Назначение функций кнопкам меню:
        self.Bind(wx.EVT_MENU, self.onSettings, id=ID_MB_SETTINGS)
        self.Bind(wx.EVT_MENU, self.onQuit, id=ID_MB_EXIT)
        self.Bind(wx.EVT_MENU, self.openFolderLocalActs, id=ID_MB_OPENFOLDERLOCAL)
        self.Bind(wx.EVT_MENU, self.openFolderActs, id=ID_MB_OPENFOLDERREPO)
        self.Bind(wx.EVT_MENU, self.openDocxTemplate, id=ID_MB_OPENDOCX)
        self.Bind(wx.EVT_MENU, self.about, id=ID_MB_ABOUT)


        #Создание панели:
        panel = wx.Panel(self)
        self.font = wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, False, u'Montserrat')
        panel.SetFont(self.font)

        #Объявление сайзеров------------------------------------------------------------------------
        #Главный сайзер программы:
        self.mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        #Центральная часть сайзеры для создания актов:
        self.centrSCreateActs = wx.BoxSizer(wx.VERTICAL)
        self.centrSCreateActsGeneral = wx.BoxSizer(wx.HORIZONTAL)
        self.centrSActNumDef = wx.BoxSizer(wx.HORIZONTAL)
        self.centrSAZSSSONumDef = wx.BoxSizer(wx.HORIZONTAL)
        #Правая часть: Сайзеры для отправки актов(горизонтальный для добавления кнопок):
        self.rightSSendActs = wx.BoxSizer(wx.VERTICAL)
        self.rightSSendActsTopBTN = wx.BoxSizer(wx.HORIZONTAL)
        self.rSizerfileworkBTN = wx.BoxSizer(wx.HORIZONTAL)


        #Центральная часть программы, поле для ввода текста и кнопка создать акт:
        #Создание кнопок "Обновить всё", "Создать Акт", "Очистить":
        self.BTNrefreshAll = wx.Button(panel, id=ID_BTN_REFALL, label="Обновить всё")
        self.BTNCreateActCS = wx.Button(panel, id=ID_BTN_CRARCT, label=u"Создать Акт")
        self.BTNclearAll = wx.Button(panel, id=ID_BTN_CLEAR, label="Очистить")
        #Определение шрифтов:
        self.fontSTactNum = wx.Font(15, wx.MODERN, wx.NORMAL, wx.BOLD, False, u'Times New Roman')
        self.fontBodyAct = wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, False, u'Times New Roman')
        #Создание заголовка "Акт №" + поле ввода акта + кнопка обновить:
        self.BTNactNum = wx.Button(panel, ID_BTN_ACTNUM, label="Акт №")
        self.BTNactNum.SetFont(self.fontSTactNum)

        self.TCActNumDef = wx.TextCtrl(panel, ID_TC_ACTNUM, wx.EmptyString, wx.DefaultPosition, size=(85, -1))
        self.TCActNumDef.SetFont(self.fontSTactNum)

        #Создание шапки документа номер АЗС номер ССО по аналогии с номером акта, текущая дата:
        self.BTNAZSnum = wx.Button(panel, ID_BTN_AZSNUM, label="АЗС №")
        self.BTNAZSnum.SetFont(self.fontBodyAct)

        self.TCAZSNumDef = wx.TextCtrl(panel, ID_TC_AZSNUM, wx.EmptyString, wx.DefaultPosition, size=(65, -1))
        self.TCAZSNumDef.SetFont(self.fontBodyAct)

        self.BTNSSOnum = wx.Button(panel, ID_BTN_SSONUM, label="ССО №")
        self.BTNSSOnum.SetFont(self.fontBodyAct)

        self.TCSSONumDef = wx.TextCtrl(panel, ID_TC_SSONUM, wx.EmptyString, wx.DefaultPosition, size=(65, -1))
        self.TCSSONumDef.SetFont(self.fontBodyAct)

        self.BTNdatenum = wx.Button(panel, ID_BTN_DATEINP, label="Дата")
        self.BTNdatenum.SetFont(self.fontBodyAct)

        self.TCdateNumDef = wx.TextCtrl(panel, ID_TC_DATEINP, wx.EmptyString, wx.DefaultPosition, size=(150, -1))
        self.TCdateNumDef.SetFont(self.fontBodyAct)

        #Создание текстового поля ввода текста акта:
        self.TCTextInputCS = wx.TextCtrl(panel,
                                         wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_MULTILINE)
        self.TCTextInputCS.SetFont(self.fontBodyAct)
        # Добавление кнопки обновить, создать, очистить в сайзер:
        self.centrSCreateActsGeneral.Add(self.BTNrefreshAll)
        self.centrSCreateActsGeneral.Add(self.BTNCreateActCS, wx.EXPAND)
        self.centrSCreateActsGeneral.Add(self.BTNclearAll)

        # Добавление виджетов номера акта в сайзер определения акта:
        self.centrSActNumDef.Add(self.BTNactNum)
        self.centrSActNumDef.Add(self.TCActNumDef)

        # Добавление виджетов номера АЗС, ССО, и текущей даты в сайзер:
        self.centrSAZSSSONumDef.Add(self.BTNAZSnum)
        self.centrSAZSSSONumDef.Add(self.TCAZSNumDef)
        self.centrSAZSSSONumDef.Add(self.BTNSSOnum, flag=wx.LEFT, border=5)
        self.centrSAZSSSONumDef.Add(self.TCSSONumDef)
        self.centrSAZSSSONumDef.Add(self.BTNdatenum, flag=wx.LEFT, border=5)
        self.centrSAZSSSONumDef.Add(self.TCdateNumDef)

        # Добавление всех элементов в центральный сайзер:
        self.centrSCreateActs.Add(self.centrSCreateActsGeneral, 0, wx.TOP | wx.LEFT | wx.RIGHT | wx.EXPAND, 5)
        self.centrSCreateActs.Add(self.centrSActNumDef, proportion=0, flag=wx.TOP | wx.LEFT | wx.RIGHT |
                                                                           wx.ALIGN_CENTER, border=5)
        self.centrSCreateActs.Add(self.centrSAZSSSONumDef, proportion=0, flag=wx.TOP | wx.LEFT | wx.RIGHT, border=5)
        self.centrSCreateActs.Add(self.TCTextInputCS, 1, wx.ALL | wx.EXPAND, 5)
        #Добавление центрального сайзера в главный сайзер:
        self.mainSizer.Add(self.centrSCreateActs, proportion=1, flag=wx.EXPAND, border=5)


        #Правая часть программы, список созданных актов и кнопки для работы с ними:
        #Создание и добавление кнопок в сайзер:
        self.send_actBTN = wx.Button(panel, id=ID_BTN_SENDACT, label="Отправить Акт")
        self.del_actBTN = wx.Button(panel, id=ID_BTN_DELACT, label="Удалить файлы")
        self.refresh_lactsBTN = wx.Button(panel, id=ID_BTN_REFLACT, label="Обновить список")
        self.rightSSendActsTopBTN.Add(self.send_actBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)
        self.rightSSendActsTopBTN.Add(self.del_actBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)
        self.rightSSendActsTopBTN.Add(self.refresh_lactsBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)

        self.open_docx_BTN = wx.Button(panel, id=ID_BTN_OPENDOCX, label="Открыть docx")
        self.recopy_act_BTN = wx.Button(panel, id=ID_BTN_RECOPYACT, label="Коп. в папку Актов")
        self.copy_pdf_inbufferBTN = wx.Button(panel, id=ID_BTN_COPYPDFBUFF, label="Коп. pdf в буфер")
        self.rSizerfileworkBTN.Add(self.open_docx_BTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)
        self.rSizerfileworkBTN.Add(self.recopy_act_BTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)
        self.rSizerfileworkBTN.Add(self.copy_pdf_inbufferBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)

        self.rightSSendActs.Add(self.rightSSendActsTopBTN, proportion=0, flag=wx.EXPAND)
        self.rightSSendActs.Add(self.rSizerfileworkBTN, proportion=0, flag=wx.EXPAND)

        #Создание списка актов
        self.OLVlocal_acts = ObjectListView(panel, wx.ID_ANY, style=wx.LC_REPORT | wx.SUNKEN_BORDER)
        #Создание столбцов
        title = ColumnDefn("Имя", "left", 240, "title", isSpaceFilling=False)
        creating = ColumnDefn("Дата создания", "left", 130, "creating",  stringConverter="%d-%m-%Y %H:%M:%S",
                              isSpaceFilling=False)
        modifine = ColumnDefn("Дата изменения", "left", 130, "modifine",  stringConverter="%d-%m-%Y %H:%M:%S",
                              isSpaceFilling=False)
        self.OLVlocal_acts.oddRowsBackColor = wx.WHITE
        self.fontOLV = wx.Font(10, wx.MODERN, wx.NORMAL, wx.NORMAL, False, u'Roboto')
        self.OLVlocal_acts.SetFont(self.fontOLV)
        self.OLVlocal_acts.SetColumns([title, creating, modifine])
        #Добавление в список актов из папки locals_act
        self.OLVlocal_acts.SetObjects(
            fileProcessing.get_listdir_docx_files_in_dict(settings.get_local_acts_path_folder()))
        #Добавление списка актов в сайзер
        self.rightSSendActs.Add(self.OLVlocal_acts, proportion=1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)

        #Добавление правого сайзера в главный сайзер.
        self.mainSizer.Add(self.rightSSendActs, proportion=0, flag=wx.EXPAND)

        #Назначение главного сайзера
        panel.SetSizer(self.mainSizer)

        #Назначение событий
        self.Bind(wx.EVT_BUTTON, self.createActOn, id=ID_BTN_CRARCT)
        self.Bind(wx.EVT_BUTTON, self.sendActOn, id=ID_BTN_SENDACT)
        self.Bind(wx.EVT_BUTTON, self.del_acts_action, id=ID_BTN_DELACT)
        self.Bind(wx.EVT_BUTTON, self.refresh_list_acts, id=ID_BTN_REFLACT)
        self.Bind(wx.EVT_BUTTON, self.open_docx, id=ID_BTN_OPENDOCX)
        self.Bind(wx.EVT_BUTTON, self.recopy_file_in_general, id=ID_BTN_RECOPYACT)
        self.Bind(wx.EVT_BUTTON, self.copy_pdf_in_clipboard, id=ID_BTN_COPYPDFBUFF)
        self.Bind(wx.EVT_BUTTON, self.act_num_refresh, id=ID_BTN_ACTNUM)
        self.Bind(wx.EVT_BUTTON, self.azs_num_refresh, id=ID_BTN_AZSNUM)
        self.Bind(wx.EVT_BUTTON, self.sso_num_refresh, id=ID_BTN_SSONUM)
        self.Bind(wx.EVT_BUTTON, self.current_data_refresh, id=ID_BTN_DATEINP)
        self.Bind(wx.EVT_BUTTON, self.refresh_all, id=ID_BTN_REFALL)
        self.Bind(wx.EVT_BUTTON, self.clear_all, id=ID_BTN_CLEAR)

    def clear_all(self, event):
        if event.GetId() == ID_BTN_CLEAR:
            self.TCActNumDef.SetLabel("")
            self.TCAZSNumDef.SetLabel("")
            self.TCSSONumDef.SetLabel("")
            self.TCdateNumDef.SetLabel("")

    def refresh_all(self, event):
        if event.GetId() == ID_BTN_REFALL:
            self.act_num_refresh(event)
            self.azs_num_refresh(event)
            self.sso_num_refresh(event)
            self.current_data_refresh(event)

    def current_data_refresh(self, event):
        if event.GetId() == ID_BTN_DATEINP or event.GetId() == ID_BTN_REFALL:
            self.TCdateNumDef.SetLabel(fileProcessing.get_current_date())

    def sso_num_refresh(self, event):
        if event.GetId() == ID_BTN_SSONUM or event.GetId() == ID_BTN_REFALL:
            txtlst = list(map(lambda x: x.strip(), self.TCTextInputCS.GetValue().split('\n')))
            if self.TCTextInputCS.GetValue():
                if fileProcessing.get_from_bodylist_ssonum(txtlst):
                    self.TCSSONumDef.SetLabel(fileProcessing.get_from_bodylist_ssonum(txtlst))

    def azs_num_refresh(self, event):
        if event.GetId() == ID_BTN_AZSNUM or event.GetId() == ID_BTN_REFALL:
            txtlst = list(map(lambda x: x.strip(), self.TCTextInputCS.GetValue().split('\n')))
            if self.TCTextInputCS.GetValue():
                if fileProcessing.get_from_bodylist_azsnum(txtlst):
                    self.TCAZSNumDef.SetLabel(fileProcessing.get_from_bodylist_azsnum(txtlst))

    def act_num_refresh(self, event):
        if event.GetId() == ID_BTN_ACTNUM or event.GetId() == ID_BTN_REFALL:
            self.TCActNumDef.SetLabel(fileProcessing.get_number_act())

    def copy_pdf_in_clipboard(self, event):
        if event.GetId() == ID_BTN_COPYPDFBUFF:
            selection = self.OLVlocal_acts.GetSelectedObjects()
            if selection:
                for i in range(len(selection)):
                    path = fileProcessing.get_path_to_file_to_string(
                        fileProcessing.get_name_pdf_from_docx(selection[i]['title']))
                    print(path)
                    if i == 0:
                        proc = subprocess.Popen(['powershell', f'Set-Clipboard -Path {path}'])
                        proc.wait()
                    else:
                        proc = subprocess.Popen(['powershell', f'Set-Clipboard -Append -Path {path}'])
                        proc.wait()

    def recopy_file_in_general(self, event):
        if event.GetId() == ID_BTN_RECOPYACT:
            selection = self.OLVlocal_acts.GetSelectedObjects()
            if selection:
                for i in range(len(selection)):
                    fileProcessing.create_pdf_file_from_docx(fileProcessing.get_path_to_file_to_string(
                        selection[i]['title']))
                    fileProcessing.copy_files_to_general_folder(selection[i]['title'])

    def open_docx(self, event):
        if event.GetId() == ID_BTN_OPENDOCX:
            selection = self.OLVlocal_acts.GetSelectedObjects()
            if selection:
                os.startfile(os.path.realpath(settings.get_local_acts_path_folder()) + '\\' + selection[0]['title'])

    def onSettings(self, event):
        with mydlg.MyDlg(self, title="Настройки") as dlg:
            dlg.ShowModal()

    def onQuit(self, event):
        self.Close()

    def openFolderLocalActs(self, event):
        os.startfile(os.path.realpath(settings.get_local_acts_path_folder()))

    def openFolderActs(self, event):
        os.startfile(os.path.realpath(settings.get_general_acts_path_folder()))

    def openDocxTemplate(self, event):
        os.startfile(os.path.realpath(settings.get_docx_templ_path()))

    def about(self, event):
        with aboutdlg.AboutDlg(self, title="Настройки") as dlg:
            dlg.ShowModal()

    def del_acts_action(self, event):
        if event.GetId() == ID_BTN_DELACT:
            selection = self.OLVlocal_acts.GetSelectedObjects()
            if selection:
                dlg = wx.MessageBox('Удалить выбранные файлы?', 'Подтверждение', wx.YES_NO | wx.NO_DEFAULT, self)
                if dlg == wx.YES:
                    for i in range(len(selection)):
                        os.remove(fileProcessing.get_path_to_file_to_string(selection[i]['title']))
                        os.remove(fileProcessing.get_path_to_file_to_string(
                            fileProcessing.get_name_pdf_from_docx(selection[i]['title'])
                        ))
                        self.refresh_list_acts(event)

    def refresh_list_acts(self, event):
        self.OLVlocal_acts.SetObjects(fileProcessing.get_listdir_docx_files_in_dict(settings.get_local_acts_path_folder()))

    def createActOn(self, event):
        if event.GetId() == ID_BTN_CRARCT and self.TCTextInputCS.GetNumberOfLines() > 0:
            txtlst = list(map(lambda x: x.strip(), self.TCTextInputCS.GetValue().split('\n')))
            progress = wx.ProgressDialog("Создание Акта...", "Этап 0 из 6: Инициализация метода создания акта",
                                         maximum=100,
                                         parent=self,
                                         style=wx.PD_AUTO_HIDE | wx.PD_APP_MODAL)
            progress.Update(10, "Этап 1 из 6: Формирование имя акта")
            filenamedocx = \
                f"Акт{self.TCActNumDef.GetValue()}_АЗС{self.TCAZSNumDef.GetValue()}_ССО{self.TCSSONumDef.GetValue()}." \
                f"docx"
            progress.Update(20, "Этап 2 из 6: Формирование пути расположения акта")
            docxpath = settings.get_local_acts_path_folder() + filenamedocx
            progress.Update(30, "Этап 3 из 6: Создание docx документа")
            fileProcessing.create_docx_file_from_bodylist(
                txtlst,
                self.TCActNumDef.GetValue(),
                self.TCAZSNumDef.GetValue(),
                self.TCSSONumDef.GetValue(),
                self.TCdateNumDef.GetValue(),
                docxpath
            )
            progress.Update(60, "Этап 4 из 6: Создание pdf документа")
            fileProcessing.create_pdf_file_from_docx(docxpath)
            progress.Update(80, "Этап 5 из 6: Копирование файлов в общую папку")
            fileProcessing.copy_files_to_general_folder(filenamedocx)
            self.refresh_list_acts(event)
            progress.Update(100, "Этап 6 из 6: Обновление списка актов")
            progress.Destroy()

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
                    mess.Attachments.Add(
                        fileProcessing.get_path_to_file_to_string(
                            fileProcessing.get_name_pdf_from_docx(selection[i]['title'])
                        )
                    )
                print('Отправка письма')
                mess.Display(True)
                print('Письмо отправлено')
