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
ID_LC_ACTLIST = 35

class ListCtrlMixinx(wx.ListCtrl, wx.lib.mixins.listctrl.ColumnSorterMixin):
    def __init__(self, parent, *args, **kw):
        wx.ListCtrl.__init__(self, parent, wx.ID_ANY, style=wx.LC_REPORT)
        wx.lib.mixins.listctrl.ColumnSorterMixin.__init__(self, 3)
        self.itemDataMap = getlistfiles.getDictfiles()

    def GetListCtrl(self):
        return self

class MyFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title='ActOf', size=wx.Size(1037, 605),
                          style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        #Объявление сайзеров:
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
        self.rightSSendActsTopBTN.Add(self.send_actBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)
        self.rightSSendActsTopBTN.Add(self.del_actBTN, 1, flag=wx.EXPAND | wx.TOP | wx.RIGHT, border=5)

        self.rightSSendActs.Add(self.rightSSendActsTopBTN, proportion=0, flag=wx.EXPAND)

        #Создание списка актов
        self.OLVlocal_acts = ObjectListView(self, wx.ID_ANY, style=wx.LC_REPORT | wx.SUNKEN_BORDER)
        #Создание столбцов
        title = ColumnDefn("Title", "left", 220, "title", isSpaceFilling=False)
        creating = ColumnDefn("Date Creating", "left", 150, "creating",  stringConverter="%d-%m-%Y %H:%M:%S", isSpaceFilling=False)
        modifine = ColumnDefn("Date Modifine", "left", 150, "modifine",  stringConverter="%d-%m-%Y %H:%M:%S", isSpaceFilling=False)
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
                        self.refresh_list_acts()


    def refresh_list_acts(self):
        self.OLVlocal_acts.SetObjects(getlistfiles.getDictFilesParam())

    def createActOn(self, event):
        if event.GetId() == ID_BTN_CRARCT and self.TCTextInputCS.GetNumberOfLines() > 0:
            txtlst = list(map(lambda x: x.strip(), self.TCTextInputCS.GetValue().split('\n')))
            docxfilemaker.createdocxnpdffiles(txtlst)
            self.refresh_list_acts()

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
