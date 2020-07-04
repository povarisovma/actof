import wx
import os
ID_BTN_APL = 15
ID_BTN_ADD = 16
ID_BTN_RMV = 17
ID_LB_TMPL = 45
ID_TC_WTEXT = 26

tmpllist = []
filelist = os.listdir('.\\template')
print(filelist)
for file in filelist:
    with open('.\\template\\' + file, 'r', encoding='utf-8') as f:
        tmpllist.append(tuple(f.read().split('\n')))
        # for line in f:
        #     print(line.strip())


# tmpllist.append(('Fuel', 'Не учтенный пролив', 'Настоящим подтверждаю, что 8.05.2020 на АЗС №31052 в результате сбоя,'
#                                                ' пролив топлива АИ-92 объемом 30,12 л. на сумму 1249,98 руб. был учтен'
#                                                ' в ССО №58 как пролив объемом 12,21 л. на сумму 506,72 руб. из-за чего'
#                                                'возникло расхождение по счетчикам ТРК. Оплата производилась по'
#                                                ' банковской карте №****4240 по RRN № 012923616562.'))
# tmpllist.append(('Fuel', 'Задвоенный пролив', 'Настоящим подтверждаю, что 8.05.2020 на АЗС №31052 в результате сбоя,'
#                                                ' пролив топлива АИ-9 объемом 30,12 л. на сумму 1249,98 руб. был учтен'
#                                                ' в ССО №58 как пролив объемом 12,21 л. на сумму 506,72 руб. из-за чего'
#                                                'возникло расхождение по счетчикам ТРК. Оплата производилась по'
#                                                ' банковской карте №****4240 по RRN № 012923616562.'))
# tmpllist.append(('DDS', 'Темпокасса по АСУ', 'Настоящим подтверждаю, что 8.05.2020 на АЗС №31052 в результате сбоя,'
#                                                ' пролив топлива АИ-9 объемом 30,12 л. на сумму 1249,98 руб. был учтен'
#                                                ' в ССО №58 как пролив объемом 12,21 л. на сумму 506,72 руб. из-за чего'
#                                                'возникло расхождение по счетчикам ТРК. Оплата производилась по'
#                                                ' банковской карте №****4240 по RRN № 012923616562.'))
# tmpllist.append(('DDS', 'Темпокасса по ККМ', 'Настоящим подтверждаю, что 8.05.2020 на АЗС №31052 в результате сбоя,'
#                                                ' пролив топлива АИ-9 объемом 30,12 л. на сумму 1249,98 руб. был учтен'
#                                                ' в ССО №58 как пролив объемом 12,21 л. на сумму 506,72 руб. из-за чего'
#                                                'возникло расхождение по счетчикам ТРК. Оплата производилась по'
#                                                ' банковской карте №****4240 по RRN № 012923616562.'))
# tmpllist.append(('DDS', 'Не верная выплата', 'Настоящим подтверждаю, что 8.05.2020 на АЗС №31052 в результате сбоя,'
#                                                ' пролив топлива АИ-9 объемом 30,12 л. на сумму 1249,98 руб. был учтен'
#                                                ' в ССО №58 как пролив объемом 12,21 л. на сумму 506,72 руб. из-за чего'
#                                                'возникло расхождение по счетчикам ТРК. Оплата производилась по'
#                                                ' банковской карте №****4240 по RRN № 012923616562.\n\tТакже подтверждаю'
#                                              ', '
#                                              'что 8.05.2020 на АЗС №30052 в результате сбоя, пролив топлива АИ-92 '
#                                              'объемом 12,21 л. на сумму 506 руб. не был учтен в ССО №58.'
#                                              ' Оплата производилась за наличный расчет.'))
# tmpllist.append(('XZ', 'test', '''Настоящим подтверждаю, что 8.05.2020 на АЗС №31052 в результате сбоя, пролив
#                                     топлива АИ-92 объемом 30,12 л. на сумму 1249,98 руб. был учтен в ССО №58 как
#                                     пролив объемом 12,21 л. на сумму 506,72 руб. из-за чего возникло расхождение по
#                                     счетчикам ТРК. Оплата производилась по банковской карте №****4240 по RRN №
#                                     012923616562.
#     Также подтверждаю, что 8.05.2020 на АЗС №30052 в результате сбоя, пролив топлива АИ-92 объемом 12,21 л. на сумму
#     506 руб. не был учтен в ССО №58. Оплата производилась за наличный расчет.
#     В связи с чем, прошу:
#     Считать пролив топлива АИ-92 по банковской карте №****4240 от 8.05.2020 объемом 30,12 л. на сумму 1249,98 руб.
#     Пролив топлива АИ-92 объемом 12,21 л. на сумму 506 руб. считать проливом за наличный расчет.
#     Выручку по банковским картам считать равной 135738,51 руб.
#     Выручку за наличный расчет считать равной 231838 руб.
#     Произвести внесение на кассе №1 по АСУ+ККМ на сумму 506 руб.'''))


class MyFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title='ActOf', pos=wx.DefaultPosition,
                          size=wx.Size(1238, 754),
                          style=wx.CAPTION | wx.DEFAULT_FRAME_STYLE | wx.MAXIMIZE | wx.TAB_TRAVERSAL)
        self.panel = wx.Panel(self)
        font = wx.SystemSettings.GetFont(wx.SYS_DEFAULT_GUI_FONT)
        font.SetPointSize(8)
        self.panel.SetFont(font)

        wx.Button(self.panel, id=ID_BTN_APL, label='Apply', pos=(10, 10), size=(70, 30))
        wx.Button(self.panel, id=ID_BTN_ADD, label='Add', pos=(90, 10), size=(70, 30))
        wx.Button(self.panel, id=ID_BTN_RMV, label='Remove', pos=(170, 10), size=(70, 30))
        wx.StaticText(self.panel, label='Template list:', pos=(90, 50), size=(70, 30))
        self.listtempl = wx.ListCtrl(self.panel, id=ID_LB_TMPL, pos=(10, 80), size=(230, 500), style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.listtempl.InsertColumn(0, 'Type', width=50)
        self.listtempl.InsertColumn(1, 'Shorttext', width=150)
        for i in range(len(tmpllist)):
            self.listtempl.Append((tmpllist[i][0], tmpllist[i][1]))


        self.tc_worktext = wx.TextCtrl(self.panel, id=ID_TC_WTEXT, pos=(300, 80), size=(700, 900), style=wx.TE_MULTILINE)



        self.Bind(wx.EVT_BUTTON, self.btnOkActiv)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.btnOkActiv, id=ID_LB_TMPL)


    def btnOkActiv(self, event):
        print('press ', event.GetId())
        if event.GetId() == ID_BTN_APL or event.GetId() == ID_LB_TMPL:
            print(tmpllist[self.listtempl.GetFirstSelected()])
            self.tc_worktext.AppendText('\t' + tmpllist[self.listtempl.GetFirstSelected()][2] + '\n')





app = wx.App()
frame = MyFrame(None).Show()

app.MainLoop()
