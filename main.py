import wx
import settings
import myframe
import templatesdb


def main():
    templatesdb.createdb()
    settings.create_settings_file()
    app = wx.App()
    frame = myframe.MyFrame(None).Show()
    app.MainLoop()


if __name__ == '__main__':
    main()


#TODO добавление обработки исключений, отсутствие номера акта, ССО, АЗС, отсутствие пути,
# и тп(разбить на несколько пунктов)
#TODO добавление инструкции для пользорвателя и раздела "о программе"

#TODO добавление возможности создания списка шаблонов, левой части программы(разбить на несколько блоков)