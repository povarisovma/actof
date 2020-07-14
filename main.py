import wx
import wx.lib.mixins.listctrl
import settings
import myframe


def main():
    settings.create_settings_file()
    app = wx.App()
    frame = myframe.MyFrame(None).Show()
    app.MainLoop()


if __name__ == '__main__':
    main()


#TODO добавление обработки исключений, отсутствие номера акта, ССО, АЗС, отсутствие пути,
# и тп(разбить на несколько пунктов)
#TODO добавление кнопки открытия файла pdf и docx из локальных актов
#TODO добавить кнопку копирования актов из папки локальных актов в папку общих актов
#TODO добавление возможности загрузки шаблона docx
#TODO добавление кнопки очистить текстовое поле
#TODO добавление текста по умолчанию в текстовом поле
#TODO добавление инструкции для пользорвателя и раздела "о программе"

#TODO добавление возможности создания списка шаблонов, левой части программы(разбить на несколько блоков)