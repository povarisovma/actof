import wx
import wx.lib.mixins.listctrl
import settings
import myframe


def main():
    settings.createConfig()
    app = wx.App()
    frame = myframe.MyFrame(None).Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
