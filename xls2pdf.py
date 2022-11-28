import os
import stat
import time
import math
import wx
import random
from ObjectListView import ObjectListView, ColumnDefn
from pdf_xlsx import process_xls_pdf
import wx.lib.scrolledpanel as scrolled


class MyForm(wx.Frame):
    # demo de fontes
    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, "Font Tutorial")

        # Add a panel so it looks the correct on all platforms
        panel = scrolled.ScrolledPanel(self)
        panel.SetAutoLayout(1)
        panel.SetupScrolling()

        fontSizer = wx.BoxSizer(wx.VERTICAL)
        families = {"FONTFAMILY_DECORATIVE":wx.FONTFAMILY_DECORATIVE, # A decorative font
                    "FONTFAMILY_DEFAULT":wx.FONTFAMILY_DEFAULT,
                    "FONTFAMILY_MODERN":wx.FONTFAMILY_MODERN,     # Usually a fixed pitch font
                    "FONTFAMILY_ROMAN":wx.FONTFAMILY_ROMAN,      # A formal, serif font
                    "FONTFAMILY_SCRIPT":wx.FONTFAMILY_SCRIPT,     # A handwriting font
                    "FONTFAMILY_SWISS":wx.FONTFAMILY_SWISS,      # A sans-serif font
                    "FONTFAMILY_TELETYPE":wx.FONTFAMILY_TELETYPE    # A teletype font
                    }
        weights = {"FONTWEIGHT_BOLD":wx.FONTWEIGHT_BOLD,
                   "FONTWEIGHT_LIGHT":wx.FONTWEIGHT_LIGHT,
                   "FONTWEIGHT_NORMAL":wx.FONTWEIGHT_NORMAL
                   }

        styles = {"FONTSTYLE_ITALIC":wx.FONTSTYLE_ITALIC,
                  "FONTSTYLE_NORMAL":wx.FONTSTYLE_NORMAL,
                  "FONTSTYLE_SLANT":wx.FONTSTYLE_SLANT
                  }
        sizes = [8, 10, 12, 14]
        for family in families.keys():
            for weight in weights.keys():
                for style in styles.keys():
                    label = "%s    %s    %s" % (family, weight, style)
                    size = random.choice(sizes)
                    font = wx.Font(size, families[family], styles[style],
                                   weights[weight])
                    txt = wx.StaticText(panel, label=label)
                    txt.SetFont(font)
                    fontSizer.Add(txt, 0, wx.ALL, 5)
        panel.SetSizer(fontSizer)
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(panel, 1, wx.EXPAND)
        self.SetSizer(sizer)


class TimedDialog(wx.Dialog):
    def __init__(self, title, message, *args, **kwargs):
        super(TimedDialog, self).__init__(None, *args,  style=wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP,
                                          pos=(0,100), **kwargs)

        self.SetSize((400, 100))
        self.SetTitle(title)
        self.Centre()

        font = wx.Font(14, wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        text = wx.StaticText(self, -1, size=(448, -1), label=message, style=(wx.ALIGN_CENTRE_HORIZONTAL | wx.TE_MULTILINE ))
        text.SetFont(font)
        text.Wrap(448)
        box = wx.BoxSizer(wx.HORIZONTAL)
        box.Add(text, 1,  wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, 15)

        self.SetSizer(box)
        self.Layout()
        self.Refresh()

        # Center form
        self.Center()
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.OnTimer)

        self.timer.Start(3000)  # 2 second interval

    def OnTimer(self, event):
        self.Close()


class MyFileDropTarget(wx.FileDropTarget):
    """"""
    def __init__(self, window):
        """Constructor"""
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        """
        When files are dropped, update the display
        """
        self.window.updateDisplay(filenames)
        return True


class FileInfo(object):
    """"""

    def __init__(self, path, date_created, date_modified, size):
        """Constructor"""
        self.name = os.path.basename(path)

        self.path = path
        self.date_created = date_created
        self.date_modified = date_modified
        self.size = size


class MainPanel(wx.Panel):
    """"""
    def __init__(self, parent):
        """Constructor"""
        wx.Panel.__init__(self, parent=parent)

        self.file_list = []

        file_drop_target = MyFileDropTarget(self)
        self.olv = ObjectListView(self, style=wx.LC_REPORT|wx.SUNKEN_BORDER)
        self.olv.SetDropTarget(file_drop_target)
        self.setFiles()

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.olv, 1, wx.EXPAND)
        self.SetSizer(sizer)

        self.addbutton = wx.Button(self, wx.ID_ANY, "PROCESSAR")
        self.Bind(wx.EVT_BUTTON, self.on_toggle_plotlist, self.addbutton)
        #sizer.Add(self.addbutton, 1, wx.EXPAND)
        sizer.Add(self.addbutton, 0, wx.ALL|wx.CENTER, 5)

    def on_toggle_plotlist(self, event):

        dlg = TimedDialog("Aviso", "teste de mensagem")
        dlg.ShowModal()

        count = self.olv.GetItemCount()
        xls_file = ""
        pdf_file = []

        if count == 0:
            return

        for row in range(count):
            item = self.olv.GetItem(row, col=0)
            print(item.GetText())

            if ".xlsx" in item.GetText():
                xls_file = item.GetText()
            if ".pdf" in item.GetText():
                pdf_file.append(item.GetText())

        if len(pdf_file) > 1:
            # xlsx ja esta em pdf, basta unir e comprimir
            process_xls_pdf(xls_file[0], pdf_file[1])
        else:
            process_xls_pdf(xls_file, pdf_file[0])

        dlg = TimedDialog(self, message="teste")
        dlg.ShowModal()

    def updateDisplay(self, file_list):
        """"""
        for path in file_list:
            file_stats = os.stat(path)
            creation_time = time.strftime("%m/%d/%Y %I:%M %p",
                                          time.localtime(file_stats[stat.ST_CTIME]))
            modified_time = time.strftime("%m/%d/%Y %I:%M %p",
                                          time.localtime(file_stats[stat.ST_MTIME]))
            file_size = file_stats[stat.ST_SIZE]
            if file_size > 1024:
                file_size = file_size / 1024.0
                file_size = "%.2f KB" % file_size

            self.file_list.append(FileInfo(path,
                                           creation_time,
                                           modified_time,
                                           file_size))

        self.olv.SetObjects(self.file_list)

    def setFiles(self):
        """"""
        self.olv.SetColumns([
            ColumnDefn("Name", "left", 220, "path"),
            ColumnDefn("Date created", "left", 150, "date_created"),
            ColumnDefn("Date modified", "left", 150, "date_modified"),
            ColumnDefn("Size", "left", 100, "size")
        ])
        self.olv.SetObjects(self.file_list)


class MainFrame(wx.Frame):
    """"""
    def __init__(self):
        """Constructor"""
        wx.Frame.__init__(self, None, title="XLS PDF Join", size=(650, 400))
        panel = MainPanel(self)
        self.Center()
        self.Show()


def main():
    """"""
    app = wx.App(False)
    frame = MainFrame()
    app.MainLoop()


if __name__ == "__main__":
    main()
