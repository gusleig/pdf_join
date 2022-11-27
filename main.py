import os
import stat
import time
import wx
from ObjectListView import ObjectListView, ColumnDefn
from pdf_xlsx import process_xls_pdf
########################################################################
class MyFileDropTarget(wx.FileDropTarget):
    """"""
    #----------------------------------------------------------------------
    def __init__(self, window):
        """Constructor"""
        wx.FileDropTarget.__init__(self)
        self.window = window

    #----------------------------------------------------------------------
    def OnDropFiles(self, x, y, filenames):
        """
        When files are dropped, update the display
        """
        self.window.updateDisplay(filenames)
        return True
########################################################################
class FileInfo(object):
    """"""
    #----------------------------------------------------------------------
    def __init__(self, path, date_created, date_modified, size):
        """Constructor"""
        self.name = os.path.basename(path)

        self.path = path
        self.date_created = date_created
        self.date_modified = date_modified
        self.size = size


class MainPanel(wx.Panel):
    """"""
    #----------------------------------------------------------------------
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
        print("Click!")
        count = self.olv.GetItemCount()
        xls_file = ""
        pdf_file = ""

        for row in range(count):
            item = self.olv.GetItem(row, col=0)
            print(item.GetText())

            if ".xlsx" in item.GetText():
                xls_file = item.GetText()
            if ".pdf" in item.GetText():
                pdf_file = item.GetText()

        process_xls_pdf(xls_file, pdf_file)

    def onclick(self, event):
        print("yay it works")

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

    #----------------------------------------------------------------------
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
        self.Show()


def main():
    """"""
    app = wx.App(False)
    frame = MainFrame()
    app.MainLoop()


if __name__ == "__main__":
    main()