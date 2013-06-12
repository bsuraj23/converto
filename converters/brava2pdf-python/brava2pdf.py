# -*- coding: utf-8 -*-

import wx
import win32com.client 
import wx.activex 
import wx.py
import wx.lib.activex
import comtypes.client
import comtypes

comtypes.client.GetModule(('{A7373E99-5D7C-47A3-BA29-8AE0B3E59949}', 1, 0))
import comtypes.gen.BRAVADTXLib as BRAVADTXLib

bravaCtrlProgID = 'BRAVADTX.BravaDTXView.1'

class BravaActiveXCtrl(wx.lib.activex.ActiveXCtrl):
    def __init__(self, parent, id = -1, pos = wx.DefaultPosition,
                 size = wx.DefaultSize, style = 0, name = '', app = None, input_file = None):
            wx.lib.activex.ActiveXCtrl.__init__(self, parent, bravaCtrlProgID,
                                                id, pos, size, style, name)
            self.app = app
            self.ctrl.Filename = input_file
    def FileLoaded(self, this, sender, event):
        print "FileLoaded"
        self.ctrl.ExportPDF(r'C:\out.pdf', 1)
    def FileLoadFailure(self, sender, event):
        print "FileLoadFailure"
    def ExportPDFSuccess(self, this, sender, event):
        print "OnExportPDFSuccess"
        self.ctrl.CloseFile()
        self.app.ExitMainLoop()
    def ExportPDFFailure(self, sender, event):
        print "OnExportPDFFailure"
        
class MyApp(wx.App): 
    def __init__(self, redirect = False, filename = None):
        wx.App.__init__(self, redirect, filename)
        self.frame = wx.Frame(None, wx.ID_ANY, title = 'brava2pdf')
        self.panel = wx.Panel(self.frame, wx.ID_ANY)

if __name__== '__main__':
    app = MyApp()
    box = wx.BoxSizer(wx.VERTICAL)

    bravaCtrl = BravaActiveXCtrl(app.panel, style = wx.SUNKEN_BORDER,
                                 app = app,
                                 input_file = r'C:\lineweights.dwg')
    box.Add(bravaCtrl, proportion = 1, flag = wx.EXPAND)

    app.panel.SetSizer(box)
    # app.frame.Show()

    # app.MainLoop()

    print "OK"