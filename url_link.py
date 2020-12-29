# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################
import webbrowser
import wx
import wx.xrc
import sqlite3
import os
import json
import winreg
import win32com
import win32con
import win32com.client



###########################################################################
## Class MyFrame1
###########################################################################

class MyFrame1(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=r"链接收藏工具", pos=wx.DefaultPosition,
                          size=wx.Size(596, 444), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
        self.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_MENU))
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_MENU))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        self.m_notebook1 = wx.Notebook(self, wx.ID_ANY, wx.DefaultPosition, wx.Size(500, 30), 0)
        self.m_panel2 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.Size(855, 20), wx.TAB_TRAVERSAL)
        self.m_panel2.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_MENU))

        gSizer1 = wx.GridSizer(0, 2, 0, 0)

        self.m_staticText32 = wx.StaticText(self.m_panel2, wx.ID_ANY, u"链接地址", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText32.Wrap(-1)
        gSizer1.Add(self.m_staticText32, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_textCtrl13 = wx.TextCtrl(self.m_panel2, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer1.Add(self.m_textCtrl13, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.m_button33 = wx.Button(self.m_panel2, wx.ID_ANY, u"打开链接", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button33.SetDefault()
        gSizer1.Add(self.m_button33, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_button34 = wx.Button(self.m_panel2, wx.ID_ANY, u"清除界面信息", wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer1.Add(self.m_button34, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_panel2.SetSizer(gSizer1)
        self.m_panel2.Layout()
        self.m_notebook1.AddPage(self.m_panel2, u"    打开操作    ", False)
        self.m_panel21 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        gSizer2 = wx.GridSizer(0, 2, 0, 0)

        self.m_staticText33 = wx.StaticText(self.m_panel21, wx.ID_ANY, u"链接名称", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText33.Wrap(-1)
        gSizer2.Add(self.m_staticText33, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_textCtrl14 = wx.TextCtrl(self.m_panel21, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                        0)
        gSizer2.Add(self.m_textCtrl14, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.m_staticText34 = wx.StaticText(self.m_panel21, wx.ID_ANY, u"链接地址", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText34.Wrap(-1)
        gSizer2.Add(self.m_staticText34, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_textCtrl15 = wx.TextCtrl(self.m_panel21, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(270, 30),
                                        wx.TE_AUTO_URL | wx.TE_MULTILINE)
        gSizer2.Add(self.m_textCtrl15, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.m_button35 = wx.Button(self.m_panel21, wx.ID_ANY, u"新增数据", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button35.SetDefault()
        gSizer2.Add(self.m_button35, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_button36 = wx.Button(self.m_panel21, wx.ID_ANY, u"清除界面信息", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button36.SetDefault()
        gSizer2.Add(self.m_button36, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_panel21.SetSizer(gSizer2)
        self.m_panel21.Layout()
        gSizer2.Fit(self.m_panel21)
        self.m_notebook1.AddPage(self.m_panel21, u"    新增操作    ", False)
        self.m_panel3 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        gSizer3 = wx.GridSizer(0, 2, 0, 0)

        self.m_staticText35 = wx.StaticText(self.m_panel3, wx.ID_ANY, u"链接名称", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText35.Wrap(-1)
        gSizer3.Add(self.m_staticText35, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_textCtrl16 = wx.TextCtrl(self.m_panel3, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer3.Add(self.m_textCtrl16, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.m_button37 = wx.Button(self.m_panel3, wx.ID_ANY, u"删除链接", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button37.SetDefault()
        gSizer3.Add(self.m_button37, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_button38 = wx.Button(self.m_panel3, wx.ID_ANY, u"清除界面信息", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button38.SetDefault()
        gSizer3.Add(self.m_button38, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_panel3.SetSizer(gSizer3)
        self.m_panel3.Layout()
        gSizer3.Fit(self.m_panel3)
        self.m_notebook1.AddPage(self.m_panel3, u"    删除操作    ", False)
        self.m_panel5 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.Size(150, -1), wx.TAB_TRAVERSAL)
        gSizer4 = wx.GridSizer(0, 2, 0, 0)

        self.m_staticText36 = wx.StaticText(self.m_panel5, wx.ID_ANY, u"链接名称", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText36.Wrap(-1)
        gSizer4.Add(self.m_staticText36, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_textCtrl17 = wx.TextCtrl(self.m_panel5, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(250, -1),
                                        0)
        gSizer4.Add(self.m_textCtrl17, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.m_staticText37 = wx.StaticText(self.m_panel5, wx.ID_ANY, u"查询结果", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText37.Wrap(-1)
        gSizer4.Add(self.m_staticText37, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_textCtrl18 = wx.TextCtrl(self.m_panel5, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(270, 30),
                                        wx.TE_MULTILINE)
        gSizer4.Add(self.m_textCtrl18, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.m_button40 = wx.Button(self.m_panel5, wx.ID_ANY, u"查询链接", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button40.SetDefault()
        gSizer4.Add(self.m_button40, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_button41 = wx.Button(self.m_panel5, wx.ID_ANY, u"打开链接", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button41.SetDefault()
        gSizer4.Add(self.m_button41, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_button42 = wx.Button(self.m_panel5, wx.ID_ANY, u"清除界面信息", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button42.SetDefault()
        gSizer4.Add(self.m_button42, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_button12 = wx.Button(self.m_panel5, wx.ID_ANY, u"创建桌面快捷方式", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button12.SetDefault()
        gSizer4.Add(self.m_button12, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_panel5.SetSizer(gSizer4)
        self.m_panel5.Layout()
        self.m_notebook1.AddPage(self.m_panel5, u"精准查询操作", False)
        self.m_panel51 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        sbSizer8 = wx.StaticBoxSizer(wx.StaticBox(self.m_panel51, wx.ID_ANY, wx.EmptyString), wx.VERTICAL)

        self.m_button17 = wx.Button(sbSizer8.GetStaticBox(), wx.ID_ANY, u"查询所有数据", wx.DefaultPosition, wx.DefaultSize,
                                    0)
        self.m_button17.SetDefault()
        sbSizer8.Add(self.m_button17, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_button18 = wx.Button(sbSizer8.GetStaticBox(), wx.ID_ANY, u"清除界面信息", wx.DefaultPosition, wx.DefaultSize,
                                    0)
        sbSizer8.Add(self.m_button18, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        sbSizer10 = wx.StaticBoxSizer(wx.StaticBox(sbSizer8.GetStaticBox(), wx.ID_ANY, wx.EmptyString), wx.VERTICAL)

        self.m_textCtrl131 = wx.TextCtrl(sbSizer10.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                         wx.DefaultSize, wx.TE_AUTO_URL | wx.TE_MULTILINE)
        sbSizer10.Add(self.m_textCtrl131, 1, wx.ALL | wx.EXPAND, 5)

        sbSizer8.Add(sbSizer10, 1, wx.EXPAND, 5)

        bSizer2.Add(sbSizer8, 1, wx.EXPAND, 5)

        self.m_panel51.SetSizer(bSizer2)
        self.m_panel51.Layout()
        bSizer2.Fit(self.m_panel51)
        self.m_notebook1.AddPage(self.m_panel51, u"全局查询操作", True)

        bSizer1.Add(self.m_notebook1, 1, wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.m_button33.Bind(wx.EVT_BUTTON, self.a)
        self.m_button34.Bind(wx.EVT_BUTTON, self.b)
        self.m_button35.Bind(wx.EVT_BUTTON, self.c)
        self.m_button36.Bind(wx.EVT_BUTTON, self.d)
        self.m_button37.Bind(wx.EVT_BUTTON, self.e)
        self.m_button38.Bind(wx.EVT_BUTTON, self.f)
        self.m_button40.Bind(wx.EVT_BUTTON, self.g)
        self.m_button41.Bind(wx.EVT_BUTTON, self.h)
        self.m_button42.Bind(wx.EVT_BUTTON, self.i)
        self.m_button12.Bind(wx.EVT_BUTTON, self.uuu_url)
        self.m_button17.Bind(wx.EVT_BUTTON, self.j)
        self.m_button18.Bind(wx.EVT_BUTTON, self.k)


    def uuu_url(self, event):
        try:
            bmurl = self.m_textCtrl18.GetValue()
            lin_url=self.m_textCtrl17.GetValue()
            if bmurl =="" or lin_url=="":
                wx.MessageBox("名称和地址都不能为空", "通知", wx.OK | wx.ICON_INFORMATION)
            else:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
                k = winreg.QueryValueEx(key, "Desktop")[0] + "{}".format("\\")
                bmpath = k + "{}".format(str(lin_url)) + ".url"
                #print(bmpath)
                ws = win32com.client.Dispatch("wscript.shell")
                scut = ws.CreateShortcut(bmpath)
                scut.TargetPath = bmurl
                scut.Save()
        except:
            wx.MessageBox("快捷方式名称取名不能带有特殊字符，请重新操作", "通知", wx.OK | wx.ICON_INFORMATION)


    # V打开链接
    def a(self, event):
        x=self.m_textCtrl13.GetValue()
        try:
            if x =="" :
                wx.MessageBox("输入地址不能为空", "通知", wx.OK | wx.ICON_INFORMATION)
            else:
                webbrowser.open(x)
        except:
            wx.MessageBox("打开异常，重新打开", "通知", wx.OK | wx.ICON_INFORMATION)

    #清除链接
    def b(self, event):
        self.m_textCtrl13.Clear()


    def c(self, event):
        Name_name=self.m_textCtrl14.GetValue()
        Name_url=self.m_textCtrl15.GetValue()
        pwd = os.getcwd()+"{}".format("\\")+"{}".format("test.db")
        conn = sqlite3.connect(pwd)
        c = conn.cursor()
        try:
            if Name_name =="" or Name_url=="":
                wx.MessageBox("链接名称或地址不能为空", "通知", wx.OK | wx.ICON_INFORMATION)
            else:
                c.execute("INSERT INTO test_url(name, link) VALUES ('%s', '%s')" % (Name_name, Name_url))
                wx.MessageBox("新增成功", "通知", wx.OK | wx.ICON_INFORMATION)
        except:
            wx.MessageBox("操作异常，重新操作", "通知", wx.OK | wx.ICON_INFORMATION)
        conn.commit()
        conn.close()

    def d(self, event):
        self.m_textCtrl14.Clear()
        self.m_textCtrl15.Clear()

    def e(self, event):
        pwd = os.getcwd()+"{}".format("\\")+"{}".format("test.db")
        conn = sqlite3.connect(pwd)
        c = conn.cursor()
        Name_link = self.m_textCtrl16.GetValue()
        try:
            if Name_link =="" :
                wx.MessageBox("链接名称不能为空", "通知", wx.OK | wx.ICON_INFORMATION)
            else:
                c.execute("DELETE FROM test_url WHERE name='%s'" % (Name_link))
                wx.MessageBox("删除成功", "通知", wx.OK | wx.ICON_INFORMATION)
        except:
            wx.MessageBox("操作异常，重新操作", "通知", wx.OK | wx.ICON_INFORMATION)
        conn.commit()
        conn.close()

    def f(self, event):
        self.m_textCtrl16.Clear()

    def g(self, event):
        pwd = os.getcwd()+"{}".format("\\")+"{}".format("test.db")
        conn = sqlite3.connect(pwd)
        c = conn.cursor()
        Url_Name = self.m_textCtrl17.GetValue()
        cursor = c.execute("select name,link from test_url where name='%s' " % Url_Name)
        #print(Url_Name)
        try:
            if Url_Name=="":
                wx.MessageBox("不能为空", "通知", wx.OK | wx.ICON_INFORMATION)
            else:
                select_name = ""
                for row in cursor:
                    #select_name = select_name + str("【链接名称：】"+tuple(row)[1])+"\n"+ str("【地址链接：】"+tuple(row)[2])
                    select_name = select_name + str(tuple(row)[1])+"\n"
                self.m_textCtrl18.SetValue(select_name)
        except:
            wx.MessageBox("异常，重新查询", "通知", wx.OK | wx.ICON_INFORMATION)
        conn.commit()
        conn.close()

    #打开链接
    def h(self, event):
        x=self.m_textCtrl18.GetValue()
        try:
            if x =="" :
                wx.MessageBox("输入地址不能为空", "通知", wx.OK | wx.ICON_INFORMATION)
            else:
                webbrowser.open(x)
        except:
            wx.MessageBox("打开异常，重新打开", "通知", wx.OK | wx.ICON_INFORMATION)

    def i(self, event):
        self.m_textCtrl17.Clear()
        self.m_textCtrl18.Clear()

    def j(self, event):
        pwd = os.getcwd()+"{}".format("\\")+"{}".format("test.db")
        conn = sqlite3.connect(pwd)
        c = conn.cursor()
        try:
            cursor = c.execute(r'select * from test_url ')
            select_name=""
            for row in cursor:
                select_name = select_name+str("【链接名称：】"+tuple(row)[1])+"\n"+ str("【地址链接：】"+tuple(row)[2])+"\n"
                #print(select_name)
            self.m_textCtrl131.SetValue(select_name)
        except:
            wx.MessageBox("没有数据", "通知", wx.OK | wx.ICON_INFORMATION)
        conn.commit()
        conn.close()

    def k(self, event):
        self.m_textCtrl131.Clear()

def main():
    app = wx.App(False)
    frame = MyFrame1(None)
    frame.Show(True)
    app.MainLoop()

if __name__ == '__main__':
    main()