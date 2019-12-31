import csv
import xlrd
import time
import wx
import os
import json
from xlutils.copy import copy


class SiteLog(wx.Frame):
    # 构造函数
    def __init__(self):
        wx.Frame.__init__(self, None, title='SiteLog', size=(640, 550))
        self.SelBtn = wx.Button(self, label='打开文件', pos=(100, 5), size=(80, 25))
        self.SelBtn.Bind(wx.EVT_BUTTON, self.open_file)
        self.RemoveBtn = wx.Button(self, label='移除文件', pos=(200, 5), size=(80, 25))
        self.RemoveBtn.Bind(wx.EVT_BUTTON, self.remove_file)
        self.OkBtn = wx.Button(self, label='OK', pos=(405, 5), size=(80, 25))
        self.OkBtn.Bind(wx.EVT_BUTTON, self.save_file)
        self.fileNames = []
        self.FileContent = wx.ListBox(self, pos=(5, 35), size=(620, 480), choices=[], style=wx.LB_SINGLE)

        # 统计sku数量
        self.sku_obj = {}

    # 打开对应方法
    def open_file(self, event):
        wildcard = 'csv files(*.csv)|*.csv'
        dialog = wx.FileDialog(None, 'select', os.getcwd(), '', wildcard, wx.FD_MULTIPLE)
        if dialog.ShowModal() == wx.ID_OK:
            file_names = dialog.GetPaths()
            self.fileNames.extend(file_names)
            self.FileContent.SetItems(file_names)
            dialog.Destroy()

    # 移除文件方法
    def remove_file(self, event):
        selection = self.FileContent.GetSelection()
        if selection is not None and selection >= 0:
            self.fileNames.pop(selection)
            self.FileContent.SetItems(self.fileNames)

    # 保存方法
    def save_file(self, event):
        if len(self.fileNames) == 0:
            return
        wildcard = 'excel files(*.xls)|*.xls'
        current_time = time.strftime("%m.%d", time.localtime())
        dialog = wx.FileDialog(self, message="保存文件", wildcard=wildcard, style=wx.FD_SAVE,
                               defaultFile='%s 义蓬UNNY美妆' % (current_time,))
        result = dialog.ShowModal()
        if result != wx.ID_OK:
            return
        path = dialog.GetPath()

        obj_wb = xlrd.open_workbook(r'model.xls', formatting_info=True)
        new_wb = copy(obj_wb)
        new_sht = new_wb.get_sheet(0)

        # 数据拼接
        i = 1
        for item0 in self.fileNames:
            with open(item0) as file:
                # reader为迭代类型
                reader = csv.reader(file)
                j = 0
                for item1 in reader:
                    if j > 0:
                        # 名字
                        new_sht.write(i, 1, item1[14])
                        # 电话
                        new_sht.write(i, 2, item1[15])
                        # 地址
                        new_sht.write(i, 3, item1[17] + item1[18] + item1[19] + item1[20])
                        # 日用品
                        new_sht.write(i, 4, '日用品')
                        # sku
                        new_sht.write(i, 5, item1[27])
                        # 备注
                        new_sht.write(i, 6, item1[38])
                        # 统计数量
                        if self.sku_obj.__contains__(item1[27]):
                            sku_num = self.sku_obj[item1[27]] + 1
                        else:
                            sku_num = 1
                        self.sku_obj[item1[27]] = sku_num
                        # 外部索引
                        i += 1
                    j += 1
        new_wb.save(path)
        print('写入成功')
        total = "共%s个，%s" % (i - 1, json.dumps(self.sku_obj, ensure_ascii=False))
        print(total)
        dlg = wx.MessageDialog(None, json.dumps(total, ensure_ascii=False), u"统计", wx.YES_NO | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            self.Close(True)
        dlg.Destroy()


if __name__ == '__main__':
    app = wx.App()
    SiteFrame = SiteLog()
    SiteFrame.Show()
    app.MainLoop()
