# -*- coding:utf-8 -*-
import sys
import win32com.client as win32
import xlrd
import os
import wx
import wx.calendar as cal

default_encoding = 'gbk'
if sys.getdefaultencoding() != default_encoding:
    reload(sys)
    sys.setdefaultencoding(default_encoding)

version="1.00"
test_date = 0
choice1 = 0
choice2 = 0

versionList = ['0.91','0.92','0.93','0.94','0.95','0.96','0.97','0.98','0.99']
testerList = [u'史建航', u'董森', u'韩伟强', u'任晓莉', u'李婷婷', u'范倩雯', u'董爽', u'仲诗禹']
fileList = [
 u'SX5_HMI_测试项目_AIR.xlsx',
 u'SX5_HMI_测试项目_BT music.xlsx',
 u'SX5_HMI_测试项目_BT Pairing.xlsx',
 u'SX5_HMI_测试项目_BT_Calls.xlsx',
 u'SX5_HMI_测试项目_CAN Settings.xlsx',
 u'SX5_HMI_测试项目_CarPlay.xlsx',
 u'SX5_HMI_测试项目_Engineering Mode.xlsx',
 u'SX5_HMI_测试项目_General.xlsx',
 u'SX5_HMI_测试项目_Home.xlsx',
 u'SX5_HMI_测试项目_IPOD.xlsx',
 u'SX5_HMI_测试项目_Link.xlsx',
 u'SX5_HMI_测试项目_Maintenance.xlsx',
 u'SX5_HMI_测试项目_PDC.xlsx',
 u'SX5_HMI_测试项目_PhoneContacts.xlsx',
 u'SX5_HMI_测试项目_Power_Moding.xlsx',
 u'SX5_HMI_测试项目_RADIO.xlsx',
 u'SX5_HMI_测试项目_Setting.xlsx',
 u'SX5_HMI_测试项目_SWDL.xlsx',
 u'SX5_HMI_测试项目_USB.xlsx',
 u'SX5_HMI_测试项目_VR.xlsx'
]

class Calendar(wx.Dialog):
    def __init__(self, parent, id, title):
        wx.Dialog.__init__(self, parent, id, title, size=(340, 240))
        
        
        self.datectrl =parent.datectrl

        vbox = wx.BoxSizer(wx.VERTICAL)
        
        calend = cal.CalendarCtrl(self, -1, wx.DateTime_Now(), \
                                  style = cal.CAL_SHOW_HOLIDAYS|\
                                  cal.CAL_SEQUENTIAL_MONTH_SELECTION)
        vbox.Add(calend, 0, wx.EXPAND|wx.ALL, 20)
        self.Bind(cal.EVT_CALENDAR, self.OnCalSelected, \
                  id=calend.GetId())

        vbox.Add((-1, 20))
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        vbox.Add(hbox, 0, wx.LEFT, 8)       
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        vbox.Add(hbox2, 0, wx.ALIGN_CENTER|wx.TOP|wx.BOTTOM, 20)     
        self.SetSizer(vbox)
        self.Show(True)
        self.Center()


    def OnCalSelected(self, event):
        global test_date
        date = str(event.GetDate())[:-9]
        date =  "20" +date[date.rfind("/")+1:]+'/'+date[:-3]
        test_date = date
        print test_date
        self.datectrl.SetLabel(str(date))
        self.Destroy()
            

class iForm(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None, title=u"数据表格处理"+version, size = (800, 600))
        panel = iPanel(self, -1)


class iPanel(wx.Panel):
    def __init__(self, parent, id):
        wx.Panel.__init__(self, parent, -1, style=wx.TAB_TRAVERSAL|wx.CLIP_CHILDREN)

        self.listctrl = wx.ListCtrl(self, -1, pos=(0,0), size=(500,600),style=wx.LB_SINGLE )#wx.LC_NO_HEADER
        self.listctrl.InsertColumn(0, u"序列",width=60)
        self.listctrl.InsertColumn(1, u"输入文件名",width=380)
        self.listctrl.InsertColumn(2, u"状态",width=60)


        self.dirpath_StaticText=wx.StaticText(self, -1, u"作业路径:", (540, 20))
        self.searfile_button = wx.Button(self, -1,u'打开文件夹', pos=(550, 50),size = (90, 60))
        self.irun_button = wx.Button(self, -1,u'开始工作！！', pos=(550, 120),size = (90, 60))
        self.idate_button = wx.Button(self, -1,u'日期', pos=(540, 255),size = (60, 20))
        
        wx.StaticText(self, -1, u"测试版本:", (540, 225))
        wx.StaticText(self, -1, u"测试人:", (540, 195))

        self.TesterChoice = wx.Choice(self, -1, (610, 190), choices=testerList)
        self.VersionChoice = wx.Choice(self, -1, (610, 220), choices=versionList)

        self.TesterChoice.Bind(wx.EVT_CHOICE, self.onTesterList)
        self.VersionChoice.Bind(wx.EVT_CHOICE, self.onVersionList)

        self.Bind(wx.EVT_BUTTON,self.OnAddLocalWork,self.searfile_button)
        self.Bind(wx.EVT_BUTTON,self.OnRun,self.irun_button)
        self.Bind(wx.EVT_BUTTON,self.OnDate,self.idate_button)

        self.datectrl=wx.StaticText(self, -1, "", pos=(600, 255)) 

        self.logText = wx.TextCtrl(self, -1, "", pos=(540, 300),size = (200,200),style = wx.TE_MULTILINE )

    def onTesterList(self,event):
        global choice1
        choice1 = self.TesterChoice.GetSelection()
        print testerList[choice1]

    def onVersionList(self,event):
        global choice2
        choice2 = self.VersionChoice.GetSelection()
        print versionList[choice2]

    def OnDate(self,event):
        mydate = Calendar(self,-1,u'请双击选择日期')


    def OnAddLocalWork(self,event):
        if self.listctrl.GetItemCount() > 0:
            self.ClearList()
        imessage="Add  Input Excel files"
        dlg = wx.DirDialog(self, message=imessage,
            defaultPath=os.getcwd(), 
            style=wx.DD_CHANGE_DIR | wx.DEFAULT_DIALOG_STYLE )
        if dlg.ShowModal() == wx.ID_OK:
            self.path = dlg.GetPath()
            self.dirpath_StaticText.SetLabel(self.path)
            self.GetFileList(self.path)
            self.OutPutFileList()
        dlg.Destroy()


    def OnRun(self,event):
        global test_date
        global choice2
        global choice1
        if self.listctrl.GetItemCount() <= 0:
            wx.MessageBox(u'No file exist!','Info',wx.OK|wx.ICON_INFORMATION)
            return
        print self.path
        self.doExcel(self.path,testerList[choice1],versionList[choice2],test_date)
        wx.MessageBox(u'完成！',u'哈哈',wx.OK|wx.ICON_INFORMATION)

    def GetFileList(self,filestr):
        self.FileList = []
        try:
            FileNames=os.listdir(filestr)
        except Exception, e:
            wx.MessageBox(u'No file exist!'+str(e),'Info',wx.OK|wx.ICON_INFORMATION)
        for EachFile in FileNames:
            if ( os.path.splitext(EachFile)[1][1:] == "xls" \
            or os.path.splitext(EachFile)[1][1:] == "xlsx" \
            and EachFile in fileList ): 
                self.FileList.append(EachFile)
        if len(self.FileList) <= 0:
            wx.MessageBox(u'No file exist!'+str(e),'Info',wx.OK|wx.ICON_INFORMATION)

    def OutPutFileList(self):
        for inum in range(0,len(self.FileList)):
            self.listctrl.InsertStringItem(inum, str(inum+1))
            self.listctrl.SetStringItem(inum, 1, self.FileList[inum])
            self.listctrl.SetStringItem(inum, 2, u"×")


    def doexcel_row_abc(self,number):
    
        if number%26==0 and number != 26:
            return chr(64+number/26-1)+'Z'
        else:
            return chr(64+number/26)+chr(64+number%26)
    
            
    def search_file(self,path,file_type):  
        queue = []
        queue.append(path);
        fpath=[]
        while len(queue) > 0:  
            tmp = queue.pop(0)  
            if(os.path.isdir(tmp)):  
                for item in os.listdir(tmp):  
                    queue.append(os.path.join(tmp, item))  
            elif(os.path.isfile(tmp)):   
                name= os.path.basename(tmp)
                dirname= os.path.dirname(tmp)
                full_path = os.path.join(dirname,name)
                abspath=os.path.abspath(tmp);
                if name[-1*len(file_type):] == file_type and name in fileList:
                       fpath.append(full_path)
        return fpath
    
    def doExcel(self,fpath,tester,test_version,test_date):
        ####################################################
        fpath = self.search_file(fpath,'xlsx')
        file_num=0
        for k in fpath:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            fname = k
            data = xlrd.open_workbook(fname)
            wb = excel.Workbooks.Open(fname)
            #文件是否是测试文件
            self.logText.AppendText(k+"\n")
            print k
            for i in range(4,len(wb.Worksheets)+1):
                table = data.sheets()[i-1]
                nrows = table.nrows+1   #随后一个输入文字的行
                ncols = table.ncols   #最后一个有文字的列
                for j in range(nrows,10,-1):
                     if(wb.Worksheets[i].Cells(j,ncols-5).Value!=None):
                         nrows=j
                         break
                ncols = ncols +1 #在下一个列开始复制
                self.logText.AppendText(wb.Worksheets[i].Name+"\n")
                self.logText.AppendText(u"行数:"+str(nrows)+"\n")
                print wb.Worksheets[i].Name
                print "ncols:"+str(ncols)
                print "nrows:"+str(nrows)
                #if(nrows >= 685):
                #    wb.Worksheets[i].Range(self.doexcel_row_abc(ncols-6)+'10:'+self.doexcel_row_abc(ncols-1)+'685').Copy()
                #    wb.Worksheets[i].Range(self.doexcel_row_abc(ncols)+'10').PasteSpecial()
                #    wb.Worksheets[i].Range(self.doexcel_row_abc(ncols-6)+'686:'+self.doexcel_row_abc(ncols-1)+str(nrows)).Copy()
                #    wb.Worksheets[i].Range(self.doexcel_row_abc(ncols)+'686').PasteSpecial()
                #else:
                if(nrows<600):
                    if(wb.Worksheets[i].Cells(13,ncols-4).Value!=None):
                        copy_range =  self.doexcel_row_abc(ncols-6)+'10:'+self.doexcel_row_abc(ncols-1)+str(nrows)
                        #print u"拷贝区域"+copy_range
                        wb.Worksheets[i].Range(copy_range).Copy()
                        wb.Worksheets[i].Range(self.doexcel_row_abc(ncols)+'10').PasteSpecial()
                #wb.Worksheets[i].Range(self.doexcel_row_abc(ncols-6)+'9:'+self.doexcel_row_abc(ncols-1)+str(nrows)).Copy()
                #wb.Worksheets[i].Range(self.doexcel_row_abc(ncols)+'9').PasteSpecial()
                else:
                    for nr in range(9,nrows+1):
                        wb.Worksheets[i].Range(self.doexcel_row_abc(ncols-6)+str(nr)+':'+self.doexcel_row_abc(ncols-1)+str(nr)).Copy()
                        wb.Worksheets[i].Range(self.doexcel_row_abc(ncols)+str(nr)).PasteSpecial()
                for j in range(12,nrows+1):
                    if(wb.Worksheets[i].Cells(j,ncols-6).Value!=None):
                        wb.Worksheets[i].Cells(j,ncols).Value = test_version
                        wb.Worksheets[i].Cells(j,ncols+1).Value = test_date
                        wb.Worksheets[i].Cells(j,ncols+2).Value = tester
                self.logText.AppendText("OK"+"\n")
                print "OK"
            self.listctrl.SetStringItem( file_num , 2, u"√")
            file_num=file_num+1
            wb.Save()
            wb.Close()   
            excel.Application.Quit()
        return 0


class iApp(wx.App):
    """Application class."""
    def __init__(self):
        wx.App.__init__(self, 0)
        return None
    def OnInit(self):
        self.MainFrame = iForm()
        self.MainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = iApp()
    app.MainLoop()