# -*- coding:utf-8 -*-
import sys
import win32com.client as win32
import xlrd
import xlwt
import os
import wx
import codecs
import sqlite3
from threading import Thread


default_encoding = 'GB18030'
if sys.getdefaultencoding() != default_encoding:
    reload(sys)
    sys.setdefaultencoding(default_encoding)

version = 1.0

class TestThread(Thread):
    def __init__(self):
    #线程实例化时立即启动
        Thread.__init__(self)
        self.start()
    def run(self):
     #线程执行的代码
        for i in range(101):
            time.sleep(0.03)
            #wx.CallAfter(Publisher.sendMessage, "update", i)
        time.sleep(0.5)
        #wx.CallAfter(Publisher.sendMessage, "update", u"线程结束")


class iForm(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None, title=u"PartList生成"+str(version)+u" PRODUCT by yanyf&huohm", size = (600, 400))
        panel = iPanel(self, -1)

class iPanel(wx.Panel):
    def __init__(self, parent, id):
        wx.Panel.__init__(self, parent, -1, style=wx.TAB_TRAVERSAL|wx.CLIP_CHILDREN)

        self.logText = wx.TextCtrl(self, -1, "", pos=(130,130), size=(350,220), style = wx.TE_MULTILINE )

        self.defpathStaticText = wx.StaticText(self, -1, u'DEF路径： ', (50, 40))
        self.srcpathStaticText = wx.StaticText(self, -1, u'SRC路径： ', (50, 70))
        self.denpathStaticText = wx.StaticText(self, -1, u'生成位置： ', (50, 100))

        self.defpathText = wx.TextCtrl(self, -1, "", pos=(130,40), size=(350,20))
        self.srcpathText = wx.TextCtrl(self, -1, "", pos=(130,70), size=(350,20))
        self.dsnpathText = wx.TextCtrl(self, -1, "", pos=(130,100), size=(350,20))

        self.defpath_button = wx.Button(self, -1, u'...', pos = (485, 40), size = (20,20))
        self.srcpath_button = wx.Button(self, -1, u'...', pos = (485, 70), size = (20,20))
        self.dsnpath_button = wx.Button(self, -1, u'...', pos = (485, 100), size = (20,20))
        
        self.startButton = wx.Button(self, -1, 'Start!', pos = (50,130), size = (60,60))


        self.Bind(wx.EVT_BUTTON, self.OnShowSrcFiles, self.srcpath_button)
        self.Bind(wx.EVT_BUTTON, self.OnShowDefFiles, self.defpath_button)
        self.Bind(wx.EVT_BUTTON, self.SetDsnPath, self.dsnpath_button)
        self.Bind(wx.EVT_BUTTON, self.OnRun, self.startButton)

        self.Srcpath = ''
        self.Defpath = ''
        self.Dsnpath = ''

    def OnShowSrcFiles(self, event):
    	imessage = "Add  Input files"
    	dlg = wx.DirDialog(self, message=imessage,
            defaultPath=os.getcwd(), 
            style=wx.DD_CHANGE_DIR | wx.DEFAULT_DIALOG_STYLE )
    	if dlg.ShowModal() == wx.ID_OK:
            self.Srcpath = dlg.GetPath()
            self.srcpathText.AppendText(self.Srcpath)
            self.GetFileList(self.Srcpath)
            self.OutFileList()
        dlg.Destroy()

    def OnShowDefFiles(self, event):
    	imessage = "Add  Input files"
    	dlg = wx.DirDialog(self, message=imessage,
            defaultPath=os.getcwd(), 
            style=wx.DD_CHANGE_DIR | wx.DEFAULT_DIALOG_STYLE )
    	if dlg.ShowModal() == wx.ID_OK:
            self.Defpath = dlg.GetPath()
            self.defpathText.AppendText(self.Defpath)
            self.GetFileList(self.Defpath)
            self.OutFileList()
        dlg.Destroy()

    def SetDsnPath(self, event):
    	imessage = "Choose output path"
    	dlg = wx.DirDialog(self, message=imessage,
    		defaultPath=os.getcwd(),
    		style=wx.DD_CHANGE_DIR | wx.DEFAULT_DIALOG_STYLE )
    	if dlg.ShowModal() == wx.ID_OK:
            self.Dsnpath = dlg.GetPath()
            self.dsnpathText.AppendText(self.Dsnpath)
        dlg.Destroy()

    def GetFileList(self, filestr):
    	self.FileList = []
        try:
            FileNames=os.listdir(filestr)
        except Exception, e:
            wx.MessageBox(u'No file exist!'+str(e),'Info',wx.OK|wx.ICON_INFORMATION)
        for EachFile in FileNames:
            if ( os.path.splitext(EachFile)[1][1:] == "xls" \
            or os.path.splitext(EachFile)[1][1:] == "vsd" ):
            #and EachFile in fileList ): 
                self.FileList.append(EachFile)
        if len(self.FileList) <= 0:
            wx.MessageBox(u'No file exist!'+str(e),'Info',wx.OK|wx.ICON_INFORMATION)

    def OutFileList(self):
    	for inum in range(0,len(self.FileList)):
            self.logText.AppendText(str(self.FileList[inum])+'\n')

    def OnRun(self, event):
    	if self.Srcpath == '' or self.Defpath == '':
    		wx.MessageBox(u'Wrong file path!','Info',wx.OK|wx.ICON_INFORMATION)
    	else:
    		self.MakePartListFile(self.Srcpath, self.Defpath, self.Dsnpath)
    		wx.MessageBox(u'完成！',u'哈哈',wx.OK|wx.ICON_INFORMATION)

    def MakePartListFile(self, srcpath, defpath, dsnpath):
    	cx = sqlite3.connect(str(dsnpath)+"\\file.db")
    	try:
            cx.execute('drop table rPart')
            cx.execute('drop table lPart')
        except sqlite3.OperationalError: pass
        finally:
        	cx.execute('''create table rPart (	id varchar(20), 
        		name varchar(50),
        		biaoshi varchar(100),
        	 	shi_1 varchar(100),
        	 	tiaojian_1 varchar(100),
        	 	xiaoqu varchar(100),
        	 	shi_2 varchar(100),
        	 	tiaojian_2 varchar(100),
        	 	'トーンダウン' varchar(100),
        	 	shi_3 varchar(100),
        	 	tiaojian_3 varchar(100),
        	 	'選択' varchar(100),
        	 	'shi_4' varchar(100),
        	 	'tiaojian_4' varchar(100),
        	 	'走行規制' varchar(30),
        	 	'短押L_ON_ユースケース' varchar(5),
        	 	'短押L_ON_時間' varchar(5),
        	 	'短押L_ON_ＢＥＥＰ音' varchar(10),
        	 	'短押L_OFF_ユースケース' varchar(5),
        	 	'短押L_OFF_時間' varchar(5),
        	 	'短押L_OFF_ＢＥＥＰ音' varchar(10),
        	 	'長押L_ON_ユースケース' varchar(5),
        	 	'長押L_ON_時間' varchar(5),
        	 	'長押L_ON_ＢＥＥＰ音' varchar(10),
        	 	'長押L_OFF_ユースケース' varchar(5),
        	 	'長押L_OFF_時間' varchar(5),
        	 	'長押L_OFF_ＢＥＥＰ音' varchar(10),
        	 	'長押L_周期_ユースケース' varchar(5),
        	 	'長押L_周期_時間' varchar(5),
        	 	'長押L_周期_ＢＥＥＰ音' varchar(10),
        	 	'備考' varchar(100))''')
        	cx.execute('''create table lPart(frameName varchar(20),
        		Num varchar(5),
        		id varchar(20))''')
        '''DEF文件夹操作'''
        self.logText.AppendText(u'收集DEF数据'+'\n')    
        fpath = self.search_file(defpath,'xls')
        AlphbetList = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
        excelList = []
        for fname in fpath:
            print fname
            self.logText.AppendText(str(fname)+'\n')    
            data = xlrd.open_workbook(fname)
            table = data.sheets()[0]
            sheetslist = []
            for i in range(0,len(data.sheets())):
            	if 'Old' in data.sheets()[i].name:
            		continue
            	sheetslist.append(i)
            for ii in range(0,len(sheetslist)):
                table = data.sheets()[sheetslist[ii]]      
                l = []
                for k in range(0,30):
                	l.append('-')    
                nrows = table.nrows       
                for i in range(2, nrows):
                	if table.cell(i, 0).value == u'部品ＩＤ':
                		l[0] = table.cell(i, 4).value
                        
                	elif table.cell(i, 0).value == u'表示':
                		j = i + 1
                		l[2] = table.cell(i, 4).value
                		
                		l[3] = ''
                		while table.cell(j, 4).value in AlphbetList:
                			l[3] = l[3] + table.cell(j, 4).value + u':' + table.cell(j, 5).value + '\n'
                			j = j + 1
                		l[1] = l[2] + '\n' + l[3]
                
                	elif table.cell(i, 0).value == u'消去':
                		if table.cell(i+1, 4).value == u'－' or table.cell(i+1, 4).value == u'-':
                			continue
                		j = i + 2
                		l[5] = table.cell(i+1, 4).value
                		l[6] = ''
                		while table.cell(j, 4).value in AlphbetList:
                			if table.cell(j, 5).value == u'－' or table.cell(j, 5).value == u'-':
                				j = j + 1
                				continue
                			else:
                			    l[6] = l[6] + table.cell(j, 4).value + u':' + table.cell(j, 5).value + '\n'
                			    j = j + 1
                		l[4] = l[5] + '\n' + l[6]
                
                	elif table.cell(i, 0).value == u'トーンダウン':
                		if table.cell(i+1, 4).value == u'－' or table.cell(i+1, 4).value == u'-':
                			continue
                		j = i + 2
                		l[8] = table.cell(i+1, 4).value
                		l[9] = ''
                		while table.cell(j, 4).value in AlphbetList:
                			if table.cell(j, 5).value == u'－' or table.cell(j, 5).value == u'-':
                				j = j + 1
                				continue
                			else:
                			    l[9] = l[9] + table.cell(j, 4).value + u':' + table.cell(j, 5).value + '\n'
                			    j = j + 1
                		l[7] = l[8] + '\n' + l[9]
                
                	elif table.cell(i, 0).value == u'走行中トーンダウン' or table.cell(i, 0).value == u'走行中消去':
                		l[13] = table.cell(i, 4).value
                
                	elif table.cell(i, 0).value == u'インジケータ':
                		if table.cell(i+1, 4).value == u'－' or table.cell(i+1, 4).value == u'-':
                			continue
                		j = i + 2
                		l[11] = table.cell(i+1, 4).value
                		l[12] = ''
                		while table.cell(j, 4).value in AlphbetList:
                			if table.cell(j, 5).value == u'－' or table.cell(j, 5).value == u'-':
                				j = j + 1
                				continue
                			else:
                			    l[12] = l[12] + table.cell(j, 4).value + u':' + table.cell(j, 5).value + '\n'
                			    j = j + 1
                		l[10] = l[11] + '\n' + l[12]
                
                	elif table.cell(i, 0).value == u'短押し（ON確定）':
                		l[14] = table.cell(i, 4).value
                		l[15] = table.cell(i+1, 4).value
                		l[16] = table.cell(i+3, 4).value
                
                	elif table.cell(i, 0).value == u'短押し（OFF確定）':
                		l[17] = table.cell(i, 4).value
                		l[18] = table.cell(i+1, 4).value
                		l[19] = table.cell(i+3, 4).value
                
                	elif table.cell(i, 0).value == u'長押し（ON確定）':
                		l[20] = table.cell(i, 4).value
                		l[21] = table.cell(i+1, 4).value
                		l[22] = table.cell(i+3, 4).value
                
                	elif table.cell(i, 0).value == u'長押し（OFF確定）':
                		l[23] = table.cell(i, 4).value
                		l[24] = table.cell(i+1, 4).value
                		l[25] = table.cell(i+3, 4).value
                
                	elif table.cell(i, 0).value == u'長押し（周期確定）':
                		l[26] = table.cell(i, 4).value
                		l[27] = table.cell(i+1, 4).value
                		l[28] = table.cell(i+3, 4).value
                
                	elif table.cell(i, 0).value == u'備考':
                		l[29] = table.cell(i, 4).value
            
                for i in range(2,nrows):
                	if table.cell(i, 0).value == u'部品名称':
                		l.insert(1,table.cell(i, 4).value)
                if len(l) > 30:
                	excelList.append(l)

        '''数据库中插入def数据'''
        for i in range(0, len(excelList)):
        	cx.execute("insert into rPart values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", tuple(excelList[i]))
        cx.commit()

        '''SCR文件夹操作'''
        fpath = self.search_file(srcpath,'.vsd')
        frameList = []
        frameName = []
        self.logText.AppendText(u'收集SCR数据'+'\n')    
        for fname in fpath:
	        print fname
	        self.logText.AppendText(str(fname)+'\n')    
	        visio = win32.gencache.EnsureDispatch('Visio.Application')
	        vs = visio.Documents.Open(fname)
        
	        for k in range(1, len(vs.Pages)+1):
	        	print vs.Pages(k).Name
	        	if k==1:
	        		frameName.append(vs.Pages(k).Name[0:-3])
	        	if vs.Pages(k).Name != u'BackGro' and k>1 and vs.Pages(k).Name[0:-3] != frameName[-1]:
	        		frameName.append(vs.Pages(k).Name[0:-3])
	        	if frameName[-1] == u'BackGro':
	        		frameName.pop()
	        	for i in range(1, vs.Pages(k).Shapes.Count):
	        		pageList = []
	        		pageList.append(vs.Pages(k).Name[0:-3])
	        		try:
	        			if(str(vs.Pages(k).Shapes(i).Name).find("List")==0):
	        				if 'Switch' in vs.Pages(k).Shapes(i).Name:
	        				    print vs.Pages(k).Shapes(i).Name
	        				    print vs.Pages(k).Shapes(i).Cells('Prop.SWPartsNum').ResultStr()
	        				    print vs.Pages(k).Shapes(i).Cells('Prop.SWPartsID').ResultStr()
	        				    pageList.append(vs.Pages(k).Shapes(i).Cells('Prop.SWPartsNum').ResultStr())
	        				    pageList.append(vs.Pages(k).Shapes(i).Cells('Prop.SWPartsID').ResultStr())
	        				    frameList.append(pageList)
	        				elif 'Design' in vs.Pages(k).Shapes(i).Name:
	        				    print vs.Pages(k).Shapes(i).Name
	        				    print vs.Pages(k).Shapes(i).Cells('Prop.DesignPartsNum').ResultStr()
	        				    print vs.Pages(k).Shapes(i).Cells('Prop.DesignPartsID').ResultStr()
	        				    pageList.append(vs.Pages(k).Shapes(i).Cells('Prop.DesignPartsNum').ResultStr())
	        				    pageList.append(vs.Pages(k).Shapes(i).Cells('Prop.DesignPartsID').ResultStr())
	        				    frameList.append(pageList)
        
	        		except Exception, e:
	        			print "error:"+str(e)+"  shape:"+vs.Pages(k).Shapes(i).Name
	        			continue
	        vs.Close()
	        visio.Application.Quit()
	        visio=None
    
		'''数据库中插入src数据'''
        for i in range(0,len(frameList)):
        	cx.execute("insert into lPart values (?,?,?)", tuple(frameList[i]))
        cx.commit()

        '''数据库查询'''
        cu = cx.execute("Select * from rPart")
        finalList = []
        print frameName
        for name in frameName:
        	print name
        	sql_cmd1="Select frameName,Num,rPart.* from lPart,rPart where frameName=\'"+name+"\' and lPart.id=rPart.id"
        	print sql_cmd1
        	cu = cx.execute(sql_cmd1)
        	for row in cu:
        		finalList.append(list(row))
        
        cx.close()

        '''从数据库中取数据放入表中'''
        self.logText.AppendText(u'写入文件'+'\n')  
        w = xlwt.Workbook()
        ws = w.add_sheet(u'sheet1')
        w.save(str(dsnpath)+'\\PartList.xls')  
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(str(dsnpath)+'\\PartList.xls')
        '''格式处理'''
        firstLine = [u'画面ID',u'部品番号',u'部品ID',u'部品名称',u'表示',u'式',u'条件',u'消去',u'式',u'条件',
            u'トーンダウン',u'式',u'条件',u'選択',u'式',u'条件',u'走行規制\n※部品が\nIM(意匠)の場合：走行中消去\nTSW（釦）の場合：走行中TD\nを表す。',
            u'ユースケース',u'時間（長押し時）',u'ＢＥＥＰ音',u'ユースケース',u'時間（長押し時）',u'ＢＥＥＰ音',u'ユースケース',u'時間（長押し時）',u'ＢＥＥＰ音',
            u'ユースケース',u'時間（長押し時）',u'ＢＥＥＰ音',u'ユースケース',u'時間（長押し時）',u'ＢＥＥＰ音',u'備考']
        for i in range(1,3):
            for j in range(1,len(firstLine)+1):
            	wb.Worksheets[1].Cells(i,j).Font.Name = u'ＭＳ Ｐゴシック'
            	wb.Worksheets[1].Cells(i,j).Font.Size = 11
            	wb.Worksheets[1].Cells(i,j).HorizontalAlignment = -4108
            	if i == 1:
            		if j == 18:
            			wb.Worksheets[1].Cells(i,j).Value = u'短押し（ON確定）'
            			wb.Worksheets[1].Cells(i,j).Interior.ColorIndex = 43
            			wb.Worksheets[1].Range('R1:T1').MergeCells = True
            			wb.Worksheets[1].Range('R1:T1').Borders.LineStyle = True
            		if j == 21:
            			wb.Worksheets[1].Cells(i,j).Value = u'短押し（OFF確定）'
            			wb.Worksheets[1].Cells(i,j).Interior.ColorIndex = 43
            			wb.Worksheets[1].Range('U1:W1').MergeCells = True
            			wb.Worksheets[1].Range('U1:W1').Borders.LineStyle = True
            		if j == 24:
            			wb.Worksheets[1].Cells(i,j).Value = u'長押し（ON確定）'
            			wb.Worksheets[1].Cells(i,j).Interior.ColorIndex = 43
            			wb.Worksheets[1].Range('X1:Z1').MergeCells = True
            			wb.Worksheets[1].Range('X1:Z1').Borders.LineStyle = True
            		if j == 27:
            			wb.Worksheets[1].Cells(i,j).Value = u'長押し（OFF確定）'
            			wb.Worksheets[1].Cells(i,j).Interior.ColorIndex = 43
            			wb.Worksheets[1].Range('AA1:AC1').MergeCells = True
            			wb.Worksheets[1].Range('AA1:AC1').Borders.LineStyle = True
            		if j == 30:
            			wb.Worksheets[1].Cells(i,j).Value = u'長押し（周期確定）'
            			wb.Worksheets[1].Cells(i,j).Interior.ColorIndex = 43
            			wb.Worksheets[1].Range('AD1:AF1').MergeCells = True
            			wb.Worksheets[1].Range('AD1:AF1').Borders.LineStyle = True
            	if i == 2:
            		if j > 17 and j < 33:
            			wb.Worksheets[1].Cells(i,j).Interior.ColorIndex = 43
            		else:
            			wb.Worksheets[1].Cells(i,j).Interior.ColorIndex = 6
            		wb.Worksheets[1].Cells(i,j).Value = firstLine[j-1]
            		wb.Worksheets[1].Cells(i,j).Borders.LineStyle = True
        k=1
        for i in range(0,len(finalList)):
        	self.logText.AppendText(u'写入'+str(finalList[i][0])+str(finalList[i][2])+'\n')
        	for j in range(0,len(finalList[i])):
        		wb.Worksheets[1].Cells(2+k,j+1).HorizontalAlignment = -4108
        		wb.Worksheets[1].Cells(2+k,j+1).Font.Name = u'ＭＳ Ｐゴシック'
        		wb.Worksheets[1].Cells(2+k,j+1).Font.Size = 11
        		wb.Worksheets[1].Cells(2+k,j+1).Borders.LineStyle = True
        		wb.Worksheets[1].Cells(2+k, j+1).Value = finalList[i][j]
        		#wb.Worksheets[1].Cells(2+k, j+1).HorizontalAlignment = 'xlCenter'
        	k = k + 1
        wb.Save()
        wb.Close()
        excel.Application.Quit()

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
                if name[-1*len(file_type):] == file_type:
                       fpath.append(full_path)
        return fpath
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