#coding: utf-8

#import interesting things
import xlrd,re,datetime

#set up our own params
showdate=datetime.date(2022,12,1)
"""
TODO:

1 for all
2 for 
"""
#whichstudyincluded=

#initialize necessary vars
filename=r"a.xlsx" #const
sheet_name="成员参会明细" #const
firstrow=9 #const
cols=5 #const
datere = re.compile(u'20[0-9][0-9]-[0-1][0-9]-[0-3][0-9] ') #const
timere = re.compile(u'[0-2][0-9]:[0-5][0-9]:[0-5][0-9]') #const
namere = re.compile(u"[0-9][0-9][\u4E00-\u9FA5][\u4E00-\u9FA5][\u4E00-\u9FA5][\u4E00-\u9FA5]|[0-9][0-9][\u4E00-\u9FA5][\u4E00-\u9FA5][\u4E00-\u9FA5]|[0-9][0-9][\u4E00-\u9FA5][\u4E00-\u9FA5]")#const to make name process as pin-point accuracy
book=xlrd.open_workbook(filename).sheet_by_name(sheet_name) #load table
totalrow=book.nrows
timetable=[ #classes const
            ["8:00","8:40"],  #first
            ["8:50","9:30"],  #second
            ["10:00","10:40"],#third
            ["10:50","11:30"],#fourth
            ["11:40","12:20"],#fifth
            ["14:30","15:10"],#sixth
            ["15:20","16:00"],#seventh
            ["16:15","16:35"],#eighth
            ["17:05","17:30"],#ninth
            ["17:40","18:20"] #tenth
            ]
morningstudy=["7:30","7:50"] #optional timetable
nightstudy=["19:00","20:50"] #optional timetable
noonstudy=["12:40","14:30"]  #optional timetable
weekendexamtable=nightstudy #of course, optional timetable
fridayexamtable=["17:05","18:05"]#default, set by your self
longlimit=datetime.datetime.strptime("05:00","%H:%M").time()
latetotal=[]
quittotal=[]
lastclass=[]

#let's begin with circlulation
for i in range(1,totalrow-firstrow):#early 9 rows excluded
    table=book.row(firstrow+i)#trans row into tuple cells
    name1=namere.findall(str(re.sub('text:', '', str(table[0]))))#not processed raw name
    if name1==[]:#teacher(who compare regex fail)
        continue #teachers are excluded
    else:#student
        name=str(name1[0])#process student name
    jointime=datetime.datetime.strptime(str(re.sub('text:', '', str(table[1]))), "'%Y-%m-%d %H:%M:%S'")
    #if str(table[2])!="'--'":
    #    quittime=datetime.datetime.strptime(str(re.sub('text:', '', str(table[2]))), "'%Y-%m-%d %H:%M:%S'")
    #else:
    #    quittime='--'
    quittime='--'
    long='--'
    #as datetime cannot define this
    #long=datetime.datetime.strptime(str(re.sub('text:', '', str(table[3]))), "'%H:%M:%S'")
    #long=quittime-jointime if str(table[2]!='--') else '--'
    #roomtype=table[4] unused
    #print("姓名：{},加入时间：{},退出时间：{},参会时长：{}".format(name,jointime,quittime,long))
    #if(jointime.date()!=showdate):
    #    continue
    '''
    TODO:
    Fix these shitty code however no more online classes now so i won't fix them
    
    else:
        for classes in range(10):
                if(classes==9):
                    pass
                else:
                    nextclassstart=datetime.datetime.strptime(showdate.strftime("%Y-%m-%d ")+timetable[classes+1][0]+":0","%Y-%m-%d %H:%M:%S")
                    nextclassend=datetime.datetime.strptime(showdate.strftime("%Y-%m-%d ")+timetable[classes+1][1]+":0","%Y-%m-%d %H:%M:%S")
                starttime=datetime.datetime.strptime(showdate.strftime("%Y-%m-%d ")+timetable[classes][0]+":0","%Y-%m-%d %H:%M:%S")
                endtime=datetime.datetime.strptime(showdate.strftime("%Y-%m-%d ")+timetable[classes][1]+":0","%Y-%m-%d %H:%M:%S")
                duringtime=jointime-starttime
                
                if(endtime-quittime<datetime.timedelta(seconds=2400) and endtime-quittime>datetime.timedelta(seconds=0)):
                        print("表格{}行：{}在第{}节课中途退出，此时离下课还有{}".format(i+firstrow+1,name,classes+1,endtime-quittime))
                        #TODO: process who exit during classsssss if()
                        for cur in range(i,totalrow-firstrow):
                            if()
                if(duringtime>datetime.timedelta(seconds=0) and duringtime<=datetime.timedelta(seconds=2400)):
                    #if exitduringclass:
                    #    print("")
                    print("表格{}行：{}在第{}节课迟到{}分钟".format(i+firstrow+1,name,classes+1,duringtime))
    '''
    with open("output.csv","wb") as file:
        file.writeline("{},{},{},{},{},{},{},{}".format(name,type(name),jointime,type(jointime),quittime,type(quittime),long,type(long)).encode())
