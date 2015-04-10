#encoding=cp936
import ConfigParser
import datetime
import random
import re
import requests
from BeautifulSoup import BeautifulSoup
import PAMIE
import sys

s = requests.session()
bizTravelID = "1321_16645" #出差项目编号，报工页面写死的
WORK_HOURS = 7.9
OT_WORK_HOURS = 4
NORMAL_TYPE = 0
OT_TYPE = 1

#读取配置文件
config = ConfigParser.ConfigParser()
config.read("config.ini")
userName = config.get("global", "userName")
passWord = config.get("global", "passWord")
projectID = config.get("global", "projectID")
projectNames = config.get("global", "projectName").split(',')
projectCnt = len(projectNames)
host = config.get("global", "host")
holidays = config.get("global", "holidays").split(',')
inputIDs = []

base_form = {}

def getLastWeekDate():
	global startDate
	global endDate
	global base_form

	today = datetime.date.today()
	startDate = today + datetime.timedelta(-7 - today.weekday()) #获取上周一的日期
	endDate = startDate + datetime.timedelta(6)
	base_form = {'startDate': startDate, 'endDate': endDate}
	print "报工开始日期: " + str(startDate) +  " 报工结束日期: " + str(endDate)
	

def sendToServer(path, form_data, desc="undefined", params=""):
    response = s.post(host + path, data = form_data, params = params)
    if cmp(desc, "undefined") == 0:
        return response

    if cmp(response.content, "success") == 0:
        print desc + u"成功"
    elif response.content == "PasswordOutOfDate": #登陆密码超时
        form_data = {'loginName': userName,
            'oldPassword': passWord,
            'password': passWord,
            'password1': passWord,
            'action':'resetPassword',
        }
        sendToServer('forgetPasswordAction.do', form_data, '重置密码');
    else:
        print desc + "失败"
        print response.content.decode('gbk')
        exit()

#登录
def login():
	login_data = {'loginName': userName,
		'password': passWord,
		'action':'checkOnline',
	}
	sendToServer('loginAction.do', login_data, u"登陆")

#增加工作项
def addProject(projectName):
    form_data={'action':'addmyworkdes',
        'taskName': projectName,
        'projectID':projectID}
    form_data.update(base_form)
    sendToServer('mywork/timesheet/addMyWorkDesAction.do', form_data, "添加项目" + projectName)

#添加出差
def addBizTravel():
	xml_data='''<?xml version="1.0" encoding="gb2312"?>
	<TreeTable xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="TreeTable.xsd" high="20" width="125" color="black" detailable="true" standardmenu="true">
	<TableData>
	<RowData><name>出差（正常栏填0.1）</name><NodeData nodeid="T_1321_16645_4408" nodetypeid="1321_16645_4408" image="29" />
	<CellData><Cell colid="ck" type="boolean">True</Cell><Cell colid="t1" /><Cell colid="t2" /></CellData></RowData>
	</TableData></TreeTable>'''
	params = {'startDate': startDate, 'endDate': endDate}
	sendToServer('mywork/timesheet/addMyTaskAction.do', xml_data, "添加出差", params)

def getWeekOfDay(desc):
    weekOfDay = {u'星期一': '0',
                 u'星期二': '1',
                 u'星期三': '2',
                 u'星期四': '3',
                 u'星期五': '4',
                 u'星期六': '5',
                 u'星期天': '6',
                 }
    return weekOfDay.get(desc[:3])

def getHolidaysWeekOfDay(page):
    result = set()
    for holiday in holidays:
		 result.add(getWeekOfDay(page.find(name="td", text=re.compile(holiday))))
    return result

#查看报工页面
def parsePage():
	response = sendToServer('mywork/timesheet/initTimeSheetAction.do', base_form)
	#查看出差一栏是否已添加:
	if bizTravelID not in response.content:
		addBizTravel()
	#查看项目一栏是否添加
	for projectName in projectNames:
		if projectName not in response.content:
			addProject(projectName)
	
	response = sendToServer('mywork/timesheet/initTimeSheetAction.do', base_form)
	soup = BeautifulSoup(response.content, fromEncoding="GBK")

	#查看是否配置文件如果有节假日日期，不填工时:
	ignoreDays = getHolidaysWeekOfDay(soup)

	#根据之前添加的项目id和inputtext_timesheet获取输入框的id：
	inputIDs = [x.get('name') for x in soup.findAll(name="input", attrs={"name": re.compile(projectID), "class" : re.compile("inputtext")})]
	fillWorkHour(inputIDs, ignoreDays)

	biz_travel_input_ids = [x.get('name') for x in soup.findAll(name="input", attrs={"name": re.compile(bizTravelID), "class" : re.compile("inputtext")})]
	fill_travel_hour(biz_travel_input_ids, ignoreDays)

def filter_holiday(holidays_arg, input_ids, request_form):
    result = []
    for input_id in input_ids:
        if input_id[-1] in holidays_arg:
            request_form[input_id] = ''
        else:
            result.append(input_id)
            
    return result, request_form

#根据项目数随机分配工时
def decomposition(remainHour):
	cnt = projectCnt
	while cnt > 0:
		if cnt == 1:
			yield round(remainHour, 1)
		else:
			n = round(random.uniform(0, remainHour), 1)
			yield n
			remainHour -= n
		cnt -= 1

#获取最后一位字符
def getLastChar(s):
	return s[-1]

#获取工时
def getHour(hours, index, workType):
    origin_index = index
    if origin_index >= projectCnt:
        origin_index = 0
        index = 0
        if workType == NORMAL_TYPE:
            hours = list(decomposition(WORK_HOURS))
        else:
            hours = list(decomposition(OT_WORK_HOURS))
    index += 1

    try:
        hour = hours[origin_index]
    except IndexError:
        hour = ''
    return (hours, hour, index)

#判断是否加班列
def isOtWork(content):
	if "gxot" in content:
		return True
	else:
		return False

#填写出差工时
def fill_travel_hour(biz_travel_input_ids, holidays_arg=set()):
    biz_hours_map = {}
    biz_hours_map.update(base_form)
    first_flag = True

    if holidays:
        biz_travel_input_ids, biz_hours_map = filter_holiday(holidays_arg, biz_travel_input_ids, biz_hours_map)

    for bizID in biz_travel_input_ids:
        if first_flag:
            first_flag = False
            pattern = re.compile('(' + bizTravelID + '.*)_')
            biz_hours_map['ptt'] = pattern.search(bizID).groups()[0]

        if isOtWork(bizID):
            biz_hours_map[bizID] = ''
        else:
            biz_hours_map[bizID] = '0.1'
    response = sendToServer('mywork/timesheet/saveTimeSheetAction.do',  biz_hours_map)

#填写普通报工
def fillWorkHour(inputIDs, holidays=set()):
    pttList = []
    workHoursMap = {}
    workHoursMap.update(base_form)

    cnt = 0
    hourIndex = 0
    otHourIndex = 0
    hours = list(decomposition(WORK_HOURS))
    otHours = list(decomposition(OT_WORK_HOURS))

    if holidays:
        inputIDs, workHoursMap = filter_holiday(holidays, inputIDs, workHoursMap)

    for inputID in sorted(inputIDs, key=getLastChar):
        if cnt < projectCnt:
            if not isOtWork(inputID):
                pattern = re.compile('(' + projectID + '.*)_')
                pttList.append(pattern.search(inputID).groups()[0])
                cnt += 1

        if cmp(inputID[-1], getWeekOfDay(u'星期六')) == 0 or cmp(inputID[-1], getWeekOfDay(u'星期天')) == 0:
            if isOtWork(inputID):
                (otHours, workHour, otHourIndex) = getHour(otHours, otHourIndex, NORMAL_TYPE)
            else:
                workHour = ''
        else:
            if isOtWork(inputID):
                (otHours, workHour, otHourIndex) = getHour(otHours, otHourIndex, OT_TYPE)
            else:
                (hours, workHour, hourIndex) = getHour(hours, hourIndex, NORMAL_TYPE)

        workHoursMap[inputID] = workHour

    for ptt in pttList:
        workHoursMap['ptt'] = ptt
        response = sendToServer('mywork/timesheet/saveTimeSheetAction.do',  workHoursMap)
    print response.content

def openPageByIE():
    url = host + 'mywork/timesheet/saveTimeSheetAction.do' + '?startDate=' + str(startDate) + '&endDate=' + str(endDate)
    PAMIE.openWorkHourPage(url)

if __name__ == '__main__':
    #getLastWeekDate()
    #TEST
    startDate = '2015-04-06'
    endDate = '2015-04-12'
    base_form = {'startDate': startDate, 'endDate': endDate}
    print u"按回车键开始自动填写上周工时"
    raw_input()

    login()
    parsePage()

    #打开报工页面人工核实提交
    openPageByIE()
