#encoding=cp936
import ConfigParser
import datetime
import re
import requests
from BeautifulSoup import BeautifulSoup

s = requests.session()
bizTravelID = "1321_16645" #出差项目编号，报工页面写死的

#测试临时用
#startDate = '2015-03-02'
#endDate = '2015-03-08'

#读取配置文件
config = ConfigParser.ConfigParser()
config.readfp(open("config.ini"), "rb")
userName = config.get("global", "userName")
passWord = config.get("global", "passWord")
projectID = config.get("global", "projectID")
projectName = config.get("global", "projectName")
host = config.get("global", "host")

def getLastWeekDate():
	global startDate
	global endDate

	today = datetime.date.today()
	startDate = today + datetime.timedelta(-7 - today.weekday()) #获取上周一的日期
	endDate = startDate + datetime.timedelta(6)
	print "报工开始日期: " + str(startDate) +  " 报工结束日期: " + str(endDate)
	

def sendToServer(path, form_data, desc="undefined", params=""):
	response = s.post(host + path, data = form_data, params = params)
	if cmp(desc, "undefined") == 0:
		return response

	if cmp(response.content, "success") == 0:
		print desc + "成功"
	else:
		print desc + "失败"
		print response.content
		exit()

#登录
def login():
	login_data = {'loginName': userName,
		'password': passWord,
		'action':'checkOnline',
	}
	sendToServer('loginAction.do', login_data, "登陆")

#增加工作项
def addProject():
	form_data={'action':'addmyworkdes',
		'startDate': startDate,
		'endDate': endDate,
		'taskName': projectName,
		'projectID':projectID,}
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
	#sendToServer('mywork/timesheet/addMyTaskAction.do?startDate=' + startDate + '&endDate=' + endDate, xml_data, "添加出差")
	sendToServer('mywork/timesheet/addMyTaskAction.do', xml_data, "添加出差", params)

#查看报工页面
def parsePage():
	form_data={
		'startDate': startDate, 
		'endDate': endDate,
	}
	response = sendToServer('mywork/timesheet/initTimeSheetAction.do', form_data)
	#查看出差一栏是否已添加:
	if bizTravelID not in response.content:
		addBizTravel()
	if projectName not in response.content:
		addProject()
	#根据之前添加的项目id和inputtext_timesheet获取输入框的id：
	response = sendToServer('mywork/timesheet/initTimeSheetAction.do', form_data)
	soup = BeautifulSoup(response.content)
	inputIDs = [x.get('name') for x in soup.findAll(name="input", attrs={"name": re.compile(projectID), "class" : re.compile("inputtext")})]
	bizTravelInputIDs = [x.get('name') for x in soup.findAll(name="input", attrs={"name": re.compile(bizTravelID), "class" : re.compile("inputtext")})]
	
	#填写出差工时
	bizHoursMap = {}
	bizHoursMap['startDate'] = startDate
	bizHoursMap['endDate'] = endDate
	firstFlag = True
	for bizID in bizTravelInputIDs:
		if firstFlag:
			firstFlag = False
			pattern = re.compile('(' + bizTravelID + '.*)_')
			bizHoursMap['ptt'] = pattern.search(bizID).groups()[0]
	
		if "gxot" in bizID:
			bizHoursMap[bizID] = ''
		else:
			bizHoursMap[bizID] = '0.1'
	response = sendToServer('mywork/timesheet/saveTimeSheetAction.do',  bizHoursMap)
	
	#填写普通报工
	workHoursMap = {}
	workHoursMap['startDate'] = startDate
	workHoursMap['endDate'] = endDate
	first_flag = True
	for inputID in inputIDs:
		if first_flag:
			first_flag = False
			pattern = re.compile('(' + projectID + '.*)_')
			workHoursMap['ptt'] = pattern.search(inputID).groups()[0]
	
		if cmp(inputID[-1], '5') == 0 or cmp(inputID[-1], '6') == 0: #周末的情况
			if "gxot" in inputID: #加班工时
				workHours = '7.9'
			else:
				workHours = ''
		else:
			if "gxot" in inputID:
				workHours = '4'
			else:
				workHours = '7.9'
			
		workHoursMap[inputID] = workHours
	
	response = sendToServer('mywork/timesheet/saveTimeSheetAction.do',  workHoursMap)
	print response.content

#除了保存，还要确认提交, 做一个菜单选择

if __name__ == '__main__':
	getLastWeekDate()
	print "按回车键开始自动填写上周工时"
	raw_input()
	
	login()
	parsePage()
