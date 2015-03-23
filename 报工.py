#encoding=cp936
import ConfigParser
import datetime
import random
import re
import requests
from BeautifulSoup import BeautifulSoup
import PAMIE

s = requests.session()
bizTravelID = "1321_16645" #������Ŀ��ţ�����ҳ��д����
WORK_HOURS = 7.9
OT_WORK_HOURS = 4
NORMAL_TYPE = 0
OT_TYPE = 1

#��ȡ�����ļ�
config = ConfigParser.ConfigParser()
config.read("config.ini")
userName = config.get("global", "userName")
passWord = config.get("global", "passWord")
projectID = config.get("global", "projectID")
projectNames = config.get("global", "projectName").split(',')
host = config.get("global", "host")
projectCnt = len(projectNames)
inputIDs = []

base_form = {}

def getLastWeekDate():
	global startDate
	global endDate
	global base_form

	today = datetime.date.today()
	startDate = today + datetime.timedelta(-7 - today.weekday()) #��ȡ����һ������
	endDate = startDate + datetime.timedelta(6)
	base_form = {'startDate': startDate, 'endDate': endDate}
	print "������ʼ����: " + str(startDate) +  " ������������: " + str(endDate)
	

def sendToServer(path, form_data, desc="undefined", params=""):
	response = s.post(host + path, data = form_data, params = params)
	if cmp(desc, "undefined") == 0:
		return response

	if cmp(response.content, "success") == 0:
		print desc + "�ɹ�"
	elif response.content == "PasswordOutOfDate": #��½���볬ʱ
		form_data = {'loginName': userName,
			'oldPassword': passWord,
			'password': passWord,
			'password1': passWord,
			'action':'resetPassword',
		}
		sendToServer('forgetPasswordAction.do', form_data, '��������');
	else:
		print desc + "ʧ��"
		print response.content.decode('utf-8')
		exit()

#��¼
def login():
	login_data = {'loginName': userName,
		'password': passWord,
		'action':'checkOnline',
	}
	sendToServer('loginAction.do', login_data, "��½")

#���ӹ�����
def addProject(projectName):
	form_data={'action':'addmyworkdes',
		'taskName': projectName,
		'projectID':projectID,}
	form_data.update(base_form)
	sendToServer('mywork/timesheet/addMyWorkDesAction.do', form_data, "�����Ŀ" + projectName)

#��ӳ���
def addBizTravel():
	xml_data='''<?xml version="1.0" encoding="gb2312"?>
	<TreeTable xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="TreeTable.xsd" high="20" width="125" color="black" detailable="true" standardmenu="true">
	<TableData>
	<RowData><name>�����������0.1��</name><NodeData nodeid="T_1321_16645_4408" nodetypeid="1321_16645_4408" image="29" />
	<CellData><Cell colid="ck" type="boolean">True</Cell><Cell colid="t1" /><Cell colid="t2" /></CellData></RowData>
	</TableData></TreeTable>'''
	params = {'startDate': startDate, 'endDate': endDate}
	sendToServer('mywork/timesheet/addMyTaskAction.do', xml_data, "��ӳ���", params)

#�鿴����ҳ��
def parsePage():
	response = sendToServer('mywork/timesheet/initTimeSheetAction.do', base_form)
	#�鿴����һ���Ƿ������:
	if bizTravelID not in response.content:
		addBizTravel()
	#�鿴��Ŀһ���Ƿ����
	for projectName in projectNames:
		if projectName not in response.content:
			addProject(projectName)

	#����֮ǰ��ӵ���Ŀid��inputtext_timesheet��ȡ������id��
	response = sendToServer('mywork/timesheet/initTimeSheetAction.do', base_form)
	soup = BeautifulSoup(response.content)
	inputIDs = [x.get('name') for x in soup.findAll(name="input", attrs={"name": re.compile(projectID), "class" : re.compile("inputtext")})]
	fillWorkHour(inputIDs)

	bizTravelInputIDs = [x.get('name') for x in soup.findAll(name="input", attrs={"name": re.compile(bizTravelID), "class" : re.compile("inputtext")})]
	fillTravelHour(bizTravelInputIDs)
	
#��д���ʱ
def fillTravelHour(bizTravelInputIDs):
	bizHoursMap = {}
	bizHoursMap.update(base_form)
	firstFlag = True
	for bizID in bizTravelInputIDs:
		if firstFlag:
			firstFlag = False
			pattern = re.compile('(' + bizTravelID + '.*)_')
			bizHoursMap['ptt'] = pattern.search(bizID).groups()[0]
	
		if isOtWork(bizID):
			bizHoursMap[bizID] = ''
		else:
			bizHoursMap[bizID] = '0.1'
	response = sendToServer('mywork/timesheet/saveTimeSheetAction.do',  bizHoursMap)
	
#������Ŀ��������乤ʱ
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

#��ȡ���һλ�ַ�
def getLastChar(s):
	return s[-1]

#��ȡ��ʱ
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

#�ж��Ƿ�Ӱ���
def isOtWork(content):
	if "gxot" in content:
		return True
	else:
		return False

#��д��ͨ����
def fillWorkHour(inputIDs):
	pttList = []
	workHoursMap = {}
	workHoursMap.update(base_form)

	cnt = 0
	hourIndex = 0
	otHourIndex = 0
	hours = list(decomposition(WORK_HOURS))
	otHours = list(decomposition(OT_WORK_HOURS))

	for inputID  in sorted(inputIDs, key=getLastChar):
		if cnt < projectCnt:
			if isOtWork(inputID) == False:
				pattern = re.compile('(' + projectID + '.*)_')
				pttList.append(pattern.search(inputID).groups()[0])
				cnt += 1
	
		if cmp(inputID[-1], '5') == 0 or cmp(inputID[-1], '6') == 0: #��ĩ�����
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
	#print response.content

def openPageByIE():
	url = host + 'mywork/timesheet/saveTimeSheetAction.do' + '?startDate=' + str(startDate) + '&endDate=' + str(endDate)
	PAMIE.openWorkHourPage(url)

if __name__ == '__main__':
	getLastWeekDate()
	#����
	#startDate = '2015-3-9'
	#endDate = '2015-3-15'
	print "���س�����ʼ�Զ���д���ܹ�ʱ"
	raw_input()
	
	login()
	parsePage()

	#�򿪱���ҳ���˹���ʵ�ύ
	openPageByIE()
