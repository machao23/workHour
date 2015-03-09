#encoding=cp936
import ConfigParser
import datetime
import re
import requests
from BeautifulSoup import BeautifulSoup

s = requests.session()
bizTravelID = "1321_16645" #������Ŀ��ţ�����ҳ��д����

#������ʱ��
#startDate = '2015-03-02'
#endDate = '2015-03-08'

#��ȡ�����ļ�
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
	startDate = today + datetime.timedelta(-7 - today.weekday()) #��ȡ����һ������
	endDate = startDate + datetime.timedelta(6)
	print "������ʼ����: " + str(startDate) +  " ������������: " + str(endDate)
	

def sendToServer(path, form_data, desc="undefined", params=""):
	response = s.post(host + path, data = form_data, params = params)
	if cmp(desc, "undefined") == 0:
		return response

	if cmp(response.content, "success") == 0:
		print desc + "�ɹ�"
	else:
		print desc + "ʧ��"
		print response.content
		exit()

#��¼
def login():
	login_data = {'loginName': userName,
		'password': passWord,
		'action':'checkOnline',
	}
	sendToServer('loginAction.do', login_data, "��½")

#���ӹ�����
def addProject():
	form_data={'action':'addmyworkdes',
		'startDate': startDate,
		'endDate': endDate,
		'taskName': projectName,
		'projectID':projectID,}
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
	#sendToServer('mywork/timesheet/addMyTaskAction.do?startDate=' + startDate + '&endDate=' + endDate, xml_data, "��ӳ���")
	sendToServer('mywork/timesheet/addMyTaskAction.do', xml_data, "��ӳ���", params)

#�鿴����ҳ��
def parsePage():
	form_data={
		'startDate': startDate, 
		'endDate': endDate,
	}
	response = sendToServer('mywork/timesheet/initTimeSheetAction.do', form_data)
	#�鿴����һ���Ƿ������:
	if bizTravelID not in response.content:
		addBizTravel()
	if projectName not in response.content:
		addProject()
	#����֮ǰ��ӵ���Ŀid��inputtext_timesheet��ȡ������id��
	response = sendToServer('mywork/timesheet/initTimeSheetAction.do', form_data)
	soup = BeautifulSoup(response.content)
	inputIDs = [x.get('name') for x in soup.findAll(name="input", attrs={"name": re.compile(projectID), "class" : re.compile("inputtext")})]
	bizTravelInputIDs = [x.get('name') for x in soup.findAll(name="input", attrs={"name": re.compile(bizTravelID), "class" : re.compile("inputtext")})]
	
	#��д���ʱ
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
	
	#��д��ͨ����
	workHoursMap = {}
	workHoursMap['startDate'] = startDate
	workHoursMap['endDate'] = endDate
	first_flag = True
	for inputID in inputIDs:
		if first_flag:
			first_flag = False
			pattern = re.compile('(' + projectID + '.*)_')
			workHoursMap['ptt'] = pattern.search(inputID).groups()[0]
	
		if cmp(inputID[-1], '5') == 0 or cmp(inputID[-1], '6') == 0: #��ĩ�����
			if "gxot" in inputID: #�Ӱ๤ʱ
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

#���˱��棬��Ҫȷ���ύ, ��һ���˵�ѡ��

if __name__ == '__main__':
	getLastWeekDate()
	print "���س�����ʼ�Զ���д���ܹ�ʱ"
	raw_input()
	
	login()
	parsePage()
