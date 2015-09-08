#!python2.7.exe
# encoding=cp936
import ConfigParser
import PAMIE
import datetime
import logging
import random
import re
import requests
import sys
import types
from BeautifulSoup import BeautifulSoup

s = requests.session()
bizTravelID = "1321_16645"  # 出差项目编号，报工页面写死的
WORK_HOURS = 8 #默认是8，如果出差后面会设置成7.9
OT_WORK_HOURS = 4
NORMAL_TYPE = 0
OT_TYPE = 1

# 读取配置文件
config = ConfigParser.ConfigParser()
config.read("config.ini")
userName = ""
passWord = ""
userNames = config.get("global", "userName").split(',')
passWords = config.get("global", "passWord").split(',')
projectID = config.get("global", "projectID")
project_names = config.get("global", "projectName").split(',')
projectCnt = len(project_names)
host = config.get("global", "host")
holidays = config.get("global", "holidays").split(',')
workdays = config.get("global", "workdays").split(',')
isTravel = config.get('global', 'isTravel')

# 全局变量
input_ids = []
base_form = {}
submit_forms = []
logger = ""


def validate(date_text):
    try:
        return datetime.datetime.strptime(date_text, '%Y%m%d')
    except ValueError:
        raise ValueError("Incorrect data format, should be YYYYMMDD")


def get_week_range(req_date, offset=0): # offset是指和本周差几周
    global startDate
    global endDate
    global base_form

    if type(req_date) is types.StringType:
        req_date = validate(req_date)
    startDate = req_date - datetime.timedelta(req_date.weekday() + offset * 7)
    endDate = startDate + datetime.timedelta(6)
    base_form = {'startDate': startDate, 'endDate': endDate}
    print u"报工开始日期: " + str(startDate) + u" 报工结束日期: " + str(endDate)


def get_last_week_date():
    today = datetime.date.today()
    get_week_range(today, 1)


def send_to_server(path, form_data, desc="undefined", params=None):
    response = s.post(host + path, data=form_data, params=params)
    if cmp(desc, "undefined") == 0:
        return response

    if cmp(response.content, "success") == 0 or response.status_code == 302 or response.status_code == 200:
        print response.content.decode('gbk')
        print desc + u"成功"
    elif response.content == "PasswordOutOfDate":  # 登陆密码超时
        reset_passwd_form_data = {
            'loginName': userName,
            'oldPassword': passWord,
            'password': passWord,
            'password1': passWord,
            'action': 'resetPassword',
        }
        send_to_server('forgetPasswordAction.do', reset_passwd_form_data, u'重置密码')
    else:
        print desc + u"失败"
        print response.content.decode('gbk')
        print response.status_code, response.url
        exit()


# 登录
def login():
    login_data = {
        'loginName': userName,
        'password': passWord,
        'action': 'checkOnline',
    }
    send_to_server('loginAction.do', login_data, u"登陆")


# 增加工作项
def add_project(project_name):
    form_data = {
        'action': 'addmyworkdes',
        'taskName': project_name,
        'projectID': projectID}
    form_data.update(base_form)
    send_to_server('mywork/timesheet/addMyWorkDesAction.do', form_data, u"添加项目" + project_name.decode('GBK'))


# 添加出差
def add_biz_travel():
    xml_data = (
        '<?xml version="1.0" encoding="gb2312"?>'
        '<TreeTable xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xsi:noNamespaceSchemaLocation="TreeTable.xsd" high="20" width="125" '
        'color="black" detailable="true" standardmenu="true">'
        '<TableData>'
        '<RowData><name>出差（正常栏填0.1）</name><NodeData nodeid="T_1321_16645_4408" '
        'nodetypeid="1321_16645_4408" image="29" />'
        '<CellData><Cell colid="ck" type="boolean">True</Cell><Cell colid="t1" />'
        '<Cell colid="t2" /></CellData></RowData>'
        '</TableData></TreeTable>')
    print xml_data
    params = {'startDate': startDate, 'endDate': endDate}
    send_to_server('mywork/timesheet/addMyTaskAction.do', xml_data, u"添加出差", params)


def get_week_of_day(desc):
    if desc is None:
        return None
    week_of_day = {
        u'星期一': '0',
        u'星期二': '1',
        u'星期三': '2',
        u'星期四': '3',
        u'星期五': '4',
        u'星期六': '5',
        u'星期天': '6',
    }
    return week_of_day.get(desc[:3])


def get_holiday_week_of_day(page):
    result = set()
    for holiday in holidays:
        result.add(get_week_of_day(page.find(name="td", text=re.compile(holiday))))
    return result

def get_workday_week_of_day(page):
    result = set()
    for workday in workdays:
        result.add(get_week_of_day(page.find(name="td", text=re.compile(workday))))
    return result

# 查看报工页面
def parse_page():
    global input_ids
    global WORK_HOURS
    global isTravel
    global logger

    response = send_to_server('mywork/timesheet/initTimeSheetAction.do', base_form)
    # 查看出差一栏是否已添加:
    if isTravel == "True" and bizTravelID not in response.content:
        add_biz_travel()
    # 查看项目一栏是否添加
    for project_name in project_names:
        if project_name not in response.content:
            add_project(project_name)

    response = send_to_server('mywork/timesheet/initTimeSheetAction.do', base_form)
    soup = BeautifulSoup(response.content, fromEncoding="GBK")
    #logger.info(response.content)

    # 查看是否配置文件如果有节假日日期，不填工时:
    holidays = get_holiday_week_of_day(soup)
    workdays = get_workday_week_of_day(soup)

    # 根据之前添加的项目id和inputtext_timesheet获取输入框的id：
    input_ids = [x.get('name') for x in soup.findAll(
        name="input",
        attrs={
            "name": re.compile(projectID),
            "class": re.compile("inputtext")})]

    if isTravel == "True":
        WORK_HOURS = 7.9
        biz_travel_input_ids = [x.get('name') for x in soup.findAll(
            name="input",
            attrs={
                "name": re.compile(bizTravelID),
                "class": re.compile("inputtext")})]
        fill_travel_hour(biz_travel_input_ids, holidays)

    fill_work_hour(input_ids, holidays, workdays)


# 设置节假日的列不填工时
def filter_holiday(holidays_arg, input_ids_arg, request_form):
    result = []
    for input_id in input_ids_arg:
        if input_id[-1] in holidays_arg:
            request_form[input_id] = ''
        else:
            result.append(input_id)
            
    return result, request_form


# 根据项目数随机分配工时
def decomposition(remain_hour):
    cnt = projectCnt
    while cnt > 0:
        if cnt == 1:
            yield round(remain_hour, 1)
        else:
            n = round(random.uniform(0, remain_hour), 1)
            yield n
            remain_hour -= n
        cnt -= 1


# 获取最后一位字符
def get_last_char(string):
    return string[-1]


# 获取工时
def get_hour(hours, index, work_type):
    origin_index = index
    if origin_index >= projectCnt:
        origin_index = 0
        index = 0
        if work_type == NORMAL_TYPE:
            hours = list(decomposition(WORK_HOURS))
        else:
            hours = list(decomposition(OT_WORK_HOURS))
    index += 1

    try:
        hour = hours[origin_index]
    except IndexError:
        hour = ''
    return hours, hour, index


# 判断是否加班列
def is_ot_work(content):
    if "gxot" in content:
        return True
    else:
        return False


# 填写出差工时
def fill_travel_hour(biz_travel_input_ids, holidays_arg=set()):
    global submit_forms
    biz_hours_map = {}
    biz_hours_map.update(base_form)
    first_flag = True

    #节假日也要填工时
    #if holidays:
        #biz_travel_input_ids, biz_hours_map = filter_holiday(holidays_arg, biz_travel_input_ids, biz_hours_map)

    for bizID in biz_travel_input_ids:
        if first_flag:
            first_flag = False
            pattern = re.compile('(' + bizTravelID + '.*)_')
            biz_hours_map['ptt'] = pattern.search(bizID).groups()[0]

        if is_ot_work(bizID):
            biz_hours_map[bizID] = ''
        else:
            biz_hours_map[bizID] = '0.1'
    send_to_server('mywork/timesheet/saveTimeSheetAction.do',  biz_hours_map)
    submit_forms.append(biz_hours_map)


# 填写普通报工
def fill_work_hour(input_ids_arg, holidays_set=set(), workdays_set=set()):
    global submit_forms
    ptt_list = []
    work_hours_map = {}
    work_hours_map.update(base_form)

    cnt = 0
    hour_index = 0
    othour_index = 0
    hours = list(decomposition(WORK_HOURS))
    ot_hours = list(decomposition(OT_WORK_HOURS))

    #节假日也要填工时
    #if holidays_set:
        #input_ids_arg, work_hours_map = filter_holiday(holidays_set, input_ids_arg, work_hours_map)

    for inputID in sorted(input_ids_arg, key=get_last_char):
        if cnt < projectCnt:
            if not is_ot_work(inputID):
                pattern = re.compile('(' + projectID + '.*)_')
                ptt_list.append(pattern.search(inputID).groups()[0])
                cnt += 1

        if inputID in holidays_set or \
            (inputID not in workdays_set and \
            cmp(inputID[-1], get_week_of_day(u'星期六')) == 0 or cmp(inputID[-1], get_week_of_day(u'星期天'))) == 0:

            if is_ot_work(inputID):
                (ot_hours, work_hour, othour_index) = get_hour(ot_hours, othour_index, NORMAL_TYPE)
            else:
                work_hour = ''
        else:
            if is_ot_work(inputID):
                (ot_hours, work_hour, othour_index) = get_hour(ot_hours, othour_index, OT_TYPE)
            else:
                (hours, work_hour, hour_index) = get_hour(hours, hour_index, NORMAL_TYPE)

        work_hours_map[inputID] = work_hour

    for ptt in ptt_list:
        work_hours_map['ptt'] = ptt
        send_to_server('mywork/timesheet/saveTimeSheetAction.do',  work_hours_map)
        submit_forms.append(work_hours_map)


def open_page_by_ie():
    url = host + ('mywork/timesheet/timeSheetMenuCardAction.do'
                  '?timeSheetFlag=person&sub=user-defined&startDate=' +
                  str(startDate) + '&endDate=' + str(endDate))
    PAMIE.open_work_hour_page(url, userName, passWord)


def query_yes_no(question, default="yes"):
    valid = {"yes": True, "y": True, "no": False, "n": False}
    if default is None:
        prompt = " [y/n] "
    elif default == "yes": #用大写来表示默认值
        prompt = " [Y/n] "
    elif default == "no":
        prompt = " [y/N] "
    else:
        raise ValueError("invalid default answer: '%s'" % default)

    while True:
        sys.stdout.write(question + prompt)
        choice = raw_input().lower()
        if default is not None and choice == '':
            return valid[default]
        elif choice in valid:
            return valid[choice]
        else:
            sys.stdout.write("Please respond with 'yes' or 'no' "
                             "(or 'y' or 'n').\n")

def setLog():
    global logger

    # 创建一个logger 
    logger = logging.getLogger('mylogger') 
    logger.setLevel(logging.DEBUG) 
       
    # 创建一个handler，用于写入日志文件 
    fh = logging.FileHandler('test.log') 
    fh.setLevel(logging.DEBUG) 
       
    # 再创建一个handler，用于输出到控制台 
    ch = logging.StreamHandler() 
    ch.setLevel(logging.DEBUG) 
       
    # 定义handler的输出格式 
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s') 
    fh.setFormatter(formatter) 
    ch.setFormatter(formatter) 
       
    # 给logger添加handler 
    logger.addHandler(fh) 
    logger.addHandler(ch)


if __name__ == '__main__':
    setLog()
    get_last_week_date()
    if not query_yes_no(u"开始上周报工吗"):
        print u"请输入要报工的日期yyyymmdd:"
        input_date = raw_input()
        get_week_range(input_date)

    
    for i in xrange(len(userNames)):
        userName = userNames[i]
        passWord = passWords[i]
        if i > 0:
            query_yes_no(u"开始对下一个用户报工:" + userName)

        print "UserNames=", userNames
        print "UserName=", userName
        login()
        parse_page()

        # 打开报工页面人工核实提交
        open_page_by_ie()

