# encoding=cp936
import ConfigParser
import PAMIE
import datetime
import random
import re
import requests
import sys
import types
from BeautifulSoup import BeautifulSoup

s = requests.session()
bizTravelID = "1321_16645"  # ������Ŀ��ţ�����ҳ��д����
WORK_HOURS = 7.9
OT_WORK_HOURS = 4
NORMAL_TYPE = 0
OT_TYPE = 1

# ��ȡ�����ļ�
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

# ȫ�ֱ���
input_ids = []
base_form = {}
submit_forms = []


def validate(date_text):
    try:
        return datetime.datetime.strptime(date_text, '%Y%m%d')
    except ValueError:
        raise ValueError("Incorrect data format, should be YYYYMMDD")


def get_week_range(req_date, offset=0): # offset��ָ�ͱ��ܲ��
    global startDate
    global endDate
    global base_form

    if type(req_date) is types.StringType:
        req_date = validate(req_date)
    startDate = req_date - datetime.timedelta(req_date.weekday() + offset * 7)
    endDate = startDate + datetime.timedelta(6)
    base_form = {'startDate': startDate, 'endDate': endDate}
    print u"������ʼ����: " + str(startDate) + u" ������������: " + str(endDate)


def get_last_week_date():
    today = datetime.date.today()
    get_week_range(today, 1)


def send_to_server(path, form_data, desc="undefined", params=None):
    response = s.post(host + path, data=form_data, params=params)
    if cmp(desc, "undefined") == 0:
        return response

    if cmp(response.content, "success") == 0:
        print desc + u"�ɹ�"
    elif response.content == "PasswordOutOfDate":  # ��½���볬ʱ
        form_data = {
            'loginName': userName,
            'oldPassword': passWord,
            'password': passWord,
            'password1': passWord,
            'action': 'resetPassword',
        }
        send_to_server('forgetPasswordAction.do', form_data, u'��������')
    else:
        print desc + u"ʧ��"
        print response.content.decode('gbk')
        exit()


# ��¼
def login():
    login_data = {
        'loginName': userName,
        'password': passWord,
        'action': 'checkOnline',
    }
    send_to_server('loginAction.do', login_data, u"��½")


# ���ӹ�����
def add_project(project_name):
    form_data = {
        'action': 'addmyworkdes',
        'taskName': project_name,
        'projectID': projectID}
    form_data.update(base_form)
    send_to_server('mywork/timesheet/addMyWorkDesAction.do', form_data, u"�����Ŀ" + project_name.decode('GBK'))


# ��ӳ���
def add_biz_travel():
    xml_data = (
        '<?xml version="1.0" encoding="gb2312"?>'
        '<TreeTable xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xsi:noNamespaceSchemaLocation="TreeTable.xsd" high="20" width="125" '
        'color="black" detailable="true" standardmenu="true">'
        '<TableData>'
        '<RowData><name>�����������0.1��</name><NodeData nodeid="T_1321_16645_4408" '
        'nodetypeid="1321_16645_4408" image="29" />'
        '<CellData><Cell colid="ck" type="boolean">True</Cell><Cell colid="t1" />'
        '<Cell colid="t2" /></CellData></RowData>'
        '</TableData></TreeTable>')
    print xml_data
    params = {'startDate': startDate, 'endDate': endDate}
    send_to_server('mywork/timesheet/addMyTaskAction.do', xml_data, u"��ӳ���", params)


def get_week_of_day(desc):
    if desc is None:
        return None
    week_of_day = {
        u'����һ': '0',
        u'���ڶ�': '1',
        u'������': '2',
        u'������': '3',
        u'������': '4',
        u'������': '5',
        u'������': '6',
    }
    return week_of_day.get(desc[:3])


def get_holiday_week_of_day(page):
    result = set()
    for holiday in holidays:
        result.add(get_week_of_day(page.find(name="td", text=re.compile(holiday))))
    return result


# �鿴����ҳ��
def parse_page():
    global input_ids
    response = send_to_server('mywork/timesheet/initTimeSheetAction.do', base_form)
    # �鿴����һ���Ƿ������:
    if bizTravelID not in response.content:
        add_biz_travel()
    # �鿴��Ŀһ���Ƿ����
    for project_name in project_names:
        if project_name not in response.content:
            add_project(project_name)

    response = send_to_server('mywork/timesheet/initTimeSheetAction.do', base_form)
    soup = BeautifulSoup(response.content, fromEncoding="GBK")

    # �鿴�Ƿ������ļ�����нڼ������ڣ����ʱ:
    ignore_days = get_holiday_week_of_day(soup)

    # ����֮ǰ��ӵ���Ŀid��inputtext_timesheet��ȡ������id��
    input_ids = [x.get('name') for x in soup.findAll(
        name="input",
        attrs={
            "name": re.compile(projectID),
            "class": re.compile("inputtext")})]

    fill_work_hour(input_ids, ignore_days)

    biz_travel_input_ids = [x.get('name') for x in soup.findAll(
        name="input",
        attrs={
            "name": re.compile(bizTravelID),
            "class": re.compile("inputtext")})]
    fill_travel_hour(biz_travel_input_ids, ignore_days)


def filter_holiday(holidays_arg, input_ids_arg, request_form):
    result = []
    for input_id in input_ids_arg:
        if input_id[-1] in holidays_arg:
            request_form[input_id] = ''
        else:
            result.append(input_id)
            
    return result, request_form


# ������Ŀ��������乤ʱ
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


# ��ȡ���һλ�ַ�
def get_last_char(string):
    return string[-1]


# ��ȡ��ʱ
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


# �ж��Ƿ�Ӱ���
def is_ot_work(content):
    if "gxot" in content:
        return True
    else:
        return False


# ��д���ʱ
def fill_travel_hour(biz_travel_input_ids, holidays_arg=set()):
    global submit_forms
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

        if is_ot_work(bizID):
            biz_hours_map[bizID] = ''
        else:
            biz_hours_map[bizID] = '0.1'
    send_to_server('mywork/timesheet/saveTimeSheetAction.do',  biz_hours_map)
    submit_forms.append(biz_hours_map)


# ��д��ͨ����
def fill_work_hour(input_ids_arg, holidays_set=set()):
    global submit_forms
    ptt_list = []
    work_hours_map = {}
    work_hours_map.update(base_form)

    cnt = 0
    hour_index = 0
    othour_index = 0
    hours = list(decomposition(WORK_HOURS))
    ot_hours = list(decomposition(OT_WORK_HOURS))

    if holidays_set:
        input_ids_arg, work_hours_map = filter_holiday(holidays_set, input_ids_arg, work_hours_map)

    for inputID in sorted(input_ids_arg, key=get_last_char):
        if cnt < projectCnt:
            if not is_ot_work(inputID):
                pattern = re.compile('(' + projectID + '.*)_')
                ptt_list.append(pattern.search(inputID).groups()[0])
                cnt += 1

        if cmp(inputID[-1], get_week_of_day(u'������')) == 0 or cmp(inputID[-1], get_week_of_day(u'������')) == 0:
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
    elif default == "yes": #�ô�д����ʾĬ��ֵ
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




if __name__ == '__main__':
    get_last_week_date()
    if not query_yes_no(u"��ʼ���ܱ�����"):
        print u"������Ҫ����������yyyymmdd:"
        input_date = raw_input()
        get_week_range(input_date)

    
    for i in xrange(len(userNames)):
        if i > 0:
            query_yes_no(u"��ʼ����һ���û�����:" + userName)
        userName = userNames[i]
        passWord = passWords[i]

        print "UserNames=", userNames
        print "UserName=", userName
        login()
        parse_page()

        # �򿪱���ҳ���˹���ʵ�ύ
        open_page_by_ie()

