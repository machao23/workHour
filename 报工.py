#!python2.7.exe
# encoding=utf-8
import ConfigParser
import PAMIE
import datetime
import json
import requests
import sys
import types

s = requests.session()

# 读取配置文件
config = ConfigParser.ConfigParser()
config.read("config.ini")
userName = ""
passWord = ""
WORK_HOURS = config.get("global", "WorkTime")
userNames = config.get("global", "userName").split(',')
passWords = config.get("global", "passWord").split(',')
projectID = config.get("global", "projectID")
projectDesc = config.get("global", "projectDesc")
projectType = config.get("global", "projectType")
projectTypeDesc = config.get("global", "projectTypeDesc")
Context = config.get("global", "Context")
host = config.get("global", "ip_host")
holidays = config.get("global", "holidays").split(',')
workdays = config.get("global", "workdays").split(',')
isTravel = config.get('global', 'isTravel')


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


def send_to_server(url, data, desc="undefined", params=None, headers=None):
    global host
    response = s.post(host + url, data=data, params=params, headers=headers)
    if cmp(desc, "undefined") == 0:
        return response

    elif cmp(response.content, "success") == 0 or response.status_code == 302 or response.status_code == 200:
        return response
    else:
        print desc + u"失败"
        print response.content.decode('utf-8')
        print response.status_code, response.url
        exit()


# 登录
def login():
    login_data = {
        'loginid': userName,
        'password': passWord,
    }
    response = send_to_server('timesheet/login.json', login_data, u"登陆")
    if u"成功" not in response.text:
        print response.text


# 判断是不是工作日
def is_workday(date):
    day_of_week = date.weekday() + 1
    if day_of_week == 6 or day_of_week == 7:
        return False
    else:
        return True


# 组装请求报文
def pack_request(date):
    return {
        "events": [
            {
                "title": projectDesc,
                "start": date.strftime("%Y-%m-%d"),
                "end": None,
                "className": [
                    "b-l b-2x b-info"
                ],
                "ProId": {
                    "keyColumn": projectID,
                    "valueColumn": projectDesc
                },
                "TSTypeId": {
                    "keyColumn": "1",
                    "valueColumn": projectTypeDesc
                },
                "LeaveTypeId": "",
                "IsTravel": True,
                "WorkTime": (is_workday(date) and 8) or 0,
                "DelayTime": (is_workday(date) and 4) or 8,
                "Context": Context,
                "_id": 1
            }
        ]
    }


# 查看报工页面
def parse_page(start, end):
    global WORK_HOURS
    global isTravel

    s_date = datetime.datetime.strptime(start, "%Y%m%d")
    e_date = datetime.datetime.strptime(end, "%Y%m%d")
    work_date = s_date
    while work_date <= e_date:
        request = pack_request(work_date)
        headers = {'content-type': 'application/json'}
        response = send_to_server(url='timesheet/calendarSave.json',
                                  data=json.dumps(request), headers=headers)
        print work_date.strftime("%Y-%m-%d"), "报工结果:", response.content.decode('utf-8')

        work_date += datetime.timedelta(days=1)


# 设置节假日的列不填工时
def filter_holiday(holidays_arg, input_ids_arg, request_form):
    result = []
    for input_id in input_ids_arg:
        if input_id[-1] in holidays_arg:
            request_form[input_id] = ''
        else:
            result.append(input_id)

    return result, request_form


def open_page_by_ie():
    PAMIE.open_work_hour_page(host, userName, passWord)


def query_yes_no(question, default="yes"):
    valid = {"yes": True, "y": True, "no": False, "n": False}
    if default is None:
        prompt = " [y/n] "
    elif default == "yes":  # 用大写来表示默认值
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
    start_date = raw_input(u"请输入要报工的开始日期yyyymmdd:")
    end_date = raw_input(u"请输入要报工的结束日期yyyymmdd:")

    for i in xrange(len(userNames)):
        userName = userNames[i]
        passWord = passWords[i]
        if i > 0:
            query_yes_no(u"开始对下一个用户报工:" + userName)

        print "UserNames=", userNames
        print "UserName=", userName
        login()
        parse_page(start_date, end_date)

        # 打开报工页面人工核实提交
        #open_page_by_ie()

