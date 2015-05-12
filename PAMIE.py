# encoding=cp936
import cPAMIE
from time import sleep


def open_work_hour_page(url, login_name, pass_word):
    ie = cPAMIE.PAMIE()
    ie.navigate('http://vp.csii.com.cn/project/login.jsp')
    ie.textBoxSet('loginName', login_name)
    ie.textBoxSet('password', pass_word)
    ie.imageClick(0)
    # sleep(1)
    # ie.navigate(url)
