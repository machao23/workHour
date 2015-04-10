# encoding=cp936
import cPAMIE
from time import sleep

def open_work_hour_page(url):
    ie = cPAMIE.PAMIE()
    ie.navigate('http://vp.csii.com.cn/project/login.jsp')
    ie.textBoxSet('loginName', '052321')
    ie.textBoxSet('password', '052321')
    ie.imageClick(0)
    sleep(1)
    ie.navigate(url)

if __name__ == '__main__':
    open_work_hour_page()
