#encoding=cp936
import cPAMIE
from time import sleep

def openWorkHourPage(url):
	ie = cPAMIE.PAMIE()

	ie.navigate('http://vp.csii.com.cn/project/login.jsp')
	ie.textBoxSet('loginName', '052321')
	ie.textBoxSet('password', '052321')
	ie.imageClick(0)
	sleep(1)

	ie.navigate(url)

#	ie.frameName = "top"
#	ie.listBoxSelect("mywork", u"��Ҫ����")
#	
#	print "���뱨��ҳ������"
#	raw_input()
#	ie.frameName = "right"
#
#	iframe = ie.getFrame('functionPage')
#	print iframe.document

if __name__ == '__main__':
	openWorkHourPage()
