from selenium import webdriver
import xlrd
import xlwt
import time
import winreg
from xlutils.copy import copy
violation_types={'x':'吸粉','f':'否','y':'引流','s':'涉黄','sz':'刷钻','sc':'首次','qz':'欺诈','qt':'其它','w':'无微聊','':'','save':'save','q':'quit','c':'correct'}
print("审核规则：",end='')
print("['x':'吸粉','f':'否','y':'引流','s':'涉黄','sz':'刷钻','sc':'首次','qz':'欺诈','qt':'其它','w':'无微聊','c':'修改']")
real_address = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',)
file_address=winreg.QueryValueEx(real_address, "Desktop")[0]
file_address+='\\'
file=input('请输入审核文件名称（无需添加后缀.xlsx或者.xls）:')
filename=file_address+file+".xlsx"
while True:
	try:
		wb=xlrd.open_workbook(filename)
	except FileNotFoundError:
		file=input('桌面查找不到该文件，请重新输入审核文件名称（请添加后缀.xlsx或者.xls）:')
		filename=file_address+file
	else:
		break
rb=copy(wb)
sheet_number=input("请输入查询第几页：")
active=True
while active:
	try:
		sheet1 = wb.sheet_by_index(int(sheet_number)-1)#根据索引获取sheet内容
	except IndexError:
		sheet_number=input("超出查找范围，查无此页，请重新输入查询第几页：")
	else:
		active=False#try成功之后要做的事
sheet2 = rb.get_sheet(int(sheet_number)-1)
col_number=input("userid在第几列：")
cols = sheet1.col_values(int(col_number)-1)
row_number=input("从第几行开始：")
col_output=input("结果输出在第几列：")
driver=webdriver.Chrome("chromedriver.exe")
driver.maximize_window()
driver.get("http://oa.58.com.cn")
active=True
while active:
	password=input("是否已经登录58盾与oa账号密码（是/否）：")
	if password=='是':
		active=False
startdate=input("请输入起始查询日期(xxxx-xx-xx):")
enddate=input("请输入结束查询日期(xxxx-xx-xx):")
driver.get("http://union.vip.58.com/bsp/index")
cluster_text = driver.find_element_by_id("bt11")
cluster_text.click()
time.sleep(2)
cluster_text_0 = driver.find_element_by_id("bt11_0")
cluster_text_0.click()
time.sleep(2)
cluster_text_1 = driver.find_element_by_id("weiliaomgrchatcha")
cluster_text_1.click()
driver.get("http://weiliaomgrweb.union.vip.58.com/contentmgr/list")
driver.find_element_by_xpath("//select[@id='searchtype']/option[text()='userid查询']").click()
driver.find_element_by_id("startdate").clear()
driver.find_element_by_id("startdate").send_keys(startdate)
driver.find_element_by_id("enddate").clear()
driver.find_element_by_id("enddate").send_keys(enddate)
driver.find_element_by_xpath("//select[@id='pageSize']/option[text()='500']").click()
def inquire(userid):
	driver.find_element_by_id("keyword").clear()
	driver.find_element_by_id("keyword").send_keys(userid)
	driver.find_element_by_id("btnsearch").click()
def check(value):
	while True:
		try:
			sub=violation_types[value]
		except KeyError:
			value=input('查无此违规类型，请重新输入:')
		else:
			break
	return sub
def excel_output(sub,n):
	if sub=='否':
		sheet2.write(n,int(col_output)-1, '否')
	elif sub=='':
		sheet2.write(n,int(col_output)-1, '')
		sheet2.write(n,int(col_output), sub)
	elif sub=='无微聊':
		sheet2.write(n,int(col_output)-1, '')
		sheet2.write(n,int(col_output)+1, sub)
	elif sub=='quit':
		break
	else:
		sheet2.write(n,int(col_output)-1, '是')
		sheet2.write(n,int(col_output), sub)
print("输入save可以进行保存/输入q可以退出")
n=1
for i in range(int(row_number)-1,sheet1.nrows):
	inquire(str(int(cols[i])))
	key=input('输入违规类型(第'+str(n)+'条):')
	sub=check(key)
	if sub=='save':
		rb.save('临时保存.xls')
		key=input('保存完毕，请继续输入本条id的违规类型:')
		sub=check(key)
		excel_output(sub,i)
	if sub=='correct':
		number_correct=input("请输入修改第几条：")
		m=i-(n-int(number_correct))
		inquire(str(int(cols[m])))
		key=input('输入违规类型:')
		sub=check(key)
		excel_output(sub,m)
		inquire(str(int(cols[i])))
		key=input('修改完毕，请继续输入本条id的违规类型:')
		sub=check(key)
		excel_output(sub,i)
	else:
		excel_output(sub,i)
	n+=1
save_name=input('已审核完毕，请输入保存文件名称：')
save_name+='.xls'
while True:
		try:
			save_address=file_address+save_name
			rb.save(save_address)
		except FileNotFoundError:
			save_name=input('文件名称重复，请重新输入：')
			save_name+='.xls'
		else:
			break
driver.close()

