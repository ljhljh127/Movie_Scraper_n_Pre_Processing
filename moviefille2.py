import time,os,shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert

##메인 웹페이지 연결
def Mainweb():
    driver = webdriver.Chrome(f"{os.getcwd()}/chromedriver.exe")
    driver.implicitly_wait(100)
    driver.get("https://www.kobis.or.kr/kobis/business/stat/offc/findFormerBoxOfficeList.do")
    ##추가 리스트 불러오기
    addlist = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div[4]/div[2]/div/a')
    addlist2 = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div[4]/div[3]/div/a')
    addlist3 = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div[4]/div[4]/div/a')
    addlist4 = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div[4]/div[5]/div/a')
    addlist.click()
    addlist2.click()
    addlist3.click()
    addlist4.click()
    return driver

##통계정보 들어가기
def Statistic():
    statistic=driver.find_element(By.XPATH,'/html/body/div[2]/div[1]/div[2]/ul/li[2]/a')
    statistic.click()
    driver.implicitly_wait(100)

# 엑셀파일 다운로드
def Exceldwn():
    exceldn=driver.find_element(By.XPATH,'/html/body/div[2]/div[2]/div/div[2]/div[1]/div[1]/div/a')
    exceldn.click()

##확인 누르기
def Confirm():
    da = Alert(driver)
    da.accept()
    time.sleep(10)
    driver.close()

##파일 renaming moviename 변수에 담긴 이름으로 바꿔줌
def Renaming(moviename):
    filepath='C:\\Users\\이정현\\Downloads'
    filename = max([filepath + '\\' + f for f in os.listdir(filepath)], key=os.path.getctime)
    shutil.move(filename, os.path.join(filepath, f"{moviename}.xls"))
    filename = filepath + "\\" + f"{moviename}.xls"
    print(f'{filename}')

def function(movieinform):
    ##영화이름 가져오기
    moviename = movieinform.text
    ##문자열의 특수문자 모두 제거 제거 안할시 renaming 과정중 오류 발생
    replace=''
    for character in moviename:
        if character.isalnum():
            replace=replace+character
    moviename=replace

    movieinform.click()
    Statistic()
    Exceldwn()
    Confirm()
    Renaming(moviename)

for i in range (1,501):
    driver=Mainweb()
    movieinform=driver.find_element(By.XPATH,f'/html/body/div/div[2]/div[2]/div[4]/div[6]/table/tbody/tr[{i}]/td[2]/a')
    function(movieinform)









