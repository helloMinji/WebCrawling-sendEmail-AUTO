from selenium import webdriver
from bs4 import BeautifulSoup
import time
import win32com.client


# chrome창을 unvisible하게
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu") #이게 안될 경우 options.add_argument("--disable-gpu")

# 사이트 열기
driver = webdriver.Chrome('chromedriver가 설치된 경로', chrome_options=options)
driver.get('웹사이트 주소') #인트라넷을 이용했었는데, 잘 실행됨
time.sleep(10) #화면이 뜨기 위한 대기시간

# 로그인
driver.find_element_by_name('USER').send_keys('아이디') #아이디
driver.find_element_by_name('PASSWORD').send_keys('비밀번호') #패스워드
driver.find_element_by_xpath('//*[@id="IMAGE1"]').click() #로그인버튼 클릭
time.sleep(20) #화면이 뜨기 위한 대기시간

driver.find_element_by_xpath('/html/body/main/div/div[3]/div[2]/div/div/div[1]/div[1]/a').click()
time.sleep(10) #원하는 부분 선택하고 대기(layout이 바뀔경우 다시 설정해야 함)
driver.find_element_by_xpath('/html/body/main/div/div[3]/div[3]/div/div/div/div[1]/div[1]/ul/li[5]/a').click()
time.sleep(10) #원하는 부분 선택하고 대기(layout이 바뀔경우 다시 설정해야 함)



html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
table = soup.find('table', class_='table download-list')

templist = table.find_all('span', class_="label label-success label-download-file") #update의 아이콘 class로 추출

# New와 Updated가 섞여있어서 Updated만 추출
updatelist = []

for item in templist:
    if item.get_text()=="Updated":
        updatelist.append(item)

ppt_list = []  # ppt파일의 주소
name_list = []  # 파일의 이름
date_list = []  # 파일 업데이트 날짜
other_list = []  # ppt가 아닌 것
isexist = False
inum = 0

# ppt가 있을 경우 그 주소를, 아니면 다른 파일이라도 주소를 저장하게 한다.

for updateone in updatelist:
    title = updateone.parent

    for fileaddr in title.next_sibling.next_sibling.find_all('a'):
        # print(fileaddr)
        if 'ppt' in fileaddr.get('href'):
            name_list.append(title.get_text().strip().replace("Updated", ""))

            ppt_list.append(fileaddr.get('href'))

            date = title.next_sibling.next_sibling.next_sibling.next_sibling.get_text().split()
            date = "".join(date[1:])
            date_list.append(date)

            isexist = True

        else:
            if not isexist:
                ppt_list.append(inum)
                other_list.append(fileaddr.get('href'))
                inum += 1



# 미리 저장해둔 확인해야 할 파일 목록
file = open('filelist 파일의 경로','r',encoding='ANSI')

temp = file.read()

alist = temp.split('\n')
c_name = []
c_year = []

# 이름과 갱신날짜 분리
for item in alist:
    if alist.index(item)%2==0:
        c_name.append(item)
    else:
        c_year.append(item)

file.close()



# 내가 갖고 있는 파일 중에 update가 생긴게 있는지 확인(이름은 같고 날짜는 다르다)

needsend_index = []

for me in c_name:
    for you in name_list:
        if me in you:
            #print(me)
            if c_year[c_name.index(me)]!=date_list[name_list.index(you)]:
                needsend_index.append(name_list.index(you))
                c_year[c_name.index(me)] = date_list[name_list.index(you)]

# 갱신날짜를 새로운 걸로 교체해서 파일에 써준다.
file = open('filelist 파일의 경로','w')

for i in range(0, len(c_name)):
    file.write(c_name[i] + '\n')
    if i == len(c_name) - 1:
        file.write(c_year[i])
    else:
        file.write(c_year[i] + '\n')

file.close()



# 메일 보내기

# 보내는 사람은 현재 pc의 outlook에 로그인된 사람

if len(needsend_index)!=0:
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = '받는 사람 메일 주소(여러명일 경우 ;로 구분)'
    #mail.CC = '참조할 사람 메일 주소'
    mail.Subject = '제목' #제목

    body = "다음의 파일이 업데이트 되었습니다. \n\n\n"
    for i in needsend_index:
        body = body + name_list[i] + " / " + date_list[i] + "\n\n"
    body = body+'\n이상입니다. 감사합니다.'

    mail.Body = body

    for i in needsend_index:
        mail.Attachments.Add('주소'+ppt_list[i])
        # 첨부파일이 있는 주소. 파일이 모두 같은 웹페이지에 있기 때문에 앞부분 주소가 동일하여 '주소'로 처리해주었다.
        # 첨부파일을 open해야 save를 위한 새로운 웹페이지가 열린다.

    mail.Send()

driver.quit()
