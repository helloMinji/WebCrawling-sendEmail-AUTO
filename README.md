# 웹과 내 데이터를 비교하기 + 아웃룩 메일 전송

이 프로그램은 두 부분으로 구성되어 있습니다.     
- 초기 페이지부터 내가 원하는 페이지까지 xpath를 이용하여 이동하고, 최종 웹사이트의 데이터와 내가 가진 txt 데이터의 내용을 비교하여 조건을 만족하는 것만 list에 저장한다.
- 저장된 list의 내용을 이용하여 outlook으로 메일을 전송한다.
           
<br>
<br>
                    


---------------------------------------------------------------

### > How to Use

#### Filelist

- txt 파일이어야 합니다.
- 내용은 name, date로 구성되어있고 엔터(Enter)로 구분합니다.
- name은 사이트 기준으로 입력해야 하고, date는 공백없이 사이트 기준으로 입력해야 합니다.                    
#### :warning: 마지막 줄이 엔터로 끝나지 않도록 해야 합니다. 불필요한 공백, 줄바꿈이 있을 경우 오류가 발생합니다.

<br>
<br>

#### Setup

1. Python 3.6과 chromedriver 설치
- python 3.6
> https://www.python.org/downloads/windows/                    
> 상위 버전일 경우 생기는 오류가 아직 해결되지 않아 3.6 사용을 권장

- chromedriver
> https://sites.google.com/a/chromium.org/chromedriver/downloads                     
> 위 링크에서 Latest Release 옆 파일을 선택해 OS 별로 알맞은 zip 파일을 압축 풀기

<br>

2. cmd에서
```
python get-pip.py
```

<br>

3. 코드가 있는 위치에서
```
pip install selenium bs4 pywin32
```

<br>

4. 코드에서 본인에 맞게 수정
- chromedrive 설치 경로 확인 : 속성창에서 확인한 경우 전부
- id, password 변경 : 사이트에 따라 'USER','PASSWORD'라는 단어가 아닐 수도 있다(html 확인)
```python3
# 사이트 열기
driver = webdriver.Chrome('설치경로/chromedriver', chrome_options=options)
driver.get('원하는 웹사이트')
time.sleep(10)  # 화면이 뜨기 위한 대기시간

# 로그인
driver.find_element_by_name('USER').send_keys('아이디')
driver.find_element_by_name('PASSWORD').send_keys('패스워드')
driver.find_element_by_xpath('//*[@id="IMAGE1"]').click() # 로그인버튼 클릭
time.sleep(20)  # 화면이 뜨기 위한 대기시간
```

- filelist 경로 확인 : 확장자까지 입력
```python3
file = open('설치경로', 'r', encoding='ANSI')

file = open('설치경로', 'w')
```

- 메일 수신자 변경 : 본인 outlook 주소 가능
```python3
# 보내는 사람은 현재 pc의 outlook에 로그인된 사람

if len(needsend_index)!=0:
  outlook = win32com.client.Dispatch('outlook.application')
  mail = outlook.CreateItem(0)
  mail.To = '메일 주소(다인일 경우 ;로 구분)'
  mail.CC = '참조 주소'
  mail.Subject = '메일 제목'
```
<br>

5. 윈도우 스케쥴러에 task 추가
![image](https://user-images.githubusercontent.com/41939828/52467306-6afe1200-2bc8-11e9-8244-92edc0674c31.PNG)
![image](https://user-images.githubusercontent.com/41939828/52467322-76e9d400-2bc8-11e9-82ce-f001ae650e53.PNG)
![image](https://user-images.githubusercontent.com/41939828/52467333-810bd280-2bc8-11e9-86f5-4d05f9fd0192.PNG)
![image](https://user-images.githubusercontent.com/41939828/52467341-8832e080-2bc8-11e9-8b9b-c1e04926858b.PNG)
![image](https://user-images.githubusercontent.com/41939828/52467348-9123b200-2bc8-11e9-829d-89197b139d21.PNG)
이걸 해야 전원을 연결하지 않아도 task가 작동된다.

<br>
<br>
<br>

---------------------------------------------------------------

### > Warning!

특정 사이트에서만 테스트했기 때문에 다른 사이트에 적용할 경우 전반적인 수정이 필요할 수 있습니다.

<br>
<br>
<br>

---------------------------------------------------------------


### > Reference

- PyCharm에 라이브러리 추가하는 법
> https://woongheelee.com/entry/PyCharm%EC%97%90%EC%84%9C-%EB%9D%BC%EC%9D%B4%EB%B8%8C%EB%9F%AC%EB%A6%AC-%EC%9E%84%ED%8F%AC%ED%8A%B8import%ED%95%98%EB%8A%94-%EB%B0%A9%EB%B2%95

- 사이트에 로그인하기
> https://blog.naver.com/PostView.nhn?blogId=popqser2&logNo=221229125022&parentCategoryNo=&categoryNo=23&viewDate=&isShowPopularPosts=true&from=search

- 태그의 트리 구조
> http://www.hanbit.co.kr/media/channel/view.html?cms_code=CMS2068924870

- a href 가져오기
> https://hudi.kr/python-bs4-%EC%82%AC%EC%9A%A9%ED%95%98%EC%97%AC-a%ED%83%9C%EA%B7%B8%EC%9D%98-href-%EC%86%8D%EC%84%B1-%EA%B0%92-%EB%AA%A8%EB%91%90-%EA%B0%80%EC%A0%B8%EC%98%A4%EA%B8%B0/

- outlook 이용하기
> https://iamaman.tistory.com/1638

- outlook에 첨부파일 추가하기
> https://stackoverrun.com/ko/q/1585555


