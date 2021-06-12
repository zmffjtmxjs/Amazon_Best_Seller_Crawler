from bs4 import BeautifulSoup
from selenium import webdriver

import time
import sys
import urllib
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import requests

query_txt = '아마존 닷컴'
query_url = 'https://www.amazon.com/bestsellers'

print("아마존 닷컴의 분야별 Best Seller 상품 정보 크롤러")

#크롬 드라이버를 활성화하지 않고 없이 html 받아오기
response = requests.get(query_url)
soup = BeautifulSoup(response.content, 'html.parser')

reple_result = soup.select('#zg_browseRoot > ul')
slist = reple_result[0].find_all('li')
sec_names = []
inputMsg = ''

for i in slist:
    sec_names.append(i.get_text())

for i in range(0, len(sec_names)):
    if (i % 3 == 0):
        inputMsg += '\n'
    inputMsg += '%2s%-32s' % (str(i + 1), ('.' + sec_names[i]))

inputMsg += '\n' * 2 + '1.위 분야 중에서 자료를 수집할 분야의 번호를 선택하세요 : '

while True:
    sec = int(input(inputMsg))

    if ((sec > 0) & (sec <= len(sec_names))):
        break
    else:
        print("잘못된 번호를 입력하였습니다.")

while True:
    cnt = int(input("2. 해당 분야에서 크롤링 할 건수는 몇건입니까?(1-100 건 사이 입력) : " ))
    if cnt < 101:
        break
    else:
        print("검색 건수는 1건 - 최대 100 건까지만 가능합니다.")

f_dir = input("3.파일을 저장할 폴더명만 쓰세요(예 : c:\\temp\\) : ")
print("\n")



sec_name = sec_names[sec - 1]

now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

resultName = s + '-' + query_txt + '-' + sec_name

f_dir += resultName

os.makedirs(f_dir)
os.chdir(f_dir)
os.makedirs(f_dir + '/images')

fileName = f_dir + '/' + resultName
imageName = f_dir + '/images/'


path = "E:/coding/3years/chrome driver/chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get(query_url)
time.sleep(1)

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

reple_result = soup.select('#zg_browseRoot > ul')
slist = reple_result[0].find_all('a')

driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[""" + str(sec) + """]/a""").click()

time.sleep(1)

def scroll_down(driver):
    driver.execute_script("window.scrollBy(0, 9300);")
    time.sleep(1)

scroll_down(driver)

bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)

ranking2 = []
title3 = []
price2 = []
score2 = []
sat_count2 = []
store2 = []
imgs = []

count = 0

while True:
    
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    
    reple_result = soup.select('#zg-center-div > #zg-ordered-list')
    slist = reple_result[0].find_all('li')
    
    for li in slist:
        f = open(fileName + '.txt', 'a', encoding = 'UTF-8')
        f.write("-" * 40 + "\n")

        #판매순위
        print("-"*70)
        try:
            ranking = li.find('span', class_='zg-badge-text').get_text().replace("#", "")
        except:
            ranking = ''
            print(ranking.replace("#", ""))
        else:
            print("1.판매순위 :", ranking)
    
            f.write('1.판매순위 :' + ranking + '\n')
        
        #제품 이미지
        try:
            src  = li.find('span', class_='zg-text-center-align').find('img')['src']
        except:
            src = ''
            
        #제품 설명
        try:
            title1 = li.find('div', class_='p13n-sc-truncated').get_text().replace('\n', '')
        except AttributeError:
            title = ''
            print(title1.replace('\n', ''))
            f.write('2.제품소개 : ' + title1 + '\n')
        else:
            title2 = title1.translate(bmp_map).replace('\n', '')
            print('2.제품소개 :', title2.replace('\n', ''))
    
        count += 1
    
        f.write('2.제품소개 : ' + title2 + '\n')
    
        #가격
        try:
            price = li.find('span', 'p13n-sc-price').get_text().replace('\n', '')
        except AttributeError:
            price = ''
    
        print('3.가격 :', price.replace('\n', ''))
        f.write('3.가격 : ' + price + '\n')
    
        #상품평 수
        try:
            sat_count = li.find('a', 'a-size-small a-link-normal').get_text().replace(',', '')
        except (IndexError, AttributeError):
            sat_count = '0'
            print('4.상품평 수 :', sat_count)
            f.write('4.상품평 수 : ' + sat_count + '\n')
        else:
            print('4.상품평 수 :', sat_count)
            f.write('4.상품평 수 : ' + sat_count + '\n')
    
        #상품 별점 구하기
        try:
            score = li.find('span', 'a-icon-alt').get_text()
        except AttributeError:
            score = ''
    
        print('5.평점 :', score)
        f.write('5.평점 : ' + score + '\n')
    
        print('-' * 70)
    
        f.close()
    
        time.sleep(0.3)
    
        ranking2.append(ranking)
        title3.append(title2.replace('\n', ''))
        price2.append(price.replace('\n', ''))
    
        try:
          sat_count2.append(sat_count)
        except IndexError:
          sat_count2.append(0)
    
        score2.append(score)
        
        #이미지 저장
        if(src != ''):
            try:
                urllib.request.urlretrieve(src, imageName + str(count) + '.jpg')        #추출한 src를 통해 이미지 다운로드
                imgs.append(imageName + str(count) + '.jpg')                            #다운로드된 이미지 경로 배열 입력
            except:
                #다운로드 실패 시 공란 처리
                imgs.append('')
        else:
            #src를 가져오지 못했을 경우 공란 처리
            imgs.append('')
        
        if count == cnt :
          break
    
    #지정한 검색 건수 도달 여부 확인
    if count == cnt :
        break
    else:
        #1 페이지 추출 후 2페이지로 넘어감
        print("\n")
        print("2페이지로 이동한 후 데이터 크롤링 진행")
        print("\n")
        driver.find_element_by_xpath("""//*[@id="zg-center-div"]/div[2]/div/ul/li[3]/a""").click()

driver.quit()



#검색 결과를 다양한 형태로 저장하기

amazon_best_seller = pd.DataFrame()
amazon_best_seller['판매순위'] = ranking2
amazon_best_seller['제품소개'] = pd.Series(title3)
amazon_best_seller['판매가격'] = pd.Series(price2)
amazon_best_seller['상품평 갯수'] = pd.Series(sat_count2)
amazon_best_seller['상품평점'] = pd.Series(score2)

#csv 형태로 저장
amazon_best_seller.to_csv(fileName + '.csv', encoding = "utf-8-sig", index = True)

#엑셀 형태로 저장하기
amazon_best_seller.to_excel(fileName + '.xlsx', index = True)

#그림추가
wb = load_workbook(filename = fileName + '.xlsx', read_only = False, data_only = False)
ws = wb.active

for i in range(0, len(imgs)):
    if(imgs[i] != ''):                                                  #이미지 파일 누락 시 건너뜀
        img = Image(imgs[i])                                            #추가 할 이미지 파일 위치
        
        cellNum = i + 2                                                 #셀 크기 조절 대상을 이미지 저장 위치에 맞춤
        
        ws.row_dimensions[cellNum].height = img.height * 0.75 + 16      #이미지 크기에 맞게 높이 조절
        ws.column_dimensions['C'].width = 102                           #제목 최대 길이에 맞게 넓이 조절
        
        ws.add_image(img, 'C' + str(cellNum))                           #이미지를 엑셀에 추가

wb.save(fileName + '.xlsx')