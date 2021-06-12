from bs4 import BeautifulSoup
from selenium import webdriver

import time
import sys
import urllib
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

print("아마존 닷컴의 분야별 Best Seller 상품 정보 추출")

query_txt = '아마존 닷컴'
query_url = 'https://www.amazon.com/bestsellers?Id=NSGoogle'
sec_names = ['Amazon Devices & Accessories', 'Amazon Launchpad', 'Appliances', 'Apps & Games', 'Arts, Crafts & Sewing', 'Audible Books & Originals', 'Automotive', 'Baby', 'Beauty & Personal Care', 'Books', 'CDs & Vinyl', 'Camera & Photo Products', 'Cell Phones & Accessories', 'Clothing, Shoes & Jewelry', 'Collectible Currencies', 'Computers & Accessories', 'Digital Educational Resources', 'Digital Music', 'Electronics', 'Entertainment Collectibles', 'Gift Cards', 'Grocery & Gourmet Food', 'Handmade Products', 'Health & Household', 'Home & Kitchen', 'Industrial & Scientific', 'Kindle Store', 'Kitchen & Dining', 'Magazine Subscriptions', 'Movies & TV', 'Musical Instruments', 'Office Products', 'Patio, Lawn & Garden', 'Pet Supplies', 'Software', 'Sports & Outdoors', 'Sports Collectibles', 'Tools & Home Improvement', 'Toys & Games', 'Video Games']

while True:
    sec = int(input('''
1.Amazon Devices & Accessories  2.Amazon Launchpad              3.Appliances                    
4.Apps & Games                  5.Arts, Crafts & Sewing         6.Audible Books & Originals     
7.Automotive                    8.Baby                          9.Beauty & Personal Care        
10.Books                        11.CDs & Vinyl                  12.Camera & Photo Products      
13.Cell Phones & Accessories    14.Clothing, Shoes & Jewelry    15.Collectible Currencies       
16.Computers & Accessories      17.Digital Educational Resources18.Digital Music                
19.Electronics                  20.Entertainment Collectibles   21.Gift Cards                   
22.Grocery & Gourmet Food       23.Handmade Products            24.Health & Household           
25.Home & Kitchen               26.Industrial & Scientific      27.Kindle Store                 
28.Kitchen & Dining             29.Magazine Subscriptions       30.Movies & TV                  
31.Musical Instruments          32.Office Products              33.Patio, Lawn & Garden         
34.Pet Supplies                 35.Software                     36.Sports & Outdoors            
37.Sports Collectibles          38.Tools & Home Improvement     39.Toys & Games                 
40.Video Games

1.위 분야 중에서 자료를 수집할 분야의 번호를 선택하세요 : '''))

    if ((sec > 0) & (sec < 41)):
        break
    else:
        print("잘못된 번호를 입력하였습니다.")

while True:
    cnt = int(input("2. 해당 분야에서 크롤링 할 건수는 몇건입니까?(1-100 건 사이 입력) : " ))
    if cnt < 101:
        break
    else:
        print("검색 건수는 1건 - 최대 100 건까지만 가능합니다.")

f_dir = "E:/coding/3years/python/Amazon_Best_Seller_Info_Crawler/"#input("3.파일을 저장할 폴더명만 쓰세요(예 : c:\\temp\\) : ")
print("\n")

now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

sec_name = sec_names[sec - 1]


os.makedirs(f_dir + s + '-'+query_txt + '-' + sec_name)
os.chdir(f_dir + s + '-' + query_txt + '-' + sec_name)

resultDir = f_dir + s + '-' + query_txt + '-' + sec_name
os.makedirs(resultDir + '/images')

resultFile = resultDir + '\\' + s + '-' + query_txt + '-' + sec_name
imageDir = resultFile + '/images/'


path = "E:/coding/3years/chrome driver/chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get(query_url)
time.sleep(1)

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

reple_result = soup.select('#zg_browseRoot > ul')
slist = reple_result[0].find_all('a')

asdf = []
for i in slist:
    asdf.append(i.get_text())
print(asdf)

time.sleep(30)

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
        f = open(resultFile + '.txt', 'a', encoding = 'UTF-8')
        f.write("-"*40 + "\n")

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
                urllib.request.urlretrieve(src, imageDir + str(count) + '.jpg')
                imgs.append(imageDir + str(count) + '.jpg')
            except:
                imgs.append('')
        else:
            imgs.append('')
    
        if count == cnt :
          break
      
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
amazon_best_seller.to_csv(resultFile + '.csv', encoding = "utf-8-sig", index = True)

#엑셀 형태로 저장
amazon_best_seller.to_excel(resultFile + '.xlsx', index = True)

#그림추가
wb = load_workbook(filename = resultFile + 'xlsx', read_only = False, data_only = False)
ws = wb.active

for i in range(0, len(imgs)):
    if(imgs[i] == ''):
        continue
    img = Image(imgs[i])
    
    cellNum = i + 2
    
    ws.row_dimensions[cellNum].height = img.height * 0.75 + 16
    ws.column_dimensions['C'].width = 122
    
    ws.add_image(img, 'C' + str(cellNum))

wb.save(resultFile + '.xlsx')