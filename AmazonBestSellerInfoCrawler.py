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
sec_names = ['Amazon Devices & Accessories', 'Amazon Launchpad', 'Appliances', 'Apps & Games', 'Arts, Crafts & Sewing',
            'Audible Books & Originals', 'Automotive', 'Baby', 'Beauty & Personal Care', 'Books', 'CDs & Vinyl', 'Camera & Photo',
            'Cell Phones & Accessories', 'Clothing, shoes & Jewelry', 'Collectible Currencies', 'Computers & Accessories', 
            'Digital Music', 'Electronics', 'Entertainment Collectibles', 'Gift Cards', 'Grocery & Gourmet Food',
            'Handmade Products', 'Health & Household', 'Home & Kitchen', 'Industrial & Scientific', 'Kindle Store',
            'Kitchen & Dining', 'Magazine Subscriptions', 'Movies & TV', 'Musical Instruments', 'Office Products',
            'Patio, Lawn & Garden', 'Pet Supplies', 'Prime Pantry', 'Smart Home', 'Software', 'Sports & Outdoors',
            'Sports Collectibies', 'Tools & Home Improvemet', 'Toys & Games', 'Video Games']

print('실행')
sec = int(input('''
        1.Amazon Devices & Accessories     2.Amazon Launchpad               3.Appliances
        4.Apps & Games                     5.Arts, Crafts & Sewing          6.Audible Books & Originals
        7.Automotive                       8.Baby                           9.Beauty & Personal Care
        10.Books                           11.CDs & Vinyl                   12.Camera & Photo
        13.Cell Phones & Accessories       14.Clothing, shoes & Jewelry     15.Collectible Currencies
        16.Computers & Accessories         17.Digital Music                 18.Electronics
        19.Entertainment Collectibles      20.Gift Cards                    21.Grocery & Gourmet Food
        22.Handmade Products               23.Health & Household            24.Home & Kitchen
        25.Industrial & Scientific         26.Kindle Store                  27.Kitchen & Dining
        28.Magazine Subscriptions          29.Movies & TV                   30.Musical Instruments
        31.Office Products                 32.Patio, Lawn & Garden          33.Pet Supplies
        34.Prime Pantry                    35.Smart Home                    36.Software
        37.Sports & Outdoors               38.Sports Collectibies           39.Tools & Home Improvemet
        40.Toys & Games                    41.Video Games

        1.위 분야 중에서 자료를 수집할 분야의 번호를 선택하세요 : '''))
cnt = int(1)#int(input("        2. 해당 분야에서 크롤링 할 건수는 몇건입니까?(1-100 건 사이 입력) : " ))
f_dir = "E:/coding/3years/python/Amazon_Best_Seller_Info_Crawler/"#input("        3.파일을 저장할 폴더명만 쓰세요(예 : c:\\temp\\) : ")
print("\n")

if cnt > 30:
    print("요청 건수가 많아서 시간이 제법 소요되오니 잠시만 기다려 주세요~~")
else:
    print("요청하신 데이터를 수집하고 있으니 잠시만 기다려 주세요~~")


now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

sec_name = sec_names[sec - 1]

os.makedirs(f_dir + s + '-'+query_txt + '-' + sec_name)
os.chdir(f_dir + s + '-' + query_txt + '-' + sec_name)

ff_dir = f_dir + s + '-' + query_txt + '-' + sec_name
os.makedirs(ff_dir + '/images')


ff_name = ff_dir + '\\' + s + '-' + query_txt + '-' + sec_name + '.txt'
fc_name = ff_dir + '\\' + s + '-' + query_txt + '-' + sec_name + '.csv'
fx_name = ff_dir + '\\' + s + '-' + query_txt + '-' + sec_name + '.xlsx'
fp_name = ff_dir + '\\images\\'


s_time = time.time()

path = "E:/coding/3years/chrome driver/chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get(query_url)
time.sleep(1)

driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[""" + str(sec) + """]/a""").click()

time.sleep(1)

def scroll_down(driver):
    driver.execute_script("window.scrollBy(0, 9300);")
    time.sleep(1)

scroll_down(driver)

bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

reple_result = soup.select('#zg-center-div > #zg-ordered-list')
slist = reple_result[0].find_all('li')

ranking2 = []
title3 = []
price2 = []
score2 = []
sat_count2 = []
store2 = []
srcs = []
imgs = []


if cnt < 51:

    count = 0

    for li in slist:
        f = open(ff_name, 'a', encoding = 'UTF-8')
        f.write("-"*40 + "\n")

        #판매순위
        print("-"*70)
        try:
          ranking = li.find('span', class_='zg-badge-text').get_text().replace("#", "")
        except AttributeError:
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
          price = li.find('span', 'p13n-sc-pric').get_text().replace('\n', '')
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
        srcs.append(src)
    
        try:
          sat_count2.append(sat_count)
        except IndexError:
          sat_count2.append(0)
    
        score2.append(score)
    
        if count == cnt :
          break

elif cnt >= 51 :

    count = 0

    for li in slist:
        f = open(ff_name, 'a', encoding = 'UTF-8')
        f.write("-"*40 + "\n")

        #판매순위
        print("-"*70)
        try:
            ranking = li.find('span', class_='zg-badge-text').get_text().replace("#", "")
        except AttributeError:
            ranking = ''
            print(ranking.replace("#", ""))
        else:
            print("1.판매순위 :", ranking)
    
            f.write('1.판매순위 :' + ranking + '\n')
          
        #제품 이미지
        try:
            src  = li.find('div', class_='a-section a-spacing-mini').find('img')['src']
        except AttributeError:
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
            price = li.find('span', 'p13n-sc-pric').get_text().replace('\n', '')
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
        srcs.append(src)
    
        try:
            sat_count2.append(sat_count)
        except IndexError:
            sat_count2.append(0)
    
        score2.append(score)
    
    
    #1 페이지 추출 후 2페이지로 넘어감
    driver.find_element_by_xpath("""//*[@id="zg-center-div"]/div[2]/div/ul/li[3]/a""").click()
    print("\n")
    print("요청하신 데이터의 수량이 많아 다음 페이지의 데이터를 추출 중이오니 잠시만 기다려 주세요~^^")
    print("\n")

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    reple_result = soup.select('#zg-center-div > #zg-ordered-list')
    slist = reple_result[0].find_all('li')

    for li in slist:
        f = open(ff_name, 'a', encoding = 'UTF-8')
        f.write("-"*40 + "\n")

        #판매순위
        print("-"*70)
        try:
            ranking = li.find('span', class_='zg-badge-text').get_text().replace("#", "")
        except AttributeError:
            ranking = ''
            print(ranking.replace("#", ""))
        else:
            print("1.판매순위 :", ranking)

        f.write('1.판매순위 :' + ranking + '\n')
      
        #제품 이미지
        try:
            src  = li.find('div', class_='a-section a-spacing-mini').find('img')['src']
        except AttributeError:
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
            price = li.find('span', 'p13n-sc-pric').get_text().replace('\n', '')
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
        srcs.append(src)

        try:
          sat_count2.append(sat_count)
        except IndexError:
          sat_count2.append(0)
    
        score2.append(score)
    
        if count == cnt :
          break

else:
  print("검색 건수는 1건 - 최대 100 건까지만 가능합니다.")

driver.quit()

#이미지 저장
for i in range(0, len(srcs)):
    if(srcs[i] != ''):
        try:
            urllib.request.urlretrieve(srcs[i], fp_name + str(i) + '.jpg')
            imgs.append(fp_name + str(i) + '.jpg')
        except:
            imgs.append('')

#Step 5. 검색 결과를 다양한 형태로 저장하기

amazon_best_seller = pd.DataFrame()
amazon_best_seller['판매순위'] = ranking2
amazon_best_seller['제품소개'] = pd.Series(title3)
amazon_best_seller['판매가격'] = pd.Series(price2)
amazon_best_seller['상품평 갯수'] = pd.Series(sat_count2)
amazon_best_seller['상품평점'] = pd.Series(score2)

#csv 형태로 저장
amazon_best_seller.to_csv(fc_name, encoding = "utf-8-sig", index = True)

#엑셀 형태로 저장하기
amazon_best_seller.to_excel(fx_name, index = True)

#그림추가
wb = load_workbook(filename = fx_name, read_only = False, data_only = False)
ws = wb.active

for i in range(0, len(imgs)):
    if(imgs[i] != ''):
        img = Image(imgs[i])
        
        cellNum = i + 2
        
        ws.row_dimensions[cellNum].height = img.height * 0.75 + 15
        ws.column_dimensions['C'].width = img.width * 0.125
        
        ws.add_image(img, 'C' + str(cellNum))

wb.save(fx_name)



e_time = time.time()
t_time = e_time - s_time

#txt 파일에 크롤링 용약 정보 저장하기
orig_stdout = sys.stdout
f = open(ff_name, 'a', encoding = 'UTF-8')
sys.stdout = f

#Step 6. 요약 정보 출력하기
print('\n')
print('=' * 50)
print('총 소요시간은 %s 초 이며, ' %t_time)
print('총 저장 건수는 %s 건 입니다.' %count)
print('=' *50)

sys.stdout = orig_stdout
f.close()

print('\n')
print('=' * 80)
print('1.요청된 총 %s 건의 리뷰 중에서 실제 크롤링 된 리뷰수는 %s 건 입니다.' %(cnt, count))
print('2.총 소요시간은 %s 초 입니다.' %round(t_time, 1))
print('3.파일 저장 완료 : txt 파일명 : %s ' %ff_name)
print('3.파일 저장 완료 : csv 파일명 : %s ' %fc_name)
print('3.파일 저장 완료 : xls 파일명 : %s ' %fx_name)
print('=' * 80)



