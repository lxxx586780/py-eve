import requests
from bs4 import BeautifulSoup
print("a-----------------------------------------------------------------")
response = requests.get("https://zkillboard.com/ship/17738/losses/")
# print(response.text)
soup = BeautifulSoup(response.content, 'lxml')
ktbody = soup.find_all('tbody')[2]
ktr = ktbody.find_all('tr')
for kth in ktr:
    kr = kth.find_all('th', class_='no-stripe')  # 寻找日期
    if(kr):
        print(kr[0].string)
    ktd = kth.find_all('td', style='width: 55px;')  # 寻找时间
    if(ktd):
        for ktda in ktd:
            print(ktda.get_text())  # 时间和损失值，没法分开，后期要处理
    ktd2 = kth.find_all('span', style='color: #F30202')  # 寻找时间
    for ktda in ktd2:
        print(ktda.get_text())

    # print(ks)
# print(km)
