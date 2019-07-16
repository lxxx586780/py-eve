import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl import Workbook
from openpyxl import load_workbook


def search_data(link):
    print("a-----------------------------------------------------------------")
    response = requests.get("https://zkillboard.com/ship/17738/losses/page/2/")
    # print(response.text)
    soup = BeautifulSoup(response.content, 'lxml')
    ktbody = soup.find_all('tbody')[2]
    ktr = ktbody.find_all('tr')

    # wb =openpyxl.Workbook()
    # sheet = wb['Sheet']
    wb = load_workbook('测试写数据.xlsx')
    wb.guess_types = True
    sheet=wb.active
    rows=[]
    for row in sheet.rows:
        rows.append(row[0].value+' '+row[1].value+' '+row[2].value)
        
    # for rowt in rows:
    #     print(rowt)
    #创建excel文档第一行
    # sheet['A1'] = '击毁日期'
    # sheet['B1'] = '击毁时间'
    # sheet['C1'] = '价值'
    # sheet['D1'] = '具体损毁链接'
    # sheet['E1'] = '安全等级'
    # sheet['F1'] = '作战星系'
    # sheet['G1'] = '作战星域'
    # sheet['H1'] = '玩家ID'
    # sheet['I1'] = '玩家所属公司'
    #创建样式（宽度）
    # sheet.column_dimensions['A'].width = 18
    # sheet.column_dimensions['D'].width = 40.0
    # sheet.column_dimensions['G'].width = 20.0
    # sheet.column_dimensions['H'].width = 30.0
    # sheet.column_dimensions['I'].width = 35.0
    # #行高
    # sheet.row_dimensions[1].height = 10
    i=1
    for kth in ktr:
        print("-----------------------------start---------------------------")
        kr = kth.find_all('th')  # 寻找日期
        if (kr):
            s_date = kr[0].get_text()

        ktd = kth.find_all('td')
        if (ktd):
            # print('击毁日期: ' + s_date)  # 数据来自上一个tr

            s_time = ktd[0].get_text()
            # print('击毁时间: ' + s_time[1:6])
            s_value = ktd[0].a.get_text()
            # print('价值：' + s_value)
            s_lostlink = 'https://zkillboard.com' + ktd[0].a['href']
            # print('具体损毁链接：' + s_lostlink)

            s_safelevel = ktd[2].find_all('span')[0].get_text()
            # print('安全等级：' + s_safelevel)
            s_location_xi = ktd[2].find_all('a')[0].get_text()
            # print('作战星系：' + s_location_xi)
            s_location_yu = ktd[2].find_all('a')[1].get_text()
            # print('作战星域：' + s_location_yu)

            s_playerId = ktd[4].find_all('a')[0].get_text()
            # print('玩家ID：' + s_playerId)
            s_playerCom = ktd[4].find_all('a')[1].get_text()
            # print('玩家所属公司：' + s_playerCom)

            # 在添加数据之前，要进行数据唯一性检查，主要参考 日期+时间（和价值量）
            # first 组合数据
            s = s_date+' '+s_time[1:6]+' '+s_value
            #匹配,有问题
            for rowt in rows:
                if s != rowt: #if语句，不加括号！！！切记
                    sheet.append([s_date, s_time[1:6], s_value, s_lostlink, s_safelevel, s_location_xi, s_location_yu, s_playerId, s_playerCom])
                else:
                    break
                
            #print(s)


            #sheet.append([s_date, s_time[1:6], s_value, s_lostlink, s_safelevel
            #, s_location_xi, s_location_yu, s_playerId, s_playerCom])
            
            # sheet['A'+str(i)] =s_date
            # sheet['B'+str(i)] =s_time[1:6]
            # sheet['C'+str(i)] =s_value
            # sheet['D'+str(i)] =s_lostlink
            # sheet['E'+str(i)] =s_safelevel
            # sheet['F'+str(i)] =s_location_xi
            # sheet['G'+str(i)] =s_location_yu
            # sheet['H'+str(i)] =s_playerId
            # sheet['I'+str(i)] =s_playerCom
     
        # 进行数据汇总
        print('------------------------------end--------------------------')
        # if(i==10):
        #     break
        # else:
        i=i+1
       
    wb.save('测试写数据.xlsx')


if __name__ == '__main__':
    search_data('a')
