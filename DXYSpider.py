# 导入所需要的模块
import requests
import json
import xlwt
import datetime

# 在添加每一个sheet之后，初始化字段
def initXLS():
    global name
    name = ['currentConfirmedCount', 'confirmedCount', 'suspectedCount', 'curedCount', 'deadCount', 'seriousCount', 'currentConfirmedIncr', 'confirmedIncr', 'suspectedIncr', 'curedIncr', 'deadIncr', 'seriousIncr', 'remark1', 'remark2', 'remark3', 'remark4', 'remark5', 'note1', 'note2', 'note3', 'updateTime']

    global row
    global outfile
    global sheet

    row = 0
    for i in range(len(name)):
        sheet.write(row, i, name[i])
    row = row + 1
    outfile.save("./疫情数据.xls")   ############

# 将dic中的内容写入excel
def writeXLS(dic):
    global row
    global outfile
    global sheet

    for k in dic:
        for i in range(len(dic[k])):
            sheet.write(row, i, dic[k][i])
        row = row + 1
    outfile.save("./疫情数据.xls")   #############

def main():
    global outfile
    global sheet

    outfile = xlwt.Workbook(encoding='utf-8')
    # 需要抓取的日期
    today = datetime.date.today()
    sheet = outfile.add_sheet(today.strftime("%Y-%m-%d"))
    initXLS()

    try:
        # 获取自爬虫运行开始（2020年1月24日下午4:00）至今，病毒研究情况以及全国疫情概览，可指定返回数据为最新发布数据或时间序列数据并打印出来。
        r = requests.get("https://lab.isaaclin.cn/nCoV/api/overall?latest=0")
        r.encoding = 'utf-8'
        if r.status_code == 200:
            text = r.text
            state=json.loads(text).get('results')
        #    print(state)
            dic = {}
            for i in range(len(state)):
                values = list(state[i].values())
                if len(name) == len(values):
                    dic[i] = values
                    print('Data', i)
            writeXLS(dic)
        else:
            print('出错了： ', r.status_code)
            print('\n')

        # 获取指定省份的疫情数据，以福建省为例,并打印，里面包括了所有市的具体数据，‘福建省’变量可根据需要进行改变
        # r = requests.get("https://lab.isaaclin.cn/nCoV/api/area?latest=1&province=福建省")
        # r.encoding = 'utf-8'
        # text = r.text
        # state=json.loads(text).get('results')
        # print(state)

    except ValueError:
        print('pyOauth2Error')

if __name__ == '__main__':
    main()
