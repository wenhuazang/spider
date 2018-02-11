import requests
import xlwt
from lxml import etree
import sys


def get_info(selector):
    info_remarks = selector.xpath("//span[@class='heightbox']")
    content_list = []
    time_list = []
    for i in range(len(info_remarks)):
        content = info_remarks[i].xpath("string(.)").encode(
            sys.stdout.encoding, "ignore").decode(
            sys.stdout.encoding)
        content = "".join(content.split())
        content_list.append(content)
        # print i, content

    info_time = selector.xpath("//span[@class='time_line']")
    for i in range(len(info_time)):
        time = info_time[i].xpath("string(.)").encode(
            sys.stdout.encoding, "ignore").decode(
            sys.stdout.encoding)
        time_list.append(time)
        # print i, time
    return zip(time_list, content_list)


def get_param(url, num_page):

    for num in range(1, num_page):
        params = {
            'poiID': 80633,
            'districtId': 151,
            'districtEName': 'Harbin',
            'pagenow': num,
            'order': 3.0,
            'star': 0.0,
            'tourist': 0.0,
            'resourceId': 20017,
            'resourcetype': 2
        }
        html = requests.get(url, params).content.decode("utf-8")
        selector = etree.HTML(html)
        get_info(selector)


def write_excel(result, sheet, i):
    for j, (t, c) in enumerate(result):
        sheet.write(i*10+j, 0, t)
        sheet.write(i*10+j, 1, c)


def main():
    # note that:
    # parameters num_page is number of page
    # poiID and resourceId, you should modify them from original website
    # parameters save_name is filename that you want to save to hard-disk
    # ########################
    # num_page = 153
    # poiID = 77060
    # resourceId = 7700
    # save_name = 'taiyangdao'
    # ########################
    # ########################
    # num_page = 442
    # poiID = 77064
    # resourceId = 7705
    # save_name = 'shengsuofeiyadajiaotang'
    # ########################
    ############
    num_page = 623
    poiID = 77071
    resourceId = 7712
    save_name = 'zhongyangdajie'
    ########################
    ############
    # num_page = 37
    # poiID = 10758158
    # resourceId = 7701
    # save_name = 'yabulihuxuechang'
    ########################

    url = 'http://you.ctrip.com/destinationsite/TTDSecond/SharedView/AsynCommentView'
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet('sheet 1')
    for num in range(1, num_page+1):
        params = {
            'poiID': poiID,
            'districtId': 151,
            'districtEName': 'Harbin',
            'pagenow': num,
            'order': 3.0,
            'star': 0.0,
            'tourist': 0.0,
            'resourceId': resourceId,
            'resourcetype': 2
        }
        print requests.get(url, params).url
        html = requests.get(url, params).content.decode("utf-8")
        selector = etree.HTML(html)
        results = get_info(selector)
        write_excel(results, sheet, num-1)
    wbk.save('xiecheng_' + save_name + '.xls')


if __name__ == '__main__':
    main()