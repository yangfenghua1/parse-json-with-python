import json,xlwt,jsonpath
def readExcel(file):
    with open(file,'r',encoding='utf8') as fr:
        data = json.load(fr) # 用json中的load方法，将json串转换成字典
    return data
def writeM():
    filename = input("Enter filename: ")
    a = readExcel(filename)
    print(a)
    title = ["该评论cid","该评论点赞","评论内容"]
    book = xlwt.Workbook() # 创建一个excel对象
    sheet = book.add_sheet('Sheet1',cell_overwrite_ok=True) # 添加一个sheet页
    for i in range(len(title)): # 循环列
        sheet.write(0,i,title[i]) # 将title数组中的字段写入到0行i列中
    cid_list=jsonpath.jsonpath(a,"$..cid")
    text_list=jsonpath.jsonpath(a,"$..text")
    digg_count_list=jsonpath.jsonpath(a,"$..digg_count")
    i = 0
    for line in text_list: #　循环字典
        i = i + 1
        sheet.write(int(i),2,line) #　将line写入到第int(line)行，第0列中
    i = 0
    for line in digg_count_list: #　循环字典
        i = i + 1
        sheet.write(int(i),1,line) #　将line写入到第int(line)行，第0列中
    i = 0
    for line in cid_list: #　循环字典
        i = i + 1
        sheet.write(int(i),0,line) #　将line写入到第int(line)行，第0列中
    book.save('output-'+filename+'-excel.xls')

if __name__ == '__main__':
    writeM()