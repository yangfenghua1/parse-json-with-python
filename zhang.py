import json,xlwt,jsonpath,urllib3,time,datetime,os
def readExcel(file):
    with open(file,'r',encoding='utf8') as fr:
        data = json.load(fr) # 用json中的load方法，将json串转换成字典
    return data
def Time2ISOString( s ):
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime( float(s) ) ) 
    
def writeM():
    method = input("Enter method: ")
    if (method == 'url'):
        http = urllib3.PoolManager()
        user_id = "98524936524" #98524936524(四平警事) "98098916818"(平安佳木斯)
        page = 0
        req_1 = {"uid":user_id,"condition":"create_time","page":page,"size":10000,"sort":[],"filterDelete":"true"}
        encoded_req_1 = json.dumps(req_1).encode('utf-8')
        r_1 = http.request('POST','https://www.doudouxia.com/data-center/douyin/detail/author',
                body=encoded_req_1,
                headers={'Content-Type':'application/json'})
        if r_1.status == 200:
            response_1 = r_1.data
            if len(response_1) > 0:
                a_1 = json.loads(response_1.decode())

        book = xlwt.Workbook() # 创建一个excel对象
        sheet0 = book.add_sheet("basic data",cell_overwrite_ok=True) # 添加一个sheet页
        title_basic_data = ["视频链接","视频标题","视频点赞数","视频评论数","视频转发数","视频创建时间"]
        for i in range(len(title_basic_data)): # 循环列
            sheet0.write(0,i+2,title_basic_data[i]) # 将title数组中的字段写入到0行i列中
        id_list=jsonpath.jsonpath(a_1,"$.result.content[*].id")
        create_time_list=jsonpath.jsonpath(a_1,"$.result.content[*].create_time")
        desc_list=jsonpath.jsonpath(a_1,"$.result.content[*].desc")
        digg_video_list=jsonpath.jsonpath(a_1,"$.result.content[*].statistics.digg_count")
        comment_video_list=jsonpath.jsonpath(a_1,"$.result.content[*].statistics.comment_count")
        share_video_list=jsonpath.jsonpath(a_1,"$.result.content[*].statistics.share_count")
        share_url_list=jsonpath.jsonpath(a_1,"$.result.content[*].share_url")        
        nick_name_list=jsonpath.jsonpath(a_1,"$.result.content[*].author.nickname")
        array = []
        array1 = []
        if create_time_list:
            for video_id in create_time_list:
                if video_id in range(1546272000,1577808000):
                    array.append(1)
                else:
                    print(video_id)
                    array.append(0)
            print(array)
            print("count 1:",array.count(1),"count 0:",array.count(0))
            for digg_id in digg_video_list:
                if digg_id in range(6000,100000000):
                    array1.append(1)
                else:
                    array1.append(0)
            print(array1)
            print(digg_video_list)
            print("count 1:",array1.count(1),"count 0:",array1.count(0))
            j = 0
            k = 0#basic data index
            for video_id in id_list: #　循环字典
                if array[j]+array1[j] < 2:
                    print(j,"out of date or small digg!!!!!")
                    j = j+1
                    continue
                else:
                    k = k+1
                    sheet0.write(k,0,nick_name_list[j])
                    sheet0.write(k,1,k)
                    sheet0.write(k,2,share_url_list[j])
                    sheet0.write(k,3,desc_list[j])
                    sheet0.write(k,4,digg_video_list[j])
                    sheet0.write(k,5,comment_video_list[j])
                    sheet0.write(k,6,share_video_list[j])
                    sheet0.write(k,7,Time2ISOString(create_time_list[j]))
                    print(j,"in range of 2019-01-01---2019-12-31")
                    print(video_id,len(id_list))
                    req_2 = {"id":video_id,"page":0,"size":100000}
                    encoded_req_2 = json.dumps(req_2).encode('utf-8')
                    r_2 = http.request('POST','https://www.doudouxia.com/data-center/douyin/comment/detail',
                            body=encoded_req_2,
                            headers={'Content-Type':'application/json'})
                    if r_2.status == 200:
                        response_2 = r_2.data
                        if len(response_2) > 0:
                            a_2 = json.loads(response_2.decode())
                        else:
                            print("length of response_2 < 0")
                    else:
                        print("status != 200")
                    title = ["该评论cid","该评论点赞","评论内容"]
                    sheet = book.add_sheet(str(k),cell_overwrite_ok=True) # 添加一个sheet页
                    for i in range(len(title)): # 循环列
                        sheet.write(3,i,title[i]) # 将title数组中的字段写入到0行i列中
                    cid_list=jsonpath.jsonpath(a_2,"$.result.content[*].cid")
                    text_list=jsonpath.jsonpath(a_2,"$.result.content[*].text")
                    digg_count_list=jsonpath.jsonpath(a_2,"$.result.content[*].digg_count")
                    if text_list:
                        i = 3
                        for line in text_list: #　循环字典
                            i = i + 1
                            sheet.write(int(i),2,line) #　将line写入到第int(line)行，第2列中
                    else:
                        print("text_list is empty")
                        continue
                    if digg_count_list:
                        i = 3
                        for line in digg_count_list: #　循环字典
                            i = i + 1
                            sheet.write(int(i),1,line) #　将line写入到第int(line)行，第1列中
                    else:
                        print("digg_count_list is empty")
                        continue
                    if cid_list:
                        i = 3
                        for line in cid_list: #　循环字典
                            i = i + 1
                            sheet.write(int(i),0,line) #　将line写入到第int(line)行，第0列中
                    else:
                        print("cid_list is empty")
                        continue
                    if desc_list:
                        sheet.write(1,0,"视频标题")
                        sheet.write(2,0,desc_list[j])
                    if digg_video_list:
                        sheet.write(1,1,"视频点赞数")
                        sheet.write(2,1,digg_video_list[j])
                    if comment_video_list:
                        sheet.write(1,2,"视频评论数")
                        sheet.write(2,2,comment_video_list[j])
                    if share_video_list:
                        sheet.write(1,3,"视频转发数")
                        sheet.write(2,3,share_video_list[j])
                    if create_time_list:
                        sheet.write(1,4,"视频创建时间")
                        sheet.write(2,4,Time2ISOString(create_time_list[j]))  
                    j = j+1
        book.save('url-output-'+nick_name_list[1]+'-'+str(page)+'-excel.xls')
        print("url request no impl")
    else:
        filename = input("Enter filename: ")
        a = readExcel(filename)
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
            sheet.write(int(i),2,line) #　将line写入到第int(line)行，第2列中
        i = 0
        for line in digg_count_list: #　循环字典
            i = i + 1
            sheet.write(int(i),1,line) #　将line写入到第int(line)行，第1列中
        i = 0
        for line in cid_list: #　循环字典
            i = i + 1
            sheet.write(int(i),0,line) #　将line写入到第int(line)行，第0列中
        book.save('output-'+filename+'-excel.xls')

if __name__ == '__main__':
    writeM()