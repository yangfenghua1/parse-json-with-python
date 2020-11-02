import json,xlwt,jsonpath,urllib3,time,datetime
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
        user_id = "98524936524"
        req_1 = {"uid":user_id,"condition":"create_time","page":0,"size":50,"sort":[],"filterDelete":"true"}
        encoded_req_1 = json.dumps(req_1).encode('utf-8')
        r_1 = http.request('POST','https://www.doudouxia.com/data-center/douyin/detail/author',
                body=encoded_req_1,
                headers={'Content-Type':'application/json'})
        if r_1.status == 200:
            response_1 = r_1.data
            if len(response_1) > 0:
                a_1 = json.loads(response_1.decode())

        book = xlwt.Workbook() # 创建一个excel对象
        id_list=jsonpath.jsonpath(a_1,"$.result.content[*].id")
        create_time_list=jsonpath.jsonpath(a_1,"$.result.content[*].create_time")
        desc_list=jsonpath.jsonpath(a_1,"$.result.content[*].desc")
        digg_video_list=jsonpath.jsonpath(a_1,"$.result.content[*].statistics.digg_count")
        comment_video_list=jsonpath.jsonpath(a_1,"$.result.content[*].statistics.comment_count")
        share_video_list=jsonpath.jsonpath(a_1,"$.result.content[*].statistics.share_count")
        create_video_list=jsonpath.jsonpath(a_1,"$.result.content[*].statistics.create_time")
        print(desc_list[1])
        array = []
        if create_time_list:
            for video_id in create_time_list:
                if video_id in range(1546272000,1577808000):
                    array.append(1)
                else:
                    print(video_id)
                    array.append(0)
            print(array)
            j = 0
            for video_id in id_list: #　循环字典
                if array[j] == 0:
                    print(j,"out of date!!!!!")
                    j = j+1
                    continue
                else:
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
                    sheet = book.add_sheet(video_id,cell_overwrite_ok=True) # 添加一个sheet页
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
                    if create_video_list:
                        sheet.write(1,4,"视频创建时间")
                        sheet.write(2,4,Time2ISOString(create_video_list[j]))  
                    j = j+1
        book.save('url-output-'+user_id+'-excel.xls')
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