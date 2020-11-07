#98524936524(四平警事) 
#105013482132(盘锦网警)
#81836097743(双鸭山公安)
#104603097611(伊春市公安局反诈骗中心) 
#3555212221(平安鹤岗) 
#98098916818(平安佳木斯) 
#89924583563(雪城公安) 
#84429877784(鹤城公安) 
#94719433130(大连e警视)
#104379730631(平安哈尔滨) 

import json,xlwt,jsonpath,urllib3,time,datetime,os
def readExcel(file):
    with open(file,'r',encoding='utf8') as fr:
        data = json.load(fr) # 用json中的load方法，将json串转换成字典
    return data
def Time2ISOString( s ):
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime( float(s) ) ) 
    
def writeM(user_id,digg_ref):
    method = 'url'#input("Enter method: ")
    if (method == 'url'):
        http = urllib3.PoolManager()
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
        http = urllib3.PoolManager()
        nick_name = ""
        k = 0#basic data index
        for page in range(10):
            req_1 = {"uid":user_id,"condition":"create_time","page":page,"size":10000,"sort":[],"filterDelete":"true"}
            encoded_req_1 = json.dumps(req_1).encode('utf-8')
            r_1 = http.request('POST','https://www.doudouxia.com/data-center/douyin/detail/author',
                    body=encoded_req_1,
                    headers={'Content-Type':'application/json'})
            if r_1.status == 200:
                response_1 = r_1.data
                if len(response_1) > 0:
                    a_1 = json.loads(response_1.decode())
            if digg_ref == 10000:
                id_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 10000)].id")
                create_time_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 10000)].create_time")
                desc_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 10000)].desc")
                digg_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 10000)].statistics.digg_count")
                comment_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 10000)].statistics.comment_count")
                share_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 10000)].statistics.share_count")
                share_url_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 10000)].share_url")        
                nick_name_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 10000)].author.nickname")
            elif digg_ref == 8000:
                id_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 8000)].id")
                create_time_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 8000)].create_time")
                desc_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 8000)].desc")
                digg_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 8000)].statistics.digg_count")
                comment_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 8000)].statistics.comment_count")
                share_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 8000)].statistics.share_count")
                share_url_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 8000)].share_url")        
                nick_name_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 8000)].author.nickname")
            else:
                id_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 6000)].id")
                create_time_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 6000)].create_time")
                desc_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 6000)].desc")
                digg_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 6000)].statistics.digg_count")
                comment_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 6000)].statistics.comment_count")
                share_video_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 6000)].statistics.share_count")
                share_url_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 6000)].share_url")        
                nick_name_list=jsonpath.jsonpath(a_1,"$.result.content[?(@.create_time in range(1546272000,1577808000) and @.share_url and @.statistics.digg_count > 6000)].author.nickname")
            if nick_name_list:
                nick_name = nick_name_list[1]
                print(nick_name,"start")
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
                if array.count(1) == 0:
                    print("page:",page,"has no video within 2019")
                    continue
                else:
                    print("page:",page,"has",array.count(1),"videos within 2019")
                for digg_id in digg_video_list:
                    if digg_id in range(digg_ref,100000000):
                        array1.append(1)
                    else:
                        array1.append(0)
                print(array1)
                print("count 1:",array1.count(1),"count 0:",array1.count(0))
                j = 0
                video_continue = 0
                for video_id in id_list: #　循环字典
                    if array[j]+array1[j] < 2:
                        print(j,"out of date or small digg!!!!!")
                        j = j+1
                        continue
                    else:
                        print(j,digg_video_list[j],Time2ISOString(create_time_list[j]))
                        k = k+1
                        #写当前视频的基础数据
                        sheet0.write(k,0,nick_name_list[j])
                        sheet0.write(k,1,k)
                        print("length of share_url_list:",len(share_url_list))
                        print("length of id_list:",len(id_list))
                        print("length of create_time_list:",len(create_time_list))
                        print("length of desc_list:",len(desc_list))
                        print("length of digg_video_list:",len(digg_video_list))
                        print("length of comment_video_list:",len(comment_video_list))
                        print("length of share_video_list:",len(share_video_list))
                        print("length of nick_name_list:",len(nick_name_list))

                        sheet0.write(k,2,share_url_list[j])
                        sheet0.write(k,3,desc_list[j])
                        sheet0.write(k,4,digg_video_list[j])
                        sheet0.write(k,5,comment_video_list[j])
                        sheet0.write(k,6,share_video_list[j])
                        sheet0.write(k,7,Time2ISOString(create_time_list[j]))

                        title = ["该评论cid","该评论点赞","评论内容"]
                        index_comment = 3
                        temp_index_comment = index_comment
                        sheet = book.add_sheet(str(k),cell_overwrite_ok=True) # 添加一个sheet页
                        for i in range(len(title)): # 循环列
                            sheet.write(index_comment,i,title[i]) # 将title数组中的字段写入到0行i列中
                        for page_comment in range(10):
                            index_comment = temp_index_comment
                            req_2 = {"id":video_id,"page":page_comment,"size":1000}
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
                            cid_list=jsonpath.jsonpath(a_2,"$.result.content[*].cid")
                            text_list=jsonpath.jsonpath(a_2,"$.result.content[*].text")
                            digg_count_list=jsonpath.jsonpath(a_2,"$.result.content[*].digg_count")
                            if text_list:
                                index_comment = temp_index_comment
                                for line in text_list: #　循环字典
                                    index_comment = index_comment + 1
                                    sheet.write(int(index_comment),2,line) #　将line写入到第int(line)行，第2列中
                            else:
                                print("text_list is empty")
                            if digg_count_list:
                                index_comment = temp_index_comment
                                for line in digg_count_list: #　循环字典
                                    index_comment = index_comment + 1
                                    sheet.write(int(index_comment),1,line) #　将line写入到第int(line)行，第1列中
                            else:
                                print("digg_count_list is empty")
                            if cid_list:
                                index_comment = temp_index_comment
                                for line in cid_list: #　循环字典
                                    index_comment = index_comment + 1
                                    sheet.write(int(index_comment),0,line) #　将line写入到第int(line)行，第0列中
                            else:
                                print("cid_list is empty")
                            temp_index_comment = index_comment
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
        book.save('url-output-'+nick_name+'-'+str(page)+'-excel.xls')
        print(nick_name,"end")
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
    user_id_list = ["98524936524","105013482132","81836097743","104603097611","3555212221","98098916818","89924583563","84429877784","94719433130","104379730631"]
    digg_ref = [10000,10000,10000,8000,8000,6000,6000,6000,6000,6000]
    for i in range(10):
        writeM(user_id_list[i],digg_ref[i])