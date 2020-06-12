# -*- coding: utf-8 -*-
"""
Created on Thu Mar 26 10:42:07 2020
企查查爬虫


tips:
    一次最好不要超过50条

gui界面可选项：注册资本 地区 每次查询多少条（如50）

地区那里好像不能一次性拆分
@author: situ
"""

from selenium import webdriver
import time
import os
import pandas as pd
import re
import cpca #cpca是chinese province city area的缩写

def qichacha_batch(driver,file_name,comp_list,j,cur,min_capital,district):
    """处理同一批次的样本"""
    inc_list = comp_list["企业名称"].tolist()
    inc_len = len(inc_list)
    
    
    
    for i in range(inc_len):
        print("正在查询第%d家公司"%i)
        comp_i = inc_list[i]
        time.sleep(3)
            
            
        if (i==0) and (j==0):    
            driver.find_element_by_id('searchkey').send_keys(comp_i)
            # 单击搜索按钮
            srh_btn = driver.find_element_by_xpath('//*[@class="index-searchbtn"]')
            srh_btn.click()
        else:
            driver.find_element_by_id('headerKey').clear()
            driver.find_element_by_id('headerKey').send_keys(comp_i)
            #搜索按钮 
            srh_btn = driver.find_element_by_xpath('//*[@class="btn btn-primary"]')
            srh_btn.click()
        
        try:
            #企查查中企业名称
            inc_full = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/a').text
        #    print(inc_full) 
            comp_list["实际查询名称"][comp_list["企业名称"]==comp_i] = inc_full
            if inc_full != comp_i:
                comp_list["备注"][comp_list["企业名称"]==comp_i] = "查询名称不匹配"
                print(comp_i+"查询名称不匹配")
                
            capital = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[1]/span[1]').text.split("：")[1]
            
        #        print(capital)  
            addr = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[3]').text
        #    print(addr) 
            comp_list["地址"][comp_list["企业名称"]==comp_i] = addr.split("：")[1]
            
            comp_list["币种"][comp_list["企业名称"]==comp_i] = [c for c in cur if c in capital][0]
            
            capital_float = float(re.findall(r'\d+', capital)[0])
            comp_list["注册资本"][comp_list["企业名称"]==comp_i] = capital_float
    #            capital_float = float(capital.split("万")[0])
            
            if len(district)==0:#是否对地区有限制
                condition = capital_float>=min_capital
            else:
                #有限制
                standard_addr = "".join(cpca.transform([addr])[["省","市","区"]].ix[0,:].tolist())
                if len([d for d in district if d in standard_addr])>0:
                    area = [d for d in district if d in standard_addr][0]
                    
                elif len(cpca.transform([addr])["市"][0])>0:
                    area = cpca.transform([addr])["市"][0]
                else:
                    area = addr[3:5]
                    
                comp_list["地区"][comp_list["企业名称"]==comp_i] = area
                
                condition = capital_float>=min_capital and area in district
                
            if condition:
                #成立日期
                date = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[1]/span[2]').text
                comp_list["成立日期"][comp_list["企业名称"]==comp_i] = date.split("：")[1]
                #联系方式
                mail_phone = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[2]').text
                mail_phone = re.sub('更多邮箱', '', mail_phone)
                mail_phone = re.sub("更多号码", '', mail_phone)
                comp_list["联系方式"][comp_list["企业名称"]==comp_i] = mail_phone
    
                # 获取网页地址，进入
                inner = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/a').get_attribute("href")
                
                #是否上市
                search_info = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]').text
                go_pub = len(re.findall(r"A股",search_info))>0 or len(re.findall(r"新三板",search_info))>0
                if go_pub:
                    inner = inner+"#base"
                    tag = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/div/span[1]').text               
                    
                driver.get(inner)    
                
                #企业类型
                comp_type = driver.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[5]/td[2]').text
                comp_list["企业类型"][comp_list["企业名称"]==comp_i] = comp_type
                #经营范围
                scope = driver.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[9]/td[2]').text
                comp_list["经营范围"][comp_list["企业名称"]==comp_i] = scope    
                
                stock_holder =  driver.find_element_by_xpath('//*[@class="seo font-14"]').text
                if go_pub:          
                    # 新三板、A股
#                    stock_holder =  driver.find_element_by_xpath('//*[@id="ipopartnerslist"]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/a').text
                    comp_list["控股信息"][comp_list["企业名称"]==comp_i] = tag + " 大股东："+stock_holder    
                else:                
                    # 普通企业    
#                    stock_holder =  driver.find_element_by_xpath('//*[@id="partnerslist"]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/a').text
                    if len(stock_holder)<5 and len(stock_holder)>0:
                        stock_holder_title = "私人股东："
                    else:
                        stock_holder_title = "大股东："
                    comp_list["控股信息"][comp_list["企业名称"]==comp_i] = stock_holder_title+stock_holder
        except:
            print(comp_i+"搜索无记录")
            comp_list["备注"][comp_list["企业名称"]==comp_i] = "搜索无记录"
    
    if len(district)==0:
        print("对地区无限制")
        comp_list["地区"] = cpca.transform(comp_list["地址"])["市"].tolist()
        print(cpca.transform(comp_list["地址"])["市"])
        print(comp_list["地区"])
        comp_list["地区"][comp_list["地区"]==""] = comp_list["地址"][comp_list["地区"]==""].str[3:5]
        print(comp_list["地址"][comp_list["地区"]==""].str[3:5])
        print(comp_list["地区"])
    
    return comp_list
    


    
    
    
    
def qichacha(file_name,batch,min_capital,district,account,password):
    cur = ["人民币","美元","欧元","加元","日元"]
    comp_list_all = pd.read_excel(file_name,dtype="str")
#    comp_list_all.head()
    ##判断下面要用到的列是否存在，不在则新增列
    col_list = ["备注","实际查询名称","地区","企业类型","经营范围","成立日期","注册资本","联系方式","地址","币种","控股信息"]
    new_col = [c for c in col_list if c not in comp_list_all.columns]   
    comp_list_all[new_col] = pd.DataFrame(columns=new_col)


    # 伪装成浏览器，防止被识破
    option = webdriver.ChromeOptions()
    option.add_argument('--user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.146 Safari/537.36"')
    driver = webdriver.Chrome(os.path.join(os.getcwd(),"chromedriver.exe"),options=option)
    # 打开登录页面
    driver.get('https://www.qichacha.com/user_login')
    # 单击用户名密码登录的标签
    tag = driver.find_element_by_xpath('//*[@id="normalLogin"]')
    tag.click()
    tag = driver.find_element_by_xpath('//*[@class="btn-weibo m-l-xs"]')
    tag.click()
    # 将用户名、密码注入
    driver.find_element_by_id('userId').send_keys(account)
    driver.find_element_by_id('passwd').send_keys(password)
    time.sleep(3)  
    # 休眠，人工完成验证步骤，等待程序单击“登录”# 单击登录按钮
    btn = driver.find_element_by_xpath('//*[@id="outer"]/div/div[2]/form/div/div[2]/div/p/a[1]')
    btn.click()
    
    #btn = driver.find_element_by_xpath('//*[@id="qcccomModal"]/div/div/button')
    #btn.click()
    
    time.sleep(10)# 

    times = round(len(comp_list_all)/batch)
    
    
    for j in range(times):
        if j==(times-1) and (j+1)*batch!=len(comp_list_all):
            comp_list = comp_list_all[(j*batch):len(comp_list_all)]
        else:
            comp_list = comp_list_all[(j*batch):((j+1)*batch)]
        
        print("正在查询第%d批次"%j)
        comp_list = qichacha_batch(driver,file_name,comp_list,j,cur,min_capital,district)
        
        print("qichacha function:")
        print(comp_list["地区"])
        
        
        if times<=1:
            writer = pd.ExcelWriter(file_name.split(".")[0]+"_done.xlsx", engine='xlsxwriter')
        else:    
            writer = pd.ExcelWriter(file_name.split(".")[0]+"_%d.xlsx"%j, engine='xlsxwriter')
        comp_list.to_excel(writer,sheet_name = "sheet1",encoding = "gbk", index=False)
        writer.save()    
        
    driver.close()
    
    #合并
    if times>1:
        file_names = os.listdir(os.getcwd())
        file_names = [f for f in file_names if len(re.findall(file_name.split(".")[0]+"_",f))>0] #只读取后缀名为csv的文件
        all_df = pd.DataFrame()
        for i in range(len(file_names)):
            temp_df = pd.read_excel(os.path.join(os.getcwd(),file_names[i]),dtype="str")
        
            all_df = pd.concat([all_df,temp_df],axis=0,ignore_index=True)
        #        print(all_df.head())
        writer2 = pd.ExcelWriter(file_name.split(".")[0]+"_done.xlsx", engine='xlsxwriter')
        all_df.to_excel(writer2,sheet_name = "sheet1",encoding = "gbk", index=False)
        writer2.save()    
        

def main():
    path = "F:/qichacha"    
    file_name = "sample.xlsx"
    

    # 搜索限制
    min_capital = 1000
#    district = ["昆明","楚雄","曲靖","昭通","玉溪","丽江","安宁","通海","嵩明"]
    district = []
    
    account = '' #微博账号
    password = '' #微博密码
    
    #一次查询几条，建议不超过50
    batch = 50
    
    os.chdir(path)

    qichacha(file_name,batch,min_capital,district,account,password)


if __name__=="__main__":
    main()



        


