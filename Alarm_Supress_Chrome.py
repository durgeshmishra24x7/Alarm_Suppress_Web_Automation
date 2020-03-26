from selenium import webdriver
from selenium.webdriver.common.keys import Keys 
import time
import xlrd
import xlsxwriter

class test1():
    
    url = "http://mas-oakhill.sso.infra.ftgroup/ms/main.do "
    driver=webdriver.Chrome("C:/Users/SHBG7410/Desktop/python/Python/Driver/chromedriver.exe")
    
    '''DEFINING FUNCTION 1'''
    def step1(self,url,driver):
        self.driver=driver
        self.url=url
        self.driver.maximize_window()
        self.driver.get(url)
        self.driver.implicitly_wait(60)
        #return(url,driver)

'''CREATING OBJECTS '''

print("I am Working, Appriciate your patience")

print("LIST of Devices not found in database :- ")

test1obj=test1()
test1obj.step1(test1obj.url,test1obj.driver)

time.sleep(10)

class test2():
    
    '''#DEFINING FUNCTION 2'''
    def step2(self):
        global q,q1
        test1obj.driver.switch_to.frame("navFrame")
        q= test1obj.driver.find_element_by_xpath(".//*[@href='activityController.do?id=-1&createNewActivity=on']")
        #("/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/p/a[1]")
        q.click()
        time.sleep(2)
        test1obj.driver.switch_to.default_content()
        test1obj.driver.switch_to.frame("mainFrame")
        q1=test1obj.driver.find_element_by_xpath("(.//*[@type='radio'])[3]")
        #("/html/body/form/table[2]/tbody/tr/td/table/tbody/tr[1]/td[2]/input[1]")
        q1.click()


'''CREATING OBJECTS'''
test2obj=test2()
test2obj.step2()

time.sleep(2)

'''++++ READING DATA FROM EXCEL SPREADSHEET +++++++++++++++++++'''
def open_file(path):
    Wb=xlrd.open_workbook(path)
    first_sheet=Wb.sheet_by_index(0)
    global L,var,c,var2,k,q2,q3,pre,rr,rr2
    
    var=(first_sheet.nrows)
    var2=(first_sheet.ncols)
    rr=0
    for j in range(2,var2):
        #print("Column ",j)
        c=(first_sheet.col_values(j))
        k=[]
        for n in range(1,var):
            k=c[n]
            rr=rr+1
            test1obj.driver.switch_to.default_content()
            test1obj.driver.switch_to.frame("mainFrame")
            q2=test1obj.driver.find_element_by_xpath(".//*[@name='newServiceId']")
            #("/html/body/form/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/input")
            q2.send_keys(k)

            test1obj.driver.switch_to.default_content()
            test1obj.driver.switch_to.frame("mainFrame")
            q3=test1obj.driver.find_element_by_xpath(".//*[@href='javascript: addServiceId()']")
            #("/html/body/form/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/a[1]")
            q3.click()

    #+++++++++++++ SUBMIT FOR APROVAL ++++++++++++++++++++++++
    test1obj.driver.switch_to.default_content()
    test1obj.driver.switch_to.frame("mainFrame")
    test1obj.driver.find_element_by_xpath("(.//*[@value='Submit For Approval'])").click()
#+++++++++++++++YES/NO++++++++++++++++++++++++++++++++++++++
    test1obj.driver.switch_to.default_content()
    test1obj.driver.switch_to.frame("mainFrame")
    test1obj.driver.find_element_by_xpath("/html/body/form/center/input[1]").click()

#++++++++++++read_List_Elements_++++++++++++++
    #test1obj.driver.switch_to.default_content()
    #test1obj.driver.switch_to.frame("mainFrame")
    global k2,star
    rr2=0
    for nn in range (1,var):
        d=nn
        cc=str(d)
        time.sleep(1)
        test1obj.driver.switch_to.default_content()
        test1obj.driver.switch_to.frame("mainFrame")
        k2=test1obj.driver.find_element_by_xpath("/html/body/form/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/select/option["+cc+"]")
        k2.click()
       # rr2=0
        if "(*)" in k2.text:
         #   rr2=rr2+1
            star=k2.text
            print(star)
            k2.click()
            rr2+=1
        else:
            continue
            
    test1obj.driver.switch_to.default_content()
    test1obj.driver.switch_to.frame("mainFrame")
    test1obj.driver.find_element_by_xpath("/html/body/form/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/a[3]").click()
   
    #+++++++++++++++ Schedule+++++++++++++++++++++++++
    #print(rr2,"   rr   ",rr)
    if rr2 != rr :
         test1obj.driver.switch_to.default_content()
         test1obj.driver.switch_to.frame("mainFrame")
         test1obj.driver.find_element_by_xpath("(.//*[@src='images/schedule_button.gif'])").click()
         time.sleep(1)
    else:
        print("NO EXTRACT FOUND")
        test1obj.driver.quit()
    
    
    #++++++++++++++ DATE PICKER+++++++++++++++++++++++
    sec_sheet=Wb.sheet_by_index(1)
    global va,va2,cal,x,st,start_time,end_time,start_date,end_date
    global mi,mi2,CST,CET
    va=(sec_sheet.nrows)
    #print("Number of rows are",va,".")
    va2=(sec_sheet.ncols)
    #print("Number of Columns are ",va2,".")
#+++++++++++++++++  TIME ++++++++++++++++++++
    stt = sec_sheet.cell(1,0)
    nu=int(round(stt.value))
    nu=str(nu)
#+++++++++++++++++++ CONDITION ++++++++++++++++
    if len(nu)<4:
        av=4-len(nu)
        if av == 1:
            nu="0"+nu
        elif av==2:
            nu="00"+nu
        else:
            nu="000"+nu
    #print(nu)
    CST=nu[:2]+":"+nu[2:4]
    #print(CST)
    for tim in range(1,97):
        mit=str(tim)
        ti=test1obj.driver.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select/option["+mit+"]")
        if CST in ti.text:
            ti.click()
      
    ett = sec_sheet.cell(1,1)
    nu2=int(round(ett.value))
    nu2=str(nu2)
    
    if len(nu2)<4:
        av2=4-len(nu2)
        if av2 == 1:
            nu2="0"+nu2
        elif av2==2:
            nu2="00"+nu2
        else:
            nu2="000"+nu2
    #print(nu2)
    
    CET=nu2[:2]+":"+nu2[2:4]
    #print ("CET",CET)
    
    for tim2 in range(1,97):
        mit2=str(tim2)
        ti2=test1obj.driver.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[4]/select/option["+mit2+"]")
        if CET in ti2.text:
            ti2.click()
    #time.sleep(1)
    global sd2,sy2,sm
#+++++++++++++++++++++DATE++++++++++++++++++
    sd = sec_sheet.cell(1,2)
    sd2=int(round(sd.value))
    sm = sec_sheet.cell(1,3)
    sy= sec_sheet.cell(1,4)
    sy2=int(round(sy.value))
    #print ("Start year :- ",str(sy2))
    global date2
    
#'###################### COMPLETE START DATE ################################################
    d2=str(sd2)+"-"+sm.value+"-"+str(sy2)
    
    #print("Complete Date = ",d)
    edd=sec_sheet.cell(1,5)
    ed2=int(round(edd.value))
    em = sec_sheet.cell(1,6)
    #print ("Start Date :- ",em.value)
    ey= sec_sheet.cell(1,7)
    ey2=int(round(ey.value))
    global date3
    date3=" End Time (in GMT)  : "+str(ed2)+"-"+em.value+"-"+str(ey2)
    #date2=" Start time (in GMT)  : "+str(sd2)+"-"+sm.value+"-"+str(sy2)
    #print("date2:-",date2)
###################### COMPLETE END DATE ################################################    
    ed=str(ed2)+"-"+em.value+"-"+str(ey2)
    #Start Calendar icon
    test1obj.driver.switch_to.default_content()
    test1obj.driver.switch_to.frame("mainFrame")
    st=test1obj.driver.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td[2]/a/img")
    st.click()
#START DATE
    sp=d2.split("-")
    #print(sp)
    date=sp[0]
    #print("Date:- ",date)
    mnth=sp[1]
    #print("Month:- ",mnth)
    year=sp[2]
    #print("Year:- ",year)

    '#++++++++++END DATE+++++++++++++++'
    #ed="7-July-2020"
    ps=ed.split("-")
    #print(sp)
    date2=ps[0]
    #print("Date:- ",date2)
    mnth2=ps[1]
    #print("Month:- ",mnth2)
    year2=ps[2]
    #print("Year:- ",year2)
    
    window_after = test1obj.driver.window_handles[1]
    test1obj.driver.switch_to_window(window_after)
    #print(window_after)
  
#'##################################### DATE PICKER #####################################################'

    #START DATE PICKER (After click on first calendar option)

    ##### MONTH+++++++++
    for re in range(1,13):
        er=str(re)
        v=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[1]/td[2]/select[1]/option["+er+"]")
        #print(er,v.text)
        if mnth in v.text:
            v.click()

     #YEAR       
    time.sleep(1)
    for kr in range(1,201):
        rk=str(kr)
        v2=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[1]/td[2]/select[2]/option["+rk+"]")
        if year in v2.text:
            v2.click()
        else:
            continue

    ds='1'
    
    #DATE
    global v3,gv3,g,g2,gnr,h,vflag
    flag=1
    
    
    for jk in range(2,7):
        kj=str(jk)
        if flag==1:
            for nr in range(1,8):
                rn=str(nr) 
                v3=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[2]/td/table/tbody/tr["+kj+"]/td["+rn+"]")
                                
                if ds in v3.text:
                    flag=0
                    h=nr
                    #print("h:-",h)
                    break
        else:
            break
    vflag=1
    for g in range(2,7):
        g2=str(g)
        if  vflag==1:
            for gnr in range(h,8):
                rng=str(gnr)
                gv3=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[2]/td/table/tbody/tr["+g2+"]/td["+rng+"]")                    
                if date in gv3.text:
                    gv3.click()
                    vflag=0
                    break
        else:
            break
        
    time.sleep(1)
    window_before = test1obj.driver.window_handles[0]
    test1obj.driver.switch_to_window(window_before)
    #print(window_before)
    #Start second Calendar icon

    test1obj.driver.switch_to.default_content()
    test1obj.driver.switch_to.frame("mainFrame")
    st2=test1obj.driver.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td[4]/a/img")
    st2.click()

    window_after = test1obj.driver.window_handles[1]
    test1obj.driver.switch_to_window(window_after)
    
    #print("Month:- ",mnth2)
  
    for fe in range(1,13):
        ef=str(fe)
        y1=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[1]/td[2]/select[1]/option["+ef+"]")
        if mnth2 in y1.text:
            y1.click()
               
    for kr2 in range(1,201):
        rk2=str(kr2)
        y2=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[1]/td[2]/select[2]/option["+rk2+"]")
        if year2 in y2.text:
            y2.click()
        else:
            continue
    #End Date

    ds2='1'
        
    flag2=1
    for jk2 in range(2,7):
        kj2=str(jk2)
        if flag2==1:
            for nr2 in range(1,8):
                rn2=str(nr2)
                y3=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[2]/td/table/tbody/tr["+kj2+"]/td["+rn2+"]")
                if ds2 in y3.text:
                    flag2=0
                    h2=nr2
                    #print("h2",h2)
                    break
        else:
            break

    vflag2=1
    for vg in range(2,7):
        vg2=str(vg)
        if  vflag2==1:
            for gvr in range(h2,8):
                rvg=str(gvr)
                gy3=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[2]/td/table/tbody/tr["+vg2+"]/td["+rvg+"]")
                if date2 in gy3.text:
                    vflag=0
                    gy3.click()
                    break
        else:
            break
        
############# Activity Duration (mins): ############################
    window_before = test1obj.driver.window_handles[0]
    test1obj.driver.switch_to_window(window_before)
    test1obj.driver.switch_to.default_content()
    test1obj.driver.switch_to.frame("mainFrame")
    ad=test1obj.driver.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td[2]")
    #print("ACTIVITY DURATION :- ",ad.text)
    bd=ad.text
    bd=str(bd)
    
    #print("string bd ",bd)

    test1obj.driver.switch_to.default_content()
    test1obj.driver.switch_to.frame("mainFrame")
    #test1obj.driver.switch_to.default_content()
    #test1obj.driver.switch_to.frame("mainFrame")
    eod=test1obj.driver.find_element_by_xpath("(.//*[@name='outageDurationMinQty'])")
    eod.clear()
    eod.send_keys(bd)

    #test1obj.driver.switch_to.default_content()
    #test1obj.driver.switch_to.frame("mainFrame")
    
    #app=test1obj.driver.find_element_by_xpath("(.//*[@name='approveButton])")
    #app.click()
    
    print("EXECUTION COMPLETED")

    

if __name__=="__main__":
    path="C:/Users/SHBG7410/Desktop/python/Python/Book1.xlsx"
    open_file(path)

#++++++++++++++++++++++++++++++++++ WRITE IN EXCEL++++++++++++++++++
global AI
test1obj.driver.switch_to.default_content()
test1obj.driver.switch_to.frame("mainFrame")
AI=test1obj.driver.find_element_by_xpath("/html/body/form/p[2]/table/tbody/tr/td/table/tbody/tr/td[1]/div")
ia=AI.text
ia=str(ia)
ca="Alarm Suppression Activity ID :  "+ia

workbook = xlsxwriter.Workbook('Activity.xlsx') 
worksheet = workbook.add_worksheet() 
f=workbook.add_format({'border':1})
worksheet.write('A1', ca,f)
worksheet.write('A2', 'Oceane ID / Remedy Reference ID : ',f)
worksheet.write('A3','',f)
worksheet.write('A4','Site Name :',f)
worksheet.write('A5','',f)
worksheet.write('A6', 'List of all devices whose alarm needs to be suppressed :',f)
worksheet.write('A7','',f)
worksheet.write('A8','',f)

worksheet.write('A9', 'List of devices which could not be added in MAS tool for suppression :',f)

#worksheet.write('B9',L1,f)
date2=" Start time (in GMT) : "+str(sd2)+"-"+sm.value+"-"+str(sy2)
#print("date2:-",date2)   
worksheet.write('A10','',f)
#print(date2+CST)
worksheet.write('A11', date2+" "+CST,f)
worksheet.write('A12', date3+" "+CET,f)

print("Report Prepared")
workbook.close() 

