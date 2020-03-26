from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import xlrd
import xlsxwriter

class test1():
    
    url = "https://self.sso.infra.ftgroup/logingassifaible.jsp?activateWindows=true&TYPE=33554433&REALMOID=06-000ad14a-2fb1-1b71-8d9e-e8be0aaad064&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=$SM$oZTcp9kVJA%2flPMtmsn9zkq6Iw0B6Jp5IWpl68ZLAVXeSQkGVKLmdn732MFgqX%2bJw&TARGET=$SM$HTTPS%3a%2f%2fself.sso.infra.ftgroup%2fAuthForm%2fredirect.jsp%3fRETURN%3dhttp$%3A%2f%2fmas-oakhill.sso.infra.ftgroup%2fms%2fmain.do"
    #url=("http://mas-oakhill.sso.infra.ftgroup/ms/main.do")
    #driver=webdriver.Chrome("C:/Users/SHBG7410/Desktop/python/Python/Driver/chromedriver.exe")
    driver=webdriver.Firefox()  

    '''DEFINING FUNCTION 1'''
    def step1(self,url,driver):
        self.driver=driver
        self.url=url
        self.driver.maximize_window()
        self.driver.get(url)
        self.driver.implicitly_wait(6)
        #return(url,driver)

'''CREATING OBJECTS '''

print("I am Working, Appriciate your patience")

print("LIST of Devices not found in database :- ")

test1obj=test1()
test1obj.step1(test1obj.url,test1obj.driver)

print("I AM WORKING, APPRECIATE YOUR PATIENCE")
def switch():
    test1obj.driver.switch_to.default_content()
    test1obj.driver.switch_to.frame("mainFrame")
    
class test2():
    
    '''#DEFINING METHOD 2'''
    def step2(self):
        global q,q1,sa,sa2,u2,p2
        global user,pwd,path2,Third_sheet,Wb2
        path2="C:/Users/SHBG7410/Desktop/python/Python/Book1.xlsx"
        #"C:/GSK/Automation/Book1.xlsx"

        Wb2=xlrd.open_workbook(path2)
        Third_sheet=Wb2.sheet_by_index(2)
        user=Third_sheet.cell(1,1)
        pwd=Third_sheet.cell(2,1)        

        sa=test1obj.driver.find_element_by_xpath("//*[@id='user']")
        sa.send_keys(user.value)

        sa2=test1obj.driver.find_element_by_xpath("//*[@id='password']")
        sa2.click()
        sa2.send_keys(pwd.value)
        time.sleep(1)
        test1obj.driver.find_element_by_xpath("//*[@id='spanLinkValidForm']").click()
        
        test1obj.driver.switch_to.frame("navFrame")
        q= test1obj.driver.find_element_by_xpath(".//*[@href='activityController.do?id=-1&createNewActivity=on']")
        q.click()
        
        switch()
        q1=test1obj.driver.find_element_by_xpath("(.//*[@type='radio'])[3]")
        q1.click()
        
'''CREATING OBJECTS'''
test2obj=test2()
test2obj.step2()

time.sleep(1)

'''++++ READING DATA FROM EXCEL SPREADSHEET +++++++++++++++++++'''


def open_file(path):
    Wb=xlrd.open_workbook(path)
    first_sheet=Wb.sheet_by_index(0)
    global L,var,c,var2,k,q2,q3,pre,rr,rr2
    switch()

    var=(first_sheet.nrows)
    #print("Number of rows are",var,".")
    var2=(first_sheet.ncols)
    rr=0
    #'print("Number of Columns are ",var2,".")

    for j in range(2,var2):
     #   'print("Column ",j)
        c=(first_sheet.col_values(j))
        k=[]
        for n in range(1,var):
            k=c[n]
            rr+=1
            switch()
            q2=test1obj.driver.find_element_by_xpath(".//*[@name='newServiceId']")
            q2.send_keys(k)
            switch()
            q3=test1obj.driver.find_element_by_xpath(".//*[@href='javascript: addServiceId()']")
            q3.click()
#+++++++++++++ SUBMIT FOR APROVAL ++++++++++++++++++++++++
    #time.sleep(1)
    switch()
    test1obj.driver.find_element_by_xpath("(.//*[@value='Submit For Approval'])").click()
#+++++++++++++++YES/NO++++++++++++++++++++++++++++++++++++++
    switch()
    test1obj.driver.find_element_by_xpath("(.//*[@name='yesButton'])").click()

#++++++++++++read_List_Elements_++++++++++++++
    global k2,star,L1,z1
    rr2=0
    #L1=[]
    for nn in range (1,var):
       # print("var value ; ",var)
        d=nn
        cc=str(d)
        switch()
        k2=test1obj.driver.find_element_by_xpath("/html/body/form/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/select/option["+cc+"]")
        #/html/body/form/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/select
        k2.click()
        if "(*)" in k2.text:
            star=k2.text
            print(star)
            k2.click()
            rr2+=1
        else:
            continue            
    #time.sleep(1)
    switch()
    test1obj.driver.find_element_by_xpath(".//*[@href='javascript: localRemoveSelected(document.generalTabForm.elements[keyImpactedElementList])']").click()
    #("/html/body/form/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/a[3]").click()
#+++++++++++++++ Schedule+++++++++++++++++++++++++
    if rr2!=rr:
        switch()
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
    switch()
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

    #START DATE PICKER
    for re in range(1,13):
        er=str(re)
        v=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[1]/td[2]/select[1]/option["+er+"]")
        #print(er,v.text)
        if mnth in v.text:
            v.click()
    time.sleep(1)
    for kr in range(1,201):
        rk=str(kr)
        v2=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[1]/td[2]/select[2]/option["+rk+"]")
        if year in v2.text:
            v2.click()
        else:
            continue
    ds='1'
        
    global v3,gv3,g,g2,gnr,h,vflag
    flag=1
    for jk in range(2,7):
        kj=str(jk)
        if flag==1:
            for nr in range(1,8):
                rn=str(nr)
                v3=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[2]/td/table/tbody/tr["+kj+"]/td["+rn+"]")
                #print("jk= ",jk,"nr= ",nr,"v3.text= ",v3.text)
                if ds in v3.text:
                    flag=0
                    h=nr
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

                #print("h",h,"value of gv3 ",gv3.text)
                #print("date :- ",date)
                
                if date in gv3.text:
                    gv3.click()
                    vflag=0
                    break
                else:
                    h=1
        else:
            break
    ###  
    time.sleep(1)
    window_before = test1obj.driver.window_handles[0]
    test1obj.driver.switch_to_window(window_before)
    #print(window_before)
    #Start second Calendar icon

    switch()
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
            #print("YES")      
          
    #print("Year:- ",year2)
    #time.sleep(1)
    for kr2 in range(1,201):
        rk2=str(kr2)
        y2=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[1]/td[2]/select[2]/option["+rk2+"]")
        if year2 in y2.text:
            y2.click()
        else:
            continue
    #End Date
    ds2="1"
    qflag=1
    for k1 in range(2,7):
        qj=str(k1)
        if qflag==1:
            for qnr in range(1,8):
                rnq=str(qnr) 
                qv3=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[2]/td/table/tbody/tr["+qj+"]/td["+rnq+"]")
                                
                if ds2 in qv3.text:
                    qflag=0
                    qh=qnr
                    #print("value of qh:- ",qh)
                    break
        else:
            break
###################################
         
    qflag2=1
    for jk2 in range(2,7):
        kj2=str(jk2)
        if qflag2==1:
            for nr2 in range(qh,8):
                rn2=str(nr2)
                y3=test1obj.driver.find_element_by_xpath("/html/body/center/form/table/tbody/tr[2]/td/table/tbody/tr["+kj2+"]/td["+rn2+"]")
                if date2 in y3.text:
                    qflag2=0
                    y3.click()
                    break
                else:
                    qh=1
        else:
            break

############# Activity Duration (mins): ############################
    window_before = test1obj.driver.window_handles[0]
    test1obj.driver.switch_to_window(window_before)
    switch()
    ad=test1obj.driver.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td[2]")
    #print("ACTIVITY DURATION :- ",ad.text)
    bd=ad.text
    bd=str(bd)
    
    #print("string bd ",bd)

    switch()
    #test1obj.driver.switch_to.default_content()
    #test1obj.driver.switch_to.frame("mainFrame")
    eod=test1obj.driver.find_element_by_xpath("(.//*[@name='outageDurationMinQty'])")
    eod.clear()
    eod.send_keys(bd)

    #switch()
    app=test1obj.driver.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td[2]")
    app.click()
    
    print("EXECUTION COMPLETED")

    

if __name__=="__main__":
    path="C:/GSK/Automation/Book1.xlsx"
    #"C:/Users/SHBG7410/Desktop/python/Python/Book1.xlsx"
    #"C:/GSK/Automation/Book1.xlsx"
    open_file(path)

#++++++++++++++++++++++++++++++++++ WRITE IN EXCEL++++++++++++++++++
global AI
switch()
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

    
    
 

