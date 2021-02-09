from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import TimeoutException
import gspread
from gspread import exceptions
from selenium.webdriver.common.action_chains import ActionChains
from oauth2client.service_account import ServiceAccountCredentials
import sys
import os
import getch,msvcrt
import time
from datetime import datetime,date
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name(

         'login.json', scope) # Your json file here

gc = gspread.authorize(credentials)
sheet = gc.open("All information").worksheet('information')

w='error'
x='none'
y="unavailable"
kin=[]
n=1
glis=[]
glist=[]
lis=[]
ki=66
ni=0
nil=0
dude=1
smart=2

options = webdriver.ChromeOptions() 
options.add_argument("user-data-dir=new user data2")

#options.add_argument("user-data-dir=C:\\Users\\rajra\\AppData\\Local\\Google\\Chrome\\User Data")
#driver= webdriver.Chrome(executable_path=r"driver2/chromedriver.exe",options=options)
#credd=str(input())
print("welcome to the shipment programming session \n 1.This program is designed to automate fedex and bludart shipment tracking.\n 2.please make sure that you have the following prerequisits satisfied.\n  a.unique project name\n  b.Order shipment sheet\n  c.Delivered sheet")
def secure_password_input(prompt=''):
    global p_s
    p_s = ''
    proxy_string = [' '] * 64
    while True:
        sys.stdout.write('\x0D' + prompt + ''.join(proxy_string))
        c = msvcrt.getch()
        if c == b'\r':
            break
        elif c == b'\x08':
            p_s = p_s[:-1]
            proxy_string[len(p_s)] = " "
        else:
            proxy_string[len(p_s)] = "*"
            p_s += c.decode()

    sys.stdout.write('\n')
    #print(p_s)
    return p_s

now = datetime.now()
global current_time
current_time = now.strftime("%I:%M:%p")
#print(current_time)


today = date.today()
global d3
d3 = today.strftime("%d/%m/%y"+"20")
#print( d3)

def startingg():
    #project=str(input())
    credentials = ServiceAccountCredentials.from_json_keyfile_name('cred.json', scope) # Your json file here

    gc = gspread.authorize(credentials)
    print("please Enter the file name of your google project and hit enter key.\n Remember it is case sensitive enter the sheet name as it is.")
    global List3

    List3=str(input())
    print("please Enter the Sheet name of your choice and hit enter key.\n Remember it is case sensitive enter the sheet name as it is.")
    global List 
    List=str(input())
    #Hyd Inventory List
    try:
        global sheet
        sheet=gc.open(List3).worksheet(List)
        global sheet2
        sheet2=gc.open(List3).worksheet("Delivered")
    except Exception as pl:
        print(pl)
        try:
            print("\nthe sheet name you entered is\n either invalid or doesn't exist")
            print("\ntry to enter a valid sheet name")
            List=str(input())
            sheet=gc.open(List3).worksheet(List)
            sheet2=gc.open(List3).worksheet("Delivered")
        except Exception:
            try:
                print("the sheet name you entered is\n either invalid or doesn't exist")
                print("try to enter a valid sheet name")
                List=str(input())
                sheet=gc.open(List3).worksheet(List)
                sheet2=gc.open(List3).worksheet("Delivered")
            except Exception:
                try:
                    print("the sheet name you entered is\n either invalid or doesn't exist")
                    print("try to enter a valid sheet name")
                    List=str(input())
                    sheet=gc.open(List3).worksheet(List)
                    sheet2=gc.open(List3).worksheet("Delivered")
                except Exception:
                    try:
                        print("the sheet name you entered is\n either invalid or doesn't exist")
                        print("try to enter a valid sheet name")
                        List=str(input())
                        sheet=gc.open(List3).worksheet(List)
                        sheet2=gc.open(List3).worksheet("Delivered")
                    except Exception:
                        print("\nyour trials are expired for this sesssion\nplease close the program and run again")
                    else:
                        pass
                else:
                    pass
            else:
                pass
        else:
            pass
    else:
        pass
    #sheets=gc.open("new inventory").worksheet('Sheet3')



    

    now = datetime.now()
    global current_time
    current_time = now.strftime("%I:%M:%p")
    print(current_time)


    today = date.today()
    global d3
    d3 = today.strftime("%d/%m/%y"+"20")
    print( d3)
    print("Enter the starting index of your choice")
    global smart
    smart=int(input())
    print("Enter the end index of your sheet")
    global dude
    dude = int(input())
    
    newll=["shipment"+List,"start"+str(smart),"end"+str(dude),current_time]
    sheet_log.insert_row(newll,3)
    function()

    




def function():
    #nil=0
    print("the program is running....do not close this window terminal")
    driver= webdriver.Chrome(executable_path=r"driver2/chromedriver.exe",options=options)
    
    driver.set_page_load_timeout(60)
    #time.sleep(100)
    global sheet2
    global sheet
    global smart
    global dude
    global sheet_log
    
    try:
        indic=[]
        dude=dude+1
        dictt={}
        today = date.today()
        d4 = today.strftime("%m/%d/%y")
        print(d4)
        for xl in range(smart,dude+1,10):
            smart1=xl
            for n in range(1,100):
                l1=[]
                l2=[]
                l3=[]
                l4=[]
                
                dictt={}
                #smart1=xl
                try:
                    
                    for a in range(smart1,smart1+10):
                        #print('smart1 is'+str(a))
                        #print('dude')
                        if a>=dude:
                            break
                        #time.sleep(1)
                        k=sheet.acell("H"+str(a)).value
                        ab=sheet.acell("G"+str(a)).value
                        if 'Fedex'== k:
                            #print(sheet.acell("D"+str(a)).value)
                            l1.append(ab)
                            #l2.append(' ')
                            print('h1')
                            print(a)
                            dictt[a]=ab
                        elif 'fedex'== k:
                            #print(sheet.acell("D"+str(a)).value)
                            l1.append(ab)
                            #l2.append(' ')
                            print('h1')
                            print(a)
                            dictt[a]=ab
                        elif 'Blue Dart'== k:
                            #print(sheet.acell("D"+str(a)).value)
                            #l1.append(' ')
                            l2.append(ab)
                            print('h2')
                            print(a)
                            dictt[a]=ab
                        elif 'Bluedart'== k:
                            #print(sheet.acell("D"+str(a)).value)
                            print('h3')
                            #l1.append(' ')
                            l2.append(ab)
                            print(a)
                            dictt[a]=ab
                        else:
                            #l1.append(' ')
                            #l2.append(' ')
                            pass

                except gspread.exceptions.APIError  as jk:
                    time.sleep(10)
                    smart=a
                    

                else:
                    break

            print(dictt) 
            #print(jk)
            try:
                if len(l1)<2:
                    l1.append('1234exclu')
                else:
                    l1 = ' '.join(l1).split()
                #print(l1)
            except Exception as mar:
                print(mar)
            print("fedex count"+str(len(l1)))
            print("bluedart count"+str(len(l2)))
            l6=l2
            def get_key(val):
                for key, value in dictt.items():
                    if val == value:
                        return key
            
                return "key doesn't exist"

            for ite in range(1,30):
                try:
                    driver.get("https://www.fedex.com/en-in/tracking.html")
                    #ddrive_call=True
                    time.sleep(4)
                    
                    driver.find_element_by_link_text("MULTIPLE TRACKING NUMBERS").click()
                    time.sleep(2)
                except (WebDriverException,TimeoutException) as w:
                    print(w)
                    print("entered here")
                    time.sleep(3)
                    driver.refresh()
                else:
                    break
            print()
            p=1
            q=1
            

            
            if len(l1)<=1:
                l1.append('exitcode')
            for i in l1:
                driver.find_element_by_xpath("//*[@id='number']/div/form/div/div/form/div[2]/div["+str(p)+"]/div/input").send_keys(l1[q-1])
                p=p+1
                q=q+1

            driver.find_element_by_xpath("//*[@id='number']/div/form/div/div/form/div[3]/button").click()
            
            for j in range(1,10):
                try:
                    
                    time.sleep(10)
                    driver.find_element_by_class_name("va_icon")
                    break
                except Exception as g:
                    print(g)
                    driver.refresh()
                    time.sleep(10)
                else:
                    pass
            print()
            
            try:
                print(driver.find_element_by_xpath("//*[@id='container']/ng-component/h1").text)
            except Exception:
                pass
            else:
                try:
                    rem=driver.find_element_by_xpath("//*[@id='container']/ng-component/ul/li[2]").text
                    l1.remove(str(rem))
                    print(l1)
                    for ite in range(1,30):
                        try:
                            driver.get("https://www.fedex.com/en-in/tracking.html")
                            #ddrive_call=True
                            time.sleep(4)
                            
                            driver.find_element_by_link_text("MULTIPLE TRACKING NUMBERS").click()
                            time.sleep(2)
                        except (WebDriverException,TimeoutException) as w:
                            print(w)
                            print("entered here")
                            time.sleep(3)
                            driver.refresh()
                        else:
                            break
                    print()
                    p=1
                    q=1



                    if len(l1)<=1:
                        l1.append('exitcode')
                    for i in l1:
                        driver.find_element_by_xpath("//*[@id='number']/div/form/div/div/form/div[2]/div["+str(p)+"]/div/input").send_keys(l1[q-1])
                        p=p+1
                        q=q+1

                    driver.find_element_by_xpath("//*[@id='number']/div/form/div/div/form/div[3]/button").click()
                    
                    for j in range(1,10):
                        try:
                            
                            time.sleep(10)
                            driver.find_element_by_class_name("va_icon")
                            break
                        except Exception as g:
                            print(g)
                            driver.refresh()
                            time.sleep(10)
                        else:
                            pass
                    print()
                except Exception as pkl:
                    print(pkl)
                    print('some duplicates are found in the list while retreiving')
                    break

            #driver.get("https://www.fedex.com/fedextrack/summary?trknbr=772400631678,772400647173,772479832559,772400656796,772479528503")
            #driver.get("https://www.fedex.com/fedextrack/summary?trknbr=772650148154,772400647173")
            #time.sleep(10)
            i=1
            for n in range(1,30):
                
                try:
                    for ij in l1:
                        x_call=False
                        try:

                            x=driver.find_element_by_xpath("//*[@id='container']/app-summary-page/div[3]/trk-shared-stylesheet-wrapper/div/div[1]/table/tbody/tr["+str(i)+"]").text
                            x_call=True
                        except Exception:
                            pass
                        else:
                            pass
                        if x_call==True:
                        
                            x=x.split("\n")
                            #y=driver.find_element_by_xpath("//*[@id='container']/app-summary-page/div[3]/trk-shared-stylesheet-wrapper/div/div[1]/table/tbody/tr[2]").text
                            #y=y.split("\n")
                            print(x)
                            #print(y)
                            z=x[0].split(" ")
                            print(z)
                            print(z[0])
                            gett=z[0]
                            for ph in range(1,100):
                                try:
                                    ke=get_key(gett)
                                    sheet.update_acell('I'+str(ke), x[1])
                                    sheet.update_acell('J'+str(ke), x[2])
                                    sheet.update_acell('K'+str(ke), d4)
                                    time.sleep(1)
                                    if  x[1]=='Delivered':
                                        #print('a3 is'+str(a3))
                                        roww=sheet.row_values(ke)
                                        sheet2.insert_row(roww,2)
                                        indic.append(ke)
                                    else:
                                        pass
                                
                                except gspread.exceptions.APIError:
                                    time.sleep(10)
                                else:
                                    break
                            
                            

                            
                        else:
                            pass
                        i=i+1
                        time.sleep(1)
                    

                except gspread.exceptions.APIError  as jk:
                    time.sleep(10)
                    smart=a
                    

                else:
                    break
            
            for ite in range(1,30):
                try:
                    driver.get("https://www.bluedart.com/tracking")
                    time.sleep(4)
                    strinn= ','.join(l2).split()
                    driver.find_element_by_xpath("//*[@id='trackingNo']").send_keys(strinn)
                    time.sleep(3)
                    driver.find_element_by_xpath("//*[@id='goBtn']").click()
                    time.sleep(2)
                except (WebDriverException,TimeoutException) as w:
                    print(w)
                    print("entered here")
                    time.sleep(3)
                    driver.refresh()
                else:
                    break

            l4=[]

            for ij in l6:
                sp=str("'")+str(ij)+'-rdrmv'+str("'")
                
                #i=driver.find_element_by_xpath("//*[@id='69627651015-rdrmv']").text
                try:
                    try:
                        y_call=True
                        i=driver.find_element_by_xpath("//*[@id="+str(sp)+"]").text
                    except Exception:
                        pass
                    if y_call==True:
                        h=i.split("\n")
                        print(h)
                        #l4.append(h)
                        bd=get_key(str(ij))
                        for ph in range(1,100):
                            try:
                                sheet.update_acell('I'+str(bd), h[3])
                                sheet.update_acell('J'+str(bd), h[4])
                                sheet.update_acell('K'+str(bd), d4)
                                if  h[3]=='Shipment Delivered':
                                
                                    roww=sheet.row_values(bd)
                                    sheet2.insert_row(roww,2)
                                    indic.append(bd)
                                else:
                                    pass
                            except gspread.exceptions.APIError:
                                time.sleep(10)
                            else:
                                break
                                
                    else:
                        pass

                except Exception as b:
                    print(b)

        #print(indic)
        
        x=0
        indic.sort()
        print(indic)
        hsmac=indic
        for tu in range(1,100):
            try:
                indic=hsmac
                for i in indic:
                    
                    j=i-x
                    print("\n")
                    print(i)
                    print(x)
                    
                    sheet.delete_row(int(j))
                    print('deleted')
                    x=x+1
                    gsmart=indic.index(i)+1
                    print(gsmart)
            except gspread.exceptions.APIError  as jk:
                print('stopped here')
                time.sleep(10)
                hsmac=indic[gsmart:]
                print('hsmac is'+str(hsmac))
                

            else:
                break








        ####program completed ....
    except Exception as ph:
        print(ph)
        print("program was interrupted for some reason \n type rerun and hit enter to rerun the program \n type exit to close the terminal")
        driver.close()
        now = datetime.now()
        
        current_time = now.strftime("%I:%M:%p")
        print(current_time)
        
        newll=["shipment prog interrupted","start"+str(smart),"end"+str(dude),current_time]
        #global sheet_log
        #sheet_log=sheet_log
        sheet_log.insert_row(newll,4)
        print(current_time)

    else:
        print("\n\nsession completed ..!! \n type rerun and hit enter to rerun the program \n type exit to close the terminal")
        driver.close()
        now = datetime.now()
        
        current_time = now.strftime("%I:%M:%p")
        print(current_time)
        
        newll=["shipment completed","start"+str(smart),"end"+str(dude),current_time]
        
        #sheet_log=sheet_log
        sheet_log.insert_row(newll,4)
    nex=str(input())
    if nex=='rerun':

        startingg()
    elif nex=='exit':

        sys.exit()

def verifyy():
    sheet = gc.open("All information").worksheet('information')
    print("please verify your credentials to enter the session")

    print("User Id:")
    id =str(input())
    print("password:")
    secure_password_input()
    #pas =str(input())
    for i in range(1,12):
        if sheet.cell(i,1).value==id and sheet.cell(i,2).value==p_s:
            sheet_called=True

            print("verified")
            global sheet_log
            sheet_log = gc.open("All information").worksheet(id)
            newl=[current_time +'*',d3+'*']
            sheet_log.insert_row(newl,2)
            startingg()
    print()
    verify2()

def verify2():
    try:

        if sheet_called==True:
            pass
        
    except Exception:
        print("user id or password you entered is incorrect")
        verifyy()
    

verifyy()


