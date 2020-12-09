import pandas as pds
import xlrd
import whois
import pythonwhois

def file_loader():
    mail_file = open(file="\\hpserver\\Devsons  Data\\Accounts\\MIR THE CHOTLI\\email list 1(James).xlsx")
    return mail_file

def xl_read():
    read_file = pds.read_excel(".\\EMAILS.xlsx",sheet_name="Sheet2")
    i = 0
    for i in range(0,100):
        curr= read_file.iloc[i]
        print(curr)

def xlsx_read():
    loc = (".\\EMAILS.xlsx")

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    for i in range(0,100):#(i,0) for all
        mail = sheet.cell_value(i, 0)
        dom = mail.split("@")[1]
        print(dom)

def chck_dom(dom):
    info = whois.whois(dom)
    print(info)
    return info

def chck_dom2(dom):
    info = pythonwhois.get_whois(dom)
    print(info)
    res = {key: info[key] for key in info.keys() & {'contacts'}}
    print(res)
    res2 = get_registrar(res)
    return res

def inp_email():
    mail = str(input("Enter your email address: \n"))
    return mail

def get_registrar(res):
    #res2 = {key: res[key] for key in res.keys() & {'registrant'}}
    res2 = res.get('contacts',{}).get('registrant')
    res3 = {key: res2[key] for key in res2.keys() & {'country'}}
    print(res3)
    return res3

x = inp_email()
y = chck_dom2(x)
