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
    print("All info: ", info)
    res = {key: info[key] for key in info.keys() & {'contacts'}}
    #res2 = get_registrar(res)
    return res

def inp_email():
    mail = str(input("Enter your email address: \n"))
    return mail

def get_registrar(res):
    #res2 = {key: res[key] for key in res.keys() & {'registrant'}}
    res2 = res.get('contacts',{}).get('registrant',{}).get('country',{})
    #res3 = {key: res2[key] for key in res2.keys() & {'country'}}
    print(res2)
    return res2

def get_country(dom_info):
    if 'contacts' in y:
        # print("\nContacts: ", y)
        if "registrant" in y.get('contacts', {}):
            # print("\nRegistrant: ", y.get('contacts',{}))
            reg_det = y.get('contacts', {})
            reg_dat = {key: reg_det[key] for key in reg_det.keys() & {'registrant'}}
            # print("trial: ", reg_dat.get('registrant'))
            if reg_dat.get('registrant') == None:
                print("\nCountry not found")
            else:
                # print("\nCountry found: ")
                dom_country = str(reg_dat.get('registrant').get('country'))
                print("Domain's Active Country:", dom_country)
        else:
            print("\nRegistrant not found")
    else:
        print("\nBool False")


x = inp_email()
y = chck_dom2(x)


