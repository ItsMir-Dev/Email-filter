import pandas as pds
import xlrd
import whois
import pythonwhois
import csv

filtered_mails = []

def main():
    loc = ".\\Mailing List Updated.xlsx"

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    for i in range(0,100): #(i,0) for all
        mail = sheet.cell_value(i, 0)
        temp_mail = mail
        dom = mail.split("@")[1]
        try:
            pythonwhois.get_whois(dom)
            if '.ke' in temp_mail or '.ug' in temp_mail or '.tz' in temp_mail:
                filtered_mails.append(temp_mail)
            else:
                dom = mail.split("@")[1]
                dom_2 = chck_dom2(dom)
                country = get_country(dom_2)
                print(country)
                if country == 'KE' or country == 'TZ' or country == 'UG':
                    filtered_mails.append(temp_mail)
        except Exception:
            pass
    add_to_file(filtered_mails)

def file_loader():
    mail_file = open(file="\\hpserver\\Devsons  Data\\Accounts\\MIR THE CHOTLI\\email list 1(James).xlsx")
    return mail_file


def xl_read():
    read_file = pds.read_excel(".\\EMAILS.xlsx",sheet_name="Sheet2")
    i = 0
    for i in range(0,100):
        curr = read_file.iloc[i]
        print(curr)


def chck_dom(dom):
    info = whois.whois(dom)
    print(info)
    return info


def chck_dom2(dom):
    try:
        info = pythonwhois.get_whois(dom)
        print("All info: ", info)
        res = {key: info[key] for key in info.keys() & {'contacts'}}
        return res
    except Exception:
        pass


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
    try:
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
                    return dom_country
            else:
                print("\nRegistrant not found")
        else:
            print("\nBool False")
    except Exception:
        pass


def add_to_file(email):
    with open('emails.csv','w+') as fp:
        write = csv.writer(fp, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        write.writerow(email)


main()