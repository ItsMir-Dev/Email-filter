import pandas as pds
import xlrd
import whois
import pythonwhois
import csv
import logging
import datetime

#array that stores filtered emails
filtered_mails = []


def main(): #main method that runs script with excel file and loads results to csv file
    logging.basicConfig(filename='log.txt',level=logging.INFO)
    loc = ".\\Mailing List Updated.xlsx"

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    for i in range(0,10): #(i,0) for all
        mail = sheet.cell_value(i, 0)
        temp_mail = mail
        dom = mail.split("@")[1]
        try:
            pythonwhois.get_whois(dom)
            whois_count = chck_dom(dom)
            dom_2 = chck_dom2(dom)
            country = get_country(dom_2)
            if '.ke' in temp_mail or '.tz' in temp_mail or '.ug' in temp_mail:
                filtered_mails.append(temp_mail)
                add_log(temp_mail, country)
            elif whois_count == 'KE' or whois_count == 'UG' or whois_count == 'TZ':
                filtered_mails.append(temp_mail)
                add_log(temp_mail,whois_count)
            elif whois_count == 'None':
                if country == 'KE' or country == 'TZ' or country == 'UG':
                    filtered_mails.append(temp_mail)
                    add_log(temp_mail,country)
            else:
                print("Domain not from East Africa")
        except Exception:
            pass
    add_to_file(filtered_mails)


def file_loader(): #open any file
    mail_file = open(file="\\hpserver\\Devsons  Data\\Accounts\\MIR THE CHOTLI\\email list 1(James).xlsx")
    return mail_file


def xl_read(): #reads an excel file .xls, .xlsx
    read_file = pds.read_excel(".\\EMAILS.xlsx",sheet_name="Sheet2")
    i = 0
    for i in range(0,100):
        curr = read_file.iloc[i]
        print(curr)


def chck_dom(dom): #gets all domain information
    info = whois.whois(dom)
    print("Whois country:", info.country)
    return info.country


def chck_dom2(dom): #get all domain information, and filters to contact info
    try:
        info = pythonwhois.get_whois(dom)
        print("Pythonwhois %s info: ", dom, info)
        res = {key: info[key] for key in info.keys() & {'contacts'}}
        return res
    except Exception:
        print(dom + " domain doesnt exist")
        pass


def inp_email(): #allows user to input an email address. Use for testing
    mail = str(input("Enter your email address: \n"))
    return mail


def get_registrar(res): #obtains registrar details of domain
    #res2 = {key: res[key] for key in res.keys() & {'registrant'}}
    res2 = res.get('contacts',{}).get('registrant',{}).get('country',{})
    #res3 = {key: res2[key] for key in res2.keys() & {'country'}}
    print(res2)
    return res2


def get_country(dom_info): #obtains country of operation for domain
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


def add_to_file(email): #adds array of emails to csv file
    with open('emails.csv','w+') as fp:
        write = csv.writer(fp, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        write.writerow(email)


def add_log(dom,country):
    date_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_txt = "INFO: " + date_time + ": Checked " +dom + " and found active country:" + country + "\n"
    with open('log.txt','a') as lf:
        lf.write(log_txt)



main()
#y = str(input("Enter your domain name: \n"))
#chck_dom2(y)
#chck_dom(y)


