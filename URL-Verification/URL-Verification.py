# -*- coding: utf-8-*-
import tldextract
import xlwings as xw
# f = open(r'C:\Users\Administrator\PycharmProjects\tjwxb\Web_Check\url.txt', "r")
# for line in f:
#     try:
#         url = tldextract.extract(line)
#         subdomain = "{0}.{1}.{2}".format(url.subdomain, url.domain, url.suffix)
#         domain = "{0}.{1}".format(url.domain, url.suffix)
#         print subdomain
#         print domain
#     except:
#         print "error"
# f.close()

app_xlsx = xw.App(visible=True,add_book=False)
wb = app_xlsx.books.open(r'C:\Users\Administrator\Desktop\URL.xlsx')
lists = wb.sheets[2].range('H2:H1359').value
# i = 2
# for line in lists:
#     try:
#         url = tldextract.extract(line)
#         if url.subdomain:
#             subdomain = "{0}.{1}.{2}".format(url.subdomain, url.domain, url.suffix)
#         else:
#             subdomain = "{0}.{1}".format(url.domain, url.suffix)
#         print subdomain
#         domain = "{0}.{1}".format(url.domain, url.suffix)
#         print domain
#         # domain = url.registered_domain
#         # print domain
#         wb.sheets[2].range('I'+ str(i)).value = subdomain
#         wb.sheets[2].range('J' + str(i)).value = domain
#     except:
#         print "error"
#     i += 1
#     print i

i = 2
for line in lists:
    try:
        url = tldextract.extract(line)
        if url.subdomain:
            if url.subdomain == 'www':
                subdomain = "{0}.{1}".format(url.domain, url.suffix)
            else:
                subdomain = "{0}.{1}.{2}".format(url.subdomain, url.domain, url.suffix)
        elif not url.suffix :
            subdomain = "{0}".format(url.domain)
        else:
            subdomain = "{0}.{1}".format(url.domain, url.suffix)
        # print subdomain
        wb.sheets[2].range('I'+ str(i)).value = subdomain
    except:
        print "error"
    i += 1
    # print i


wb.save(r'd:\URL.xlsx')
wb.close()
app_xlsx.quit()