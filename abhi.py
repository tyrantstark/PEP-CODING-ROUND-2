import requests#for requesting  html
from bs4 import BeautifulSoup #for parsing html
import xlsxwriter#for maintaining excel sheet
def href_generator(href):#return href of college
    href_start_index = href.index("href=") + 6
    href_end_index = href_start_index + (href[href_start_index:].find(">")-1)
    params = href[href_start_index:href_end_index]
    return params

college_duniya_base_url = "https://collegedunia.com"# base url of college_duniya
#url of college_dunia b.tech  Ajmer city
college_dunia_url ="https://collegedunia.com/btech/hisar-colleges"
#USING REQUEST MODULE FOR REQUESTING PAGE
response = requests.get(url=college_dunia_url,headers={"User-Agent":"Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36",
                                              "Accept-Language":"en-US,en;q=0.9"})
page_html = response.text
#USING BEAUTIFULSOUP FOR PARSING
soup = BeautifulSoup(page_html,"html.parser")
top_college = list(soup.find_all("div",class_="jsx-765939686 clg-name-address"))[:6]#ONLY FETCHING TOP 6 COLLEGE FROM COLLEGE DUNIYA
top_6 = {}#AN ENMPTY DICT FOR STROING HREF OF TOP 6 COLLGE KEY AS COLLEGE NAME : HREF AS VALUES
for top in top_college:
    name = top.get_text()
    params = href_generator(str(top))
    top_6[name]=params
#FINGING DATA FROM EACH COLLEGE WEBSITE
for college in top_6.keys():
    response = requests.get(url = f"{college_duniya_base_url}{top_6[college]}",headers={"User-Agent":"Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36","Accept-Language":"en-US,en;q=0.9"})
    page_data = response.text
    soup = BeautifulSoup(page_data,"html.parser")
    try:#FETCHING SUMMARY
        summry = soup.find("div",class_="jsx-3399697518 jsx-1406691586 article-body").find("p").get_text()
    except:
        summry = "No summry on website"
    all_tuples = []#FOR STORING ALL ENTRY IN LIST
    schema = ["No Schema on website","NULL","NULL"]#IN CASE OF TABLE NOT PRESENT ON WEBSITE
    try:
        table_data = soup.find("table",class_="jsx-2675951502 table table-striped text-center")
        scherma_names = table_data.find_all("th",class_="jsx-2675951502")
        schema = [attribute.get_text() for attribute in scherma_names]#GETTING THE SCHEMA OF TABLE
        tuples_data = table_data.find_all("td",class_="jsx-2675951502")#GETTING ALL ROWS
        all_value = [entry.get_text() for entry in tuples_data]#STORING ALL VALUE PRESENT IN TABLE
        for saperator in range(0,len(all_value),3):
            all_tuples.append(all_value[saperator:saperator+3])#MAKING 2D ARRAY BY SCHEMA SIZE SUBARRAY
    except:
        all_tuples.append(["No Table on website","NULL","NULL"])#IN CASE OF TABLE NOT PRESENT ON WEBSITE
    workbook = xlsxwriter.Workbook(f'{college}.xlsx')#EXCEL FILE MANAGER
    worksheet = workbook.add_worksheet()
    try:
        init=2
        worksheet.write('A1', "Summary")
        worksheet.write('A2',summry)
        worksheet.write('B1', schema[0])
        worksheet.write('C1', schema[1])
        worksheet.write('D1', schema[2])

    except:
        pass
    for row_no in range(len(all_tuples)):#LOOP FOR DOING ENTRY IN EXCEL SHEET
        try:
            worksheet.write(f'B{init}', all_tuples[row_no][0])
            worksheet.write(f'C{init}', all_tuples[row_no][1])
            worksheet.write(f'D{init}', all_tuples[row_no][2])
            init+=1
        except:
            init+=1
            pass
    workbook.close()
    #DONE
