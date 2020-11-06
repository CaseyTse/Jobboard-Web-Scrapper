from bs4 import BeautifulSoup
import requests
import sys
import xlwt
from xlwt import Workbook
import datetime
#Programmer: Casey Tse
#Date: September 12,2020
#Purpose: Project is to further develop skills with scripting using python
#Project is going to be to scrape job postings on indeed, and import the data to an excel file
#The WS will extract the title, Company, Location, Experience, Description, Date into an excel file
#This will allow for easier and more consistent job postings to be seen, and can be further filtered by an excel file.
#The Job Scrapper will scrape all the data purely from indeed.com

#page_url = sys.argv[1]# using the command line can input the url
#

wb = Workbook()
sheet = wb.add_sheet("Sheet 1",cell_overwrite_ok = True)

sheet.write(0,0,"Title")
sheet.write(0,1,"Vote Count")
sheet.write(0,2,"Link")

front_page = "https://news.ycombinator.com/"
page_format = "https://news.ycombinator.com/news?p="

Article_list = []
Links_list = []
Popular_Articles = []
#takes the list full of dictionies and puts it into an excel file
def Add_To_Excel(list):
    for i in range(1,len(list)):
        print(list[i])
        sheet.write(i,0,list[i]["Title"])
        sheet.write(i,1,list[i]["Votes"])
        sheet.write(i,2,list[i]["Link"])


        x = datetime.datetime.now()
        year = x.strftime("%Y")
        month = x.strftime("%m")
        day = x.strftime("%d")
        date = (day + "_" + month + "_" + year)
        wb.save("Popular_Hacker_News_Articles_" + date +".xls")
#page_url = "https://news.ycombinator.com/"
def Scrape_Page(page_url):


    page = requests.get(page_url)
    print(page.status_code)
    if not page.status_code == 200:
        print("Page unsucessful")
        return(False)
    else:
        print("Page sucessfully accessed!")

        #get the page contents
        soup = BeautifulSoup(page.content,'html.parser')

        #Create a list of all the articles found on the page.
        for article in (soup.find_all(class_="storylink")):
            Article_list.append(article.get_text())
            Links_list.append(article.get("href"))
        #loop through all articles and get the scores.
        for index,posting in enumerate(soup.find_all(class_="subtext")):

            #Check if  the posting has any score to it, and clean it up
            if posting.find(class_="score"):
                votes = posting.find(class_="score")
                points = votes.get_text()
                cleaned_points = points.split(" ")#Extract the number of points
                if int(cleaned_points[0]) > 100:# see if the article has more than 100 points
                    Popular_Articles.append({"Title": Article_list[index],"Votes": cleaned_points[0],"Link":Links_list[index]})
            Add_To_Excel(Popular_Articles)





if __name__ ==  "__main__":
    Scrape_Page(front_page)
    # for index in range(10):
    #     page_url = page_format + str(index)
    #     Scrape_Page(page_url)
