import subprocess
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import re
import pandas as pd

# User agent added to avoid 'enable cookies' error
headers = {"user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 12_4) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/102.0.5005.63 Safari/537.36"}
todaysDate = datetime.now()
todaysDateFormatted = todaysDate.strftime("%B %d, %Y")  # Looks like 'June 06, 2022'

# Empty lists for content, links, titles, dates
list_titles = []
list_desc = []
list_links = []
list_dates = []
list_sources = []

# Hacker News variables
thn = "https://thehackernews.com/"
thnReq = requests.get(thn, headers=headers)
thnPageText = thnReq.content
thnSoup = BeautifulSoup(thnPageText, "html.parser")
thn_frontpage = thnSoup.find_all("div", class_="body-post clear")  # identify only the news articles

# Bleeping Computer variables
bc = "https://www.bleepingcomputer.com/"
bcReq = requests.get(bc, headers=headers)
bcPageText = bcReq.content
bcSoup = BeautifulSoup(bcPageText, "html.parser")
bc_frontpage = bcSoup.find_all("div", class_="bc_latest_news_text")

# Scraping Bleeping Computer's front page for today's articles
for n in bc_frontpage:
    titles = n.find("h4").get_text()
    articleHeader = n.find("h4") # container for article links
    links = articleHeader.find("a")["href"]
    desc = n.find("p").get_text()
    source = "Bleeping Computer"
    dates = n.find("li", class_="bc_news_date").get_text()

    if dates == todaysDateFormatted:  # Appends output to list for Excel formatting
        list_dates.append(dates)
        list_titles.append(titles)
        list_links.append(links)
        list_sources.append(source)
        list_desc.append(desc)

# Scraping The Hacker News' front page for today's articles
for n in thn_frontpage:
    dates = str(n.find_all("div", attrs={"class": "item-label"}))
    datesFormatted = str(re.findall("([A-Za-z]+\s\d{2},\s\d{4})", dates)).translate({ord(c): None for c in "[]\'"})
    titles = n.find("h2", attrs={"class": "home-title"}).get_text()
    links = n.find("a")["href"]
    source = "The Hacker News"
    desc = n.find("div", attrs={"class": "home-desc"}).get_text().replace('\xa0','')
    shortenedDesc = str(desc.split('. ')[0]).strip('"\xa0"[]')  # just grabs the first part, otherwise too long

    if datesFormatted == todaysDateFormatted:
        list_dates.append(datesFormatted)
        list_titles.append(titles)
        list_links.append(links)
        list_sources.append(source)
        list_desc.append(shortenedDesc)


df = pd.DataFrame(  # Creates pandas dataframe to output 'columns:list values'
    {"Article Title": list_titles,
     "Date": list_dates,
     "Source": list_sources,
     "Description": list_desc,
     "Article Link": list_links}, )

# Formats and finds save location for Excel spreadsheet
writer = pd.ExcelWriter(r"C:\temp\NewsScraper.xlsx",
                        engine="xlsxwriter")
df.to_excel(writer, sheet_name="Today\'s News", index=False, na_rep="NaN")

# Get the xlsxwriter workbook and worksheet objects in order to add formatting
workbook = writer.book
worksheet = writer.sheets["Today\'s News"]

# Define background and formatting for headers
header_format = workbook.add_format({
    'bold': True,
    'fg_color': '#c6d9ec',
    'border': 1})

# Apply header formatting
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)

# Auto-adjusts columns' width
for column in df:
    column_width = max(df[column].astype(str).map(len).max(), len(column))
    col_idx = df.columns.get_loc(column)
    writer.sheets["Today\'s News"].set_column(col_idx, col_idx, column_width)

# Wrap the title and description columns
wrap = workbook.add_format({'text_wrap': True})
worksheet.set_column('D:D', 100, wrap)
worksheet.set_column('A:A', 65, wrap)

writer.save()

# Opens file
subprocess.Popen(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE NewsScraper.xlsx")
