import requests
from bs4 import BeautifulSoup
import datetime
import pandas as pd

today = datetime.datetime.today()
file_date = date = today.strftime("%Y%m%d")

keywords = ["コーヒー", "珈琲", "coffee"]
nikkei_url = "https://www.nikkei.com/search?keyword="

class NikkeiSearch:

    def __init__(self):
        self.articles = []
        self.title_list = []
        self.desc_list = []

    def search(self, name):
        response = requests.get(nikkei_url + f"{name}&volume=20")
        response_html = response.text
        soup = BeautifulSoup(response_html, "html.parser")
        titles = soup.find_all(name="h3", class_="nui-card__title")
        descriptions = soup.find_all(name="a", class_="nui-card__excerpt")
        title_list = [title.text.replace("　", "").replace(" ", "").replace("\n", "") for title in titles]
        url_list = [desc.get_attribute_list("href") for desc in descriptions]
        desc_list = [desc.text.replace("　", "").replace(" ", "").replace("\n", "") for desc in descriptions]
        return title_list, url_list, desc_list

    def export_data(self, title, url, desc):
        for n in range(len(title)):
            article = {}
            article["title"] = title[n]
            article["description"] = desc[n]
            article["url"] = url[n]
            self.articles.append(article)
        df = pd.DataFrame(self.articles)
        df.duplicated(subset=["title"])
        df.to_excel(f"{file_date}_nikkei_tenki.xlsx", encoding='utf_8_sig', index=False)

nikkei = NikkeiSearch()

for keyword in keywords:
    title_list, url_list, desc_list = nikkei.search(keyword)
    nikkei.export_data(title_list, url_list, desc_list)


