import requests
from bs4 import BeautifulSoup
import datetime
import pandas as pd

today = datetime.datetime.today()
file_date = date = today.strftime("%Y%m%d")

keywords = ["コーヒー", "珈琲", "coffee"]
nikkei_url = "https://www.nikkei.com/"

class NikkeiSearch:

    def __init__(self):
        self.articles = []
        self.title_list = []
        self.desc_list = []
        self.url_list = []

    def search(self, name):
        response = requests.get(f"{nikkei_url}search?keyword={name}&volume=20")
        response_html = response.text
        soup = BeautifulSoup(response_html, "html.parser")
        titles = soup.find_all(name="h3", class_="nui-card__title")
        descriptions = soup.find_all(name="a", class_="nui-card__excerpt")
        self.title_list = [title.text.replace("　", "").replace(" ", "").replace("\n", "") for title in titles]
        self.url_list = [desc.get_attribute_list("href")[0] for desc in descriptions]
        self.desc_list = [desc.text.replace("　", "").replace(" ", "").replace("\n", "") for desc in descriptions]

    def format_data(self):
        for n in range(len(self.title_list)):
            article = {}
            article["title"] = self.title_list[n]
            article["description"] = self.desc_list[n]
            article["url"] = self.url_list[n]
            self.articles.append(article)

    def export_data(self):
        df = pd.DataFrame(self.articles)
        df.duplicated(subset=["title"])
        df.to_excel(f"{file_date}_nikkei_coffee.xlsx", encoding='utf_8_sig', index=False)


nikkei = NikkeiSearch()

for keyword in keywords:
    nikkei.search(keyword)
    nikkei.format_data()
nikkei.export_data()


