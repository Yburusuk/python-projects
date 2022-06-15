from os import link
from requests import session
import requests_html, re, lxml, cloudscraper
from bs4 import BeautifulSoup
import openpyxl, pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
def main():

    options = webdriver.ChromeOptions()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument("window-size=1280,720")
    driver = webdriver.Chrome(executable_path=ChromeDriverManager().install())

    tournaments = "https://www.atptour.com/en/scores/results-archive?year=2021"
    driver.get(tournaments)
    html_tournaments = driver.page_source
    soup_tournaments = BeautifulSoup(html_tournaments, 'lxml')

    links = soup_tournaments.find_all("a", class_="tourney-title")
    
    href_list = []
    for i in links:
        if not (("Cancelled" in i["data-ga-label"]) or ("ATP Cup" in i["data-ga-label"])
        or ("ATP Finals" in i["data-ga-label"]) or ("Laver Cup" in i["data-ga-label"])
        or ("Olympics" in i["data-ga-label"])):
            href_list.append(i["href"])

    links_ = []
    for i in href_list:
        val01 = "https://www.atptour.com" + i
        links_.append(val01)
    
    links_new = []
    for i in links_:
        val02 = i.replace("tournaments", "scores/archive")
        val03 = val02.replace("overview", "2021/draws")
        links_new.append(val03)
    
    driver.quit()

    mm = 0
    def code(url):
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument("window-size=1280,720")
        options.add_argument('--profile-directory=Default') 
        
        driver = webdriver.Chrome(executable_path=ChromeDriverManager().install())
        driver.get(url)
        html = driver.page_source
        soup = BeautifulSoup(html, 'lxml')
        
        sgl_data= soup.find_all("span", class_="item-value")
        sgl = int((sgl_data[0].text).strip())
        
        t_name = (soup.find("a", class_="tourney-title").text).strip()

        table = soup.find_all("a", class_= "scores-draw-entry-box-players-item")

        players = []
        for i in table:
            detail = i.text
            players.append(detail.strip())
        
        players_d = {}
        for i in range(len(players)):
            players_d[players[i]] = players.count(players[i])
        
        players_sorted = {}

        sorted_info = sorted(players_d, key = players_d.get)
        for i in sorted_info:
            players_sorted[i] = players_d[i]
        
        def Key(val):
            a = [m for m,n in players_sorted.items() if n == val]
            return a

        wl = ["Lost in R128", "Lost in R64", "Lost in R32", "Lost in R16",
            "Lost in Quarterfinal", "Lost in Semifinal", "Lost in Final", "Winner"]
        col = ["H","G","F","E","D","C","B","A"]

        results = {}
        columns = []
        n = 0
        if (16 < sgl < 33):
            n = 6
            while (n > 0):
                results[wl[n+1]] = Key(n)
                columns.append(col[n+1])
                n -= 1
        elif (32 < sgl < 65):
            n = 7
            while (n > 0):
                results[wl[n]] = Key(n)
                columns.append(col[n])
                n -= 1
        elif (64 < sgl < 129):
            n = 8
            while (n > 0):
                results[wl[n-1]] = Key(n)
                columns.append(col[n-1])
                n -= 1
        
        df = pd.DataFrame(dict([ (k,pd.Series(v)) for k,v in results.items() ]))
        df.fillna("", inplace=True)
        filem = f"{mm}.{t_name}.xlsx"
        df.to_excel(filem, sheet_name=t_name, index=False)

        wb = openpyxl.load_workbook(filename=filem)
        worksheet = wb.active

        r_key = [m for m in results.keys()]
        r_val = [i for i in (results.get(m) for m in r_key)]
        longest = [len(max(r_val[i], key=(len))) for i in range(len(r_key))]
        
        for m, n in zip(columns, longest):
            worksheet.column_dimensions[m].width = n+1

        wb.save(filem)

        driver.quit()
    

    for i in links_new:
        code(i)
        mm += 1
    
main()
