from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import Select
import pandas as pd
import time
import csv
import os
import re
import json
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import html5lib
import random

# ランダムな待機時間を追加する関数
def add_random_wait():
    # ランダムな待機時間を生成（1から5秒の間でランダム）
    wait_time = random.uniform(5, 10)
    print(f"Waiting for {wait_time} seconds...")
    # 待機時間分スリープ
    time.sleep(wait_time)

def extract_data(soup):
    data = {}
    def find_data_by_th(th_string):
        th_element = soup.find('th', string=th_string)
        return th_element.find_next_sibling('td').get_text(strip=True) if th_element else None
    data['登録番号'] = find_data_by_th('登録番号')
    data['最初の登録年月日'] = find_data_by_th('最初の登録年月日')
    th_tags = soup.find_all('th')
    for th in th_tags:
        flat_text = ''.join(th.stripped_strings)
        if '有効期間' in flat_text and '起算日' in flat_text:
            data['有効期間（起算日）'] = th.find_next_sibling('td').get_text(strip=True) if th.find_next_sibling('td') else ""
        elif '有効期間' in flat_text and '満了日' in flat_text:
            data['有効期間（満了日）'] = th.find_next_sibling('td').get_text(strip=True) if th.find_next_sibling('td') else ""
        elif '主たる事務所の所在地' in flat_text:
            data['主たる事務所の所在地'] = th.find_next_sibling('td').get_text(strip=True) if th.find_next_sibling('td') else ""
    data['法人・個人の別'] = find_data_by_th('法人・個人の別')
    phone_number = find_data_by_th('電話番号')
    if phone_number and not '-' in phone_number:
        if phone_number.startswith('011'):
            phone_number = phone_number[:3] + '-' + phone_number[3:6] + '-' + phone_number[6:]
        elif len(phone_number) == 10:
            phone_number = phone_number[:2] + '-' + phone_number[2:6] + '-' + phone_number[6:]
    data['電話番号'] = phone_number
    company_name_th = soup.find('th', string='商号又は名称')
    if company_name_th:
        company_name_td = company_name_th.find_next_sibling('td')
        data['商号又は名称フリガナ'] = company_name_td.find('p', class_='phonetic').get_text(strip=True)
        data['商号又は名称'] = company_name_td.contents[-1].strip() 
    representative_name_th = soup.find('th', string='代表者の氏名')
    if representative_name_th:
        representative_name_td = representative_name_th.find_next_sibling('td')
        data['代表者の氏名フリガナ'] = representative_name_td.find('p', class_='phonetic').get_text(strip=True)
        data['代表者の氏名'] = representative_name_td.contents[-1].strip()
    for key, value in data.items():
        if value:
            data[key] = value.replace('\u3000', '　')
    return data

def extract_office_data(soup2, office_name):
    office_data = []
    rows = soup2.select('table.re_summ_sc2 tr')
    for row in rows:
        columns = row.find_all(['th', 'td'])
        row_data = [col.get_text(strip=True) for col in columns]
        if any(row_data) and row_data[1] == office_name:
            address_cell = row.find('td', style='width : 272px ;')
            address_text = address_cell.get_text(strip=True, separator=' ')
            postal_code, prefecture, city, other_address = split_address(address_text)
            phone_number = row_data[4]
            if phone_number and not '-' in phone_number:
                if phone_number.startswith('011'):
                    phone_number = phone_number[:3] + '-' + phone_number[3:6] + '-' + phone_number[6:]
                elif len(phone_number) == 10:
                    phone_number = phone_number[:2] + '-' + phone_number[2:6] + '-' + phone_number[6:]
            row_data[4] = phone_number
            row_data[3] = postal_code
            row_data.insert(4, prefecture)
            row_data.insert(5, city)
            row_data.insert(6, other_address)
            office_data.append(row_data)
    return office_data

PREFECTURES = [
    '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県', '茨城県',
    '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県', '新潟県', '富山県',
    '石川県', '福井県', '山梨県', '長野県', '岐阜県', '静岡県', '愛知県', '三重県',
    '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県', '鳥取県', '島根県',
    '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県', '福岡県',
    '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'
]

def split_address(address):
    # 郵便番号を取得
    postal_code_match = re.match(r'〒(\d{3}-\d{4})', address)
    postal_code = postal_code_match.group(1) if postal_code_match else None
    address = address[len(postal_code)+1:] if postal_code else address
    address = address.strip()# 先頭と末尾の空白を削除
    # 奈良県大和郡山市の住所の判定
    if address.startswith("奈良県大和郡山市"):
        prefecture = "奈良県"
        city = "大和郡山市"
        other_address = address.replace("奈良県大和郡山市", "", 1).strip()
        return postal_code, prefecture, city, other_address
    else:
        # 都道府県を取得
        prefecture = next((p for p in PREFECTURES if p in address), None)
        address = address.replace(prefecture, '', 1).strip() if prefecture else address
        # 他の特定の市区郡名を検索
        special_cities = ("市川市", "市原市", "野々市市", "四日市市", "廿日市市", "余市郡", "高市郡", "郡山市", "郡上市", "蒲郡市", "小郡市")
        city = next((sc for sc in special_cities if sc in address), None)
        # 特定の市区郡名が見つからなければ一般的な市区郡名を検索
        if not city:
            city_match = re.search(r'(.+?[市区郡])', address)
            city = city_match.group(1) if city_match else None
        other_address = address[len(city):].strip() if city else address
        return postal_code, prefecture, city, other_address

# Pandas DataFrameの初期化
data_columns = [
    "番号", "登録番号", "最初の登録年月日", "有効期間（起算日）", "有効期間（満了日）", "法人・個人の別",
    "商号又は名称フリガナ", "商号又は名称", "代表者の氏名フリガナ", "代表者の氏名",
    "主たる事務所の所在地", "電話番号", "No.", "名称", "事務所の区分", "所在地郵便番号",
    "所在地都道府県", "所在地市区郡", "所在地その他", "事務所電話番号"
]
all_data_df = pd.DataFrame(columns=data_columns)
caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "eager"
options = webdriver.chrome.options.Options()

# User-Agentの設定を追加
#user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36' 
user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36' 
options.add_argument(f'user-agent={user_agent}')

profile_path = os.path.join(os.path.expanduser('~'), 'Library', 'Application Support', 'Google', 'Chrome', 'Default')
options.add_argument('--user-data-dir=' + profile_path)
options.add_argument('--ignore-certificate-errors')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--disable-popup-blocking")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
options.add_experimental_option("prefs", {
    "profile.managed_default_content_settings.images": 2  # 画像の読み込みの無効化
})
options.add_argument('--headless')
options.add_argument('window-size=1920x1080')
service = ChromeService(ChromeDriverManager().install())
#driver = webdriver.Chrome(desired_capabilities=caps, service=service, options=options)
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver,60)

chintai = "https://etsuran2.mlit.go.jp/TAKKEN/chintaiKensaku.do"
driver.get(chintai)

soup = BeautifulSoup(driver.page_source, 'lxml')
select = soup.find('select', id='kenCode')
options = select.find_all('option')
prefectures = [opt['value'] for opt in options if opt['value']]
START_PAGE = 1
END_PAGE = 100
page_num = START_PAGE

try:
    for pref in [prefectures[7]]:
        # 都道府県のプルダウンを選択
        dropdown = driver.find_element(By.ID,'kenCode')
        dropdown.click()
        dropdown.find_element(By.CSS_SELECTOR,f'option[value="{pref}"]').click()
        # 本店のプルダウンは選択しない→一覧の事務所名のみを事務所タブで取得する
        # 降順を選択
        driver.find_element(By.ID, 'rdoSelectKo').click()
        # 検索表示50件
        disp_count_dropdown = driver.find_element(By.ID, 'dispCount')
        disp_count_dropdown.click()
        option_50 = disp_count_dropdown.find_element(By.CSS_SELECTOR, 'option[value="50"]')
        option_50.click()
        # 検索ボタンをクリック
        search_btn = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'img[onclick="js_Search(0)"]')))
        search_btn.click()
        # ページ遷移のドロップダウンメニューを選択
        page_dropdown = driver.find_element(By.ID, 'pageListNo1')
        page_dropdown.click()
        page_option = page_dropdown.find_element(By.CSS_SELECTOR, f'option[value="{START_PAGE}"]')
        page_option.click()
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'img[onclick="js_Search(0)"]')))
        while True:
            try:
                # ランダムな待機時間を追加
                add_random_wait()
                
                links_elements = driver.find_elements(By.CSS_SELECTOR, 'td a[onclick^="js_ShowDetail"]')
                office_names_elements = driver.find_elements(By.CSS_SELECTOR, 'td[style="text-align:left; white-space : nowrap;"]:nth-child(5)')
                links = [link.get_attribute('onclick') for link in links_elements]
                office_names = [office.text.strip() for office in office_names_elements]
                for idx, link_onclick in enumerate(links):
                    office_name = office_names[idx]
                    #print(office_name)
                    no_element = driver.find_elements(By.CSS_SELECTOR, 'tr:not(.trev) td[style="text-align:right;"]')[idx]
                    data_value = no_element.text.strip() if no_element else ""
                    data = {"番号": data_value}
                    print(data)
                    link_element = driver.find_element(By.CSS_SELECTOR, f'td a[onclick="{link_onclick}"]')
                    link_element.click()
                    houjin = wait.until(EC.presence_of_element_located((By.XPATH, '//th[text()="法人・個人の別"]')))
                    soup = BeautifulSoup(driver.page_source, 'html.parser')
                    data.update(extract_data(soup))
                    #print(json.dumps(data, ensure_ascii=False, indent=2))
                    #print(data)
                    current_url = driver.current_url
                    element = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@onclick, 'jimusyo_submit();')]")))#事務所ページ
                    element.click()
                    wait.until(lambda d: d.current_url != current_url)
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.re_summ_sc2")))
                    soup2 = BeautifulSoup(driver.page_source, 'html.parser')
                    office_data_list = extract_office_data(soup2, office_name)
                    #print(office_data_list)
                    driver.back()
                    main_data_row = pd.Series({header: data.get(header, "") for header in data_columns[:12]}) # 最初の12項目
                    for office_data in office_data_list:
                        if len(office_data) == 8:
                            office_data_series = pd.Series(office_data, index=data_columns[12:])
                            all_data_row = pd.concat([main_data_row, office_data_series], axis=0)
                            all_data_row = all_data_row.reindex(data_columns)  # 列名を data_columns に合わせる
                            all_data_df = pd.concat([all_data_df, all_data_row.to_frame().T], ignore_index=True)
                        else:
                            print(f"Error: Office data length mismatch for {office_name}")
                    driver.back() 
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'img[onclick="js_Search(0)"]')))#検索ボタンがある一覧ページ戻るまで待機
                    # 30秒待機
                    time.sleep(30)                
                if page_num > END_PAGE:
                    break  # 終了ページを超えたらループを終了
                next_page = driver.find_elements(By.XPATH, '//img[@onclick="js_Search(\'2\')"]')
                if not next_page:
                    break
                current_page_number = driver.find_element(By.ID, 'pageListNo1').get_attribute('value')
                next_page[0].click()
                wait.until(lambda driver: driver.find_element(By.ID, 'pageListNo1').get_attribute('value') != current_page_number)
                page_num += 1
                print(page_num)
            except TimeoutException as e:
                print(f"Timeout occurred at page {page_num}: {e}")
                break
    all_data_df.to_csv("chintai_茨城_202407.csv", index=False, encoding='utf-8-sig')
    print("CSVファイルに出力しました。")
except Exception as e:
    print(f"An error occurred: {e}")
    # エラー発生時に現時点でのデータをCSVに出力
    all_data_df.to_csv("chintai_partial.csv", index=False, encoding='utf-8-sig')
    print("エラー発生時のデータをCSVファイルに出力しました。")
print("おわり")