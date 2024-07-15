import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import openpyxl
import time
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter


def get_html_links(base_url):
    response = requests.get(base_url)
    response.raise_for_status()

    soup = BeautifulSoup(response.content, 'html.parser')

    html_links = []
    for link in soup.find_all('a', href=True):
        href = link['href']
        if href.endswith('.html'):
            full_url = "https://www.data.jma.go.jp" + href
            html_links.append(full_url)

    return html_links


def filter_links_by_date(html_links, start_date, end_date):
    filtered_links = []
    for url in html_links:
        try:
            date_str = url.split('/')[-1][2:10]
            date_obj = datetime.strptime(date_str, '%Y%m%d')
            if start_date <= date_obj <= end_date:
                unique_part = url.split('/')[-1][2:]
                corrected_url = f"https://www.data.jma.go.jp/vois/data/tokyo/STOCK/volinfo/VG{unique_part}"
                filtered_links.append(corrected_url)
        except ValueError:
            continue

    return filtered_links


def extract_info_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    pre_tag = soup.find('pre')
    if not pre_tag:
        return None

    content = pre_tag.get_text()
    lines = content.split('\n')
    info = {}

    info['第2報'] = '*' if '第2報' in content else ''

    for line in lines:
        if line.startswith('日　　時：'):
            datetime_str = line.split('：')[1].split('（')[0].strip()
            jst = datetime.strptime(datetime_str, '%Y年%m月%d日%H時%M分')
            utc = jst - timedelta(hours=9)
            info['噴火日時（日本時間）'] = jst.strftime('%Y/%m/%d %H:%M')
            info['噴火日時（UTC）'] = utc.strftime('%Y/%m/%d %H:%M')
        elif line.startswith('現　　象：'):
            phenomenon = line.split('：')[1].strip()
            if phenomenon in ["噴火", "爆発", "噴火したもよう", "停止したもよう", "連続噴火停止", "連続噴火継続"]:
                info['現象'] = phenomenon
            else:
                info['現象'] = "？"
        elif line.startswith('有色噴煙：'):
            colored_smoke = line.split('：')[1].strip()
            if "火口上" in colored_smoke:
                info['有色噴煙（1）'] = "火口上"
                info['有色噴煙（2）'] = colored_smoke.split('火口上')[1].split('m')[0].strip()
            elif "不明" in colored_smoke:
                info['有色噴煙（1）'] = "不明"
                info['有色噴煙（2）'] = ""
            elif colored_smoke == "":
                info['有色噴煙（1）'] = ""
                info['有色噴煙（2）'] = ""
            else:
                info['有色噴煙（1）'] = "？"
                info['有色噴煙（2）'] = "？"
        elif line.startswith('白色噴煙：'):
            white_smoke = line.split('：')[1].strip()
            if "火口上" in white_smoke:
                info['白色噴煙（1）'] = "火口上"
                info['白色噴煙（2）'] = white_smoke.split('火口上')[1].split('m')[0].strip()
            elif "不明" in white_smoke:
                info['白色噴煙（1）'] = "不明"
                info['白色噴煙（2）'] = ""
            elif white_smoke == "":
                info['白色噴煙（1）'] = ""
                info['白色噴煙（2）'] = ""
            else:
                info['白色噴煙（1）'] = "？"
                info['白色噴煙（2）'] = "？"
        elif line.startswith('流　　向：'):
            direction = line.split('：')[1].strip()
            info['流向'] = direction[:3]

    return info


def save_to_excel(data, filename):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "噴火情報"

    headers = ["噴火日時（日本時間）", "噴火日時（UTC）", "現象", "有色噴煙（1）", "有色噴煙（2）", "白色噴煙（1）",
               "白色噴煙（2）", "流向", "第2報"]
    sheet.append(headers)

    for entry in data:
        row = [entry.get(header, "") for header in headers]
        sheet.append(row)

    workbook.save(filename)


def main():
    base_url = "https://www.data.jma.go.jp/vois/data/tokyo/STOCK/volinfo/volinfo.php?info=VG&id=506"
    html_links = get_html_links(base_url)

    start_date = datetime(2015, 7, 1)
    end_date = datetime(2024, 3, 31)

    filtered_links = filter_links_by_date(html_links, start_date, end_date)

    if not filtered_links:
        print("その範囲にHTMLファイルなし")
        return

    filtered_links.reverse()

    session = requests.Session()
    retries = Retry(total=5, backoff_factor=1, status_forcelist=[502, 503, 504, 500])
    session.mount('https://', HTTPAdapter(max_retries=retries))

    data = []
    for url in filtered_links:
        try:
            response = session.get(url)
            response.raise_for_status()
            info = extract_info_from_html(response.content)
            if info:
                data.append(info)
            time.sleep(1)  # リクエスト間に1秒の遅延を追加
        except requests.exceptions.RequestException as e:
            print(f"Error fetching {url}: {e}")

    save_to_excel(data, 'Hunka_info.xlsx')
    print(f"'Hunka_info.xlsx'に保存完了")


if __name__ == "__main__":
    main()
