import requests
import csv
from lxml import etree

host = "https://scholar.google.ca"
headers = {
    "user-agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36"
}

start_url = "https://scholar.google.ca/citations?hl=en&view_op=search_authors&mauthors=Universit%C3%A9+Laval&btnG="


def test(start_url):
    res = requests.get(start_url, headers=headers)
    selector = etree.HTML(res.text)
    for item in selector.xpath('//div[@class="gsc_1usr"]'):
        data = dict()
        user_url = host + item.xpath(".//a[not(@class)]")[0].attrib["href"]
        name = item.xpath("string(.//a[not(@class)])")
        title = item.xpath('string(.//div[@class="gs_ai_aff"])')
        img_url = host + item.xpath('.//img')[0].attrib["src"]

        data["name"] = name
        data["title"] = title
        data["img_url"] = img_url
        data.update(user_detail(user_url))
        print(data)
        save_data(data)
    next_page = selector.xpath('//button[@aria-label="Next"]')
    if next_page:
        print(host +
              next_page[0].attrib["onclick"].replace("window.location=\'", "").
              replace("\'", "").replace(r"\x3d", "=").replace(r"\x26", "&"))
        return test(host + next_page[0].attrib["onclick"].replace(
            "window.location=\'", "").replace("\'", "").replace(r"\x3d", "=").
                    replace(r"\x26", "&"))


def save_data(data, method="a+", file_name="google.csv"):
    with open(file_name, method, encoding="utf-8", newline="") as f:
        csv_f = csv.writer(f)
        csv_f.writerow([
            data.get("name", ""),
            data.get("title", ""),
            data.get("img_url", ""),
            data.get("Citations", ""),
            data.get("h-index", ""),
            data.get("i10-index", ""),
            data.get("year", "")
        ])


def user_detail(user_url):
    """
    :param user_url: eg: https://scholar.google.ca/citations?hl=en&user=Be-s2h4AAAAJ
    :return: eg:
        {
        'year': {
            '2017': '1396',
            '2018': '1441',
            '2019': '803'
         },
        'Citations': {'ALL': '10106', 'Since 2014': '6944'},
        'h-index': {'ALL': '51', 'Since 2014': '41'},
        'i10-index': {'ALL': '190', 'Since 2014': '162'}
        }

    """
    res = requests.get(user_url, headers=headers)
    selector = etree.HTML(res.text)
    detail = dict()
    year = dict()
    for index, span in enumerate(
            list(selector.xpath('//span[@class="gsc_g_t"]'))):
        year[span.xpath("string(.)")] = selector.xpath(
            'string(//a[@class="gsc_g_a"][{}])'.format(index + 1))
    detail["year"] = year
    for tr in selector.xpath('//table[@id="gsc_rsb_st"]/tbody//tr'):
        detail[[td.xpath("string(.)") for td in tr.xpath('.//td')][0]] = {
            "ALL": [td.xpath("string(.)") for td in tr.xpath('.//td')][1],
            "Since 2014":
            [td.xpath("string(.)") for td in tr.xpath('.//td')][2]
        }
    return detail


def main():
    test(start_url)


if __name__ == '__main__':
    main()
