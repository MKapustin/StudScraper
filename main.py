import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor
from xlwt import Workbook
from queue import Queue


def get_page_soup(url):
    response = requests.get(url)
    if response.status_code == 200:
        return BeautifulSoup(response.content, features="lxml")


def get_all_pagination_links(url):
    main_page_soup = get_page_soup(url)
    if main_page_soup:
        pagination_links = Queue()
        pagination_links.put(url)
        pagination_a_items = main_page_soup.find_all("a", {
            "class": "pagination-link"})
        for pagination_a_item in pagination_a_items:
            pagination_links.put("https://www.hertie-school.org" + pagination_a_item.attrs["href"])
        return pagination_links


def get_person_email(person_private_url):
    person_page_soup = get_page_soup(person_private_url)
    if person_page_soup:
        try:
            person_email = person_page_soup.find_all("div", {
                "class": "text-block"})[0].p.a.attrs["href"][7:]
        except IndexError:
            person_email = person_page_soup.find_all("li", {
                "class": "item mail"})[0].a.attrs["href"]
        except AttributeError:
            person_email = "AttributeError"
        return person_email


def add_persons_info_from_pagination_links(persons_info, pagination_links):
    while not pagination_links.empty():
        pagination_link = pagination_links.get()
        page_soup = get_page_soup(pagination_link)
        for person_div in page_soup.find_all("div", {
            "class": "grid-item col-xs-12 col-sm-6 col-md-3 type-link has-circle-image color-scheme--red"}):
            person_info = [person_div.h2.span.text.strip(),
                           person_div.find("div", {"class": "grid-item-text"}).text.strip(),
                           get_person_email("https://www.hertie-school.org" + person_div.a.attrs["href"])]
            persons_info.put(person_info)


def get_all_persons_info(
        main_url="https://www.hertie-school.org/en/who-we-are/people?no_cache=1&tx_lfpeopledirectory_list%5BselectedAreas%5D=&tx_lfpeopledirectory_list%5BselectedProgrammes%5D=&tx_lfpeopledirectory_list%5BselectedRoles%5D=&tx_lfpeopledirectory_list%5BselectedThemes%5D=&cHash=d65692a89b447c12e06cd55792edcc7d",
        workers_amount=6
):
    pagination_links = get_all_pagination_links(main_url)
    persons_info = Queue()  # [(Name, Status, Link to private page), ... ]
    if not pagination_links.empty():
        with ThreadPoolExecutor(max_workers=workers_amount) as executor:
            for index in range(workers_amount):
                executor.submit(add_persons_info_from_pagination_links, persons_info, pagination_links)
    return persons_info


def write_data_to_excel(data_list):
    wb = Workbook()
    sheet = wb.add_sheet('Sheet 1')
    for data_item_idx, data_item in enumerate(data_list):
        for data_idx, data in enumerate(data_item):
            sheet.write(data_item_idx, data_idx, data)
    wb.save('UniData.xlsx')


if __name__ == "__main__":
    write_data_to_excel(list(get_all_persons_info().queue))
