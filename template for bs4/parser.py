import csv
import time

import openpyxl
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from progress.bar import ShadyBar
from progress.spinner import Spinner


def read_from_exel(file_path):
    """Read data from the exel in first row and return list"""
    book = openpyxl.open(file_path, read_only=True)
    sheet = book.active
    return [sheet[row][0].value.lower().strip() for row in range(1, sheet.max_row + 1)]


def write_to_exel(speciality, link, recommend_job_experience):
    """Write to exel file data"""
    with open('data.csv', 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(
            (
                f'{speciality}',
                f'{link}',
                f'{recommend_job_experience}'
            )
        )


def create_csv():
    """Function to crete new csv file in current directory"""
    with open('data.csv', 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(
            (
                "Specialities",
                "Links",
                "Recommend job experience"
            )
        )


def process_url_domain(first_url, second_url):
    """Return true if domain in links has equal word"""
    first_index = first_url[::-1].find('/')
    second_index = second_url[1:].find('/')
    first_string = first_url[::-1][:first_index + 1][::-1]
    second_string = second_url[:second_index + 1]
    return first_string == second_string


def collect_data(box_of_search, urls):
    ua = UserAgent()

    headers = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/'
                  'webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'User_Agent': ua.random
    }

    # bar for get progress result in terminal
    spinner = Spinner()
    with ShadyBar(f'  Parsing ...', max=len(urls)) as bar:

        for url in urls:

            # request to site and parsing data
            response = requests.get(url=url, headers=headers)

            soup = BeautifulSoup(response.text, 'lxml')

            # find all tags with href
            jobs = soup.find_all('a')
            collection = []

            # find all needed specialities
            for item in jobs:
                spinner.next()
                search_name_or_link = item.text.lower()
                index = 0

                # search in tags (h1, ..., h6) name of speciality for determine domain of link
                for i in range(2, 7):
                    var = item.find(f'h{i}')
                    if var:
                        index = i
                        break

                if index:
                    # search speciality key word in string from site
                    for word in box_of_search:
                        if word in item.find(f'h{index}').text.lower():
                            # if link no have protocol prefix
                            if 'http' in item.get("href"):
                                collection.append(f'{item.find(f"h{index}").text}: '
                                                  f'{item.get("href")}')
                            else:
                                collection.append(f'{item.find(f"h{index}").text}: '
                                                  f'{url + item.get("href")}')
                else:
                    for word in box_of_search:
                        if word in search_name_or_link:
                            if 'http' in item.get("href"):
                                collection.append(f'{search_name_or_link}: '
                                                  f'{item.get("href")}')
                            else:
                                if process_url_domain(url, item.get("href")):
                                    index2 = item.get("href")[1:].find('/')
                                    collection.append(f'{search_name_or_link}: '
                                                      f'{url + item.get("href")[index2 + 1:]}')
                                else:
                                    collection.append(f'{search_name_or_link}: '
                                                      f'{url + item.get("href")}')

            for i in range(len(collection)):
                # analyzing all valid vacancy on years of experience
                spinner.next()

                url = collection[i][collection[i].find(': ') + 2:]
                src = requests.get(url=url, headers=headers)
                soup = BeautifulSoup(src.text, 'lxml')
                string = soup.find_all('li')

                # process string
                new_string = [j.text.lower() for j in string]
                string = ' '.join(new_string)

                # filter years of experience
                if string.find('years') == -1:
                    write_to_exel(collection[i][:collection[i].find(': ') + 2],
                                  collection[i][collection[i].find(': ') + 2:], 'empty')
                else:
                    statement = string[string.find("years") - 4:string.find("years") + 6]
                    if ('1' in statement and '10' not in statement) or '2' in statement or \
                            ('3' in statement and '+' not in statement):
                        write_to_exel(collection[i][:collection[i].find(': ') + 2],
                                      collection[i][collection[i].find(': ') + 2:], f'{statement.strip()}')

            bar.next()
            spinner.next()

    print('Done!')
    time.sleep(3)


def main():
    file_path_specialities = input("Input file name, example [data.xlsx]: ")
    file_path_links = input("Input file name, example [data.xlsx]: ")

    create_csv()
    try:
        collect_data(read_from_exel(file_path_specialities), read_from_exel(file_path_links))
    except (FileExistsError, FileNotFoundError):
        print("No such file or directory!")


if __name__ == '__main__':
    main()
