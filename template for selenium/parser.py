import csv

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

options = webdriver.ChromeOptions()
options.add_argument("user-agent=Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:84.0) Gecko/20100101 Firefox/84.0")
options.add_argument("--disable-blink-features=AutomationControlled")
s = Service(executable_path="D:\\Templates_to_data_parsing\\template for selenium\\chromedriver.exe")
driver = webdriver.Chrome(service=s, options=options)


def read_from_exel(file_path):
    """Read data from the exel in first row and return list"""
    book = openpyxl.open(file_path, read_only=True)
    sheet = book.active
    return [sheet[row][0].value.lower().strip() for row in range(1, sheet.max_row + 1)]


def write_to_exel(code, link):
    """Write to exel file data"""
    with open('data.csv', 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(
            (
                f'{code}',
                f'{link}'
            )
        )


def create_csv():
    """Function to crete new csv file in current directory"""
    with open('data.csv', 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(
            (
                "Codes",
                "Links"
            )
        )


def collect_data(links, box):
    # count for pages
    count = -1

    try:
        for link in links:
            count += 1

            for page in range(1, box[count]):
                driver.get(url=f'{link[:-1]}{page}')

                driver.implicitly_wait(5)
                print(f'Current url: {driver.current_url}')
                print(driver.window_handles)

                # find all clickable elements
                items = driver.find_elements(By.CLASS_NAME, value="SYep11sJh1qGxMJNwO1X")

                for i in range(len(items)):
                    try:
                        # go to needed page
                        driver.get(url=f'{link[:-1]}{page}')
                        items = driver.find_elements(By.CLASS_NAME, value="SYep11sJh1qGxMJNwO1X")
                        print(f'Current url: {driver.current_url}')

                        driver.implicitly_wait(5)
                        # go to next page
                        items[i].click()
                        print(driver.window_handles)

                        # get needed data from tags
                        data = driver.find_element(By.CLASS_NAME, value="nsLdAwtaP1DS5KR540i9").text

                        write_to_exel(f'{data[data.find(":") + 2:]}', f'{driver.current_url}')

                    except BaseException as e:
                        driver.get(url=f'{link[:-1]}{page}')
                        driver.implicitly_wait(5)
                        print('Error!')

    except Exception as ex:
        print(ex)
    finally:
        driver.close()
        driver.quit()


def main():
    create_csv()
    collect_data(read_from_exel('links.xlsx'), [135, 138, 60, 21, 257, 16, 44, 153, 131, 15, 15, 70])


if __name__ == '__main__':
    main()
