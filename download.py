import os
import time
import pandas as pd
import xlwings as xw
import requests
from bs4 import BeautifulSoup


def get_date_and_download_path(url):
    """
    The method returns the dates of the report in the URL provided and their matching files paths.
    :param url: String. The health department url
    :return: Lists. The dates of the report in the URL provided and their matching files paths.
    """
    html = requests.get(url).content.decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    td_date = soup.findAll('td', attrs={'class': 'gvDate'})
    dates = [t.text.strip() for t in td_date]
    td_path = soup.findAll('td', attrs={'class': "Icon legislationIcon"})
    download_pathes = [t.contents[0].attrs['href'] for t in td_path]
    return dates, download_pathes


def delete_redundant_report(data_path):
    files_names = os.listdir(data_path)
    for name in files_names:
        if name.endswith('.2015.xlsx') or name.endswith('01.2016.xlsx') or name.endswith('02.2016.xlsx') or\
                name.endswith('03.2016.xlsx') or name.endswith('04.2020.xlsx') or name.endswith('05.2020.xlsx') or\
                name.endswith('06.2020.xlsx'):
            os.remove("{}/{}".format(data_path,name))


def download_data(data_path):
    """
    The method download the health department data to the provided data_path
    :param data_path: The path the data will be download to.
    """
    url_base = 'https://www.health.gov.il/UnitsOffice/HD/PH/epidemiology/Pages/epidemiology_report.aspx?WPID=WPQ7&PN='
    addons = [n for n in range(1,6)]
    if not os.path.exists(data_path):
        os.makedirs(data_path)
    for addon in addons:
        dates, download_pathes = get_date_and_download_path(url_base+str(addon))
        place = 0
        print(addon)
        for d in download_pathes:
            myfile = requests.get(d)
            open("{}/{}.xlsx".format(data_path,dates[place]), 'wb').write(myfile.content)
            print("{}/{}.xlsx".format(data_path,dates[place]))
            place = place +1
        time.sleep(15)
    delete_redundant_report(data_path)


def get_cities_names(file_path):
    excel_app = xw.App(visible=False)
    wb = excel_app.books.open(file_path)
    sheet = wb.sheets[0]
    cities = sheet['B41:P41'].options(pd.DataFrame, index=False, header=True).value
    return cities


def get_diseases_names(file_path):
    excel_app = xw.App(visible=False)
    wb = excel_app.books.open(file_path)
    sheet = wb.sheets[0]
    diseases1 = sheet['A43:A84'].options(pd.DataFrame, index=False, header=False).value
    diseases2 = sheet['A87:A128'].options(pd.DataFrame, index=False, header=False).value
    diseases = pd.concat([diseases1,diseases2], ignore_index=True)
    return diseases


def get_data_from_files(data_path):
    files_names = os.listdir(data_path).copy()
    cities = get_cities_names("{}/{}".format(data_path,files_names[0]))
    diseases = get_diseases_names("{}/{}".format(data_path,files_names[0]))
    columns = cities.copy()
    columns.append(['disease','Date'])
    final_data = pd.DataFrame(columns=columns)
    for name in files_names:
        if name.split('$')[0] != '~':
            excel_app = xw.App(visible=False)
            wb = excel_app.books.open("{}/{}".format(data_path,name))
            sheet = wb.sheets[0]
            data1 = sheet['B43:P84'].options(pd.DataFrame, index=False, header=False).value
            data2 = sheet['B87:P128'].options(pd.DataFrame, index=False, header=False).value
            data = pd.concat([data1,data2], ignore_index=True)
            data.columns = list(cities)
            data['Disease'] = list(diseases[0])
            data['Date'] = name.split('.x')[0]
            final_data = pd.concat([final_data,data], ignore_index=True)
            wb.close()
            excel_app.quit()
            print(name)
    return final_data


def save_data(raw_path, saving_path):
    download_data(raw_path)
    get_data_from_files(raw_path).to_csv('{}/health_data.csv'.format(saving_path))


def load_data(raw_path, saving_path):
    if not os.path.exists('{}/health_data.csv'.format(saving_path)):
        save_data(raw_path, saving_path)
    return pd.read_csv('{}/health_data.csv'.format(saving_path))


