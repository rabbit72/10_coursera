import sys
import os.path
import re
import random
import requests
from datetime import datetime
from collections import OrderedDict
from bs4 import BeautifulSoup as soup
from openpyxl import Workbook


def fetch_page(url, params=None):
    response = requests.get(url, params)
    response.encoding = 'utf8'
    return response.text


def get_random_courses(xml_page, number_courses=20):
    soup_courses = soup(xml_page, 'xml')
    courses_list = soup_courses.text.split()
    if number_courses:
        return random.sample(courses_list, number_courses)
    return courses_list


def get_start_date(course):
    attrs = {'class': 'startdate rc-StartDateString caption-text'}
    start_date = course.find(attrs=attrs)
    return start_date.text if start_date else None


def get_user_rating(course):
    raw_rating = course.find('div', {'class': 'ratings-text bt3-hidden-xs'})
    if not raw_rating:
        return None
    course_rating = re.findall('\d\.\d', raw_rating.text)
    return float(course_rating[0])


def get_name_course(course):
    name_course = course.find('h1', {'class': 'title display-3-text'})
    return name_course.text if name_course else None


def get_language_course(course):
    language = course.find(attrs={'class': 'rc-Language'})
    return language.text if language else None


def get_weeks_number(course):
    weeks_number = course.find(attrs={'class': 'rc-WeekView'})
    return len(weeks_number) if weeks_number else None


def get_courses_info(courses_pages):
    courses_info = []
    for course_page in courses_pages:
        course = soup(course_page, 'lxml')
        course_info = OrderedDict([
            ('name', get_name_course(course)),
            ('language', get_language_course(course)),
            ('start_date', get_start_date(course)),
            ('weeks_number', get_weeks_number(course)),
            ('user_rating', get_user_rating(course))
        ])
        courses_info.append(course_info)
    return courses_info


def fill_workbook(workbook=Workbook()):
    sheet = workbook.active
    sheet.title = 'courses'
    title_row = [title for title in courses_info[0].keys()]
    sheet.append(title_row)
    for course_info in courses_info:
        sheet.append([one_info for one_info in course_info.values()])
    return workbook


def save_workbook(workbook, directory_for_save=None):
    file_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S.xlsx')
    current_directory = '.'
    if not directory_for_save:
        directory_for_save = current_directory
    file_path = os.path.join(directory_for_save, file_name)
    workbook.save(file_path)


if __name__ == '__main__':
    try:
        path_for_save = sys.argv[1]
        if not os.path.isdir(path_for_save):
            exit('Directory for saving not found')
        url_xml_coursers = 'https://www.coursera.org/sitemap~www~courses.xml'
        courses_urls = get_random_courses(fetch_page(url_xml_coursers))
        courses_pages = [fetch_page(course_url) for course_url in courses_urls]
        courses_info = get_courses_info(courses_pages)
        save_workbook(fill_workbook(), path_for_save)
    except IndexError:
        exit('Path for saving not input')
    except requests.ConnectionError:
        exit('Check your connection')
