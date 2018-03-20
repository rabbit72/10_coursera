import sys
import os.path
import re
import random
import requests
from datetime import datetime
from collections import OrderedDict
from bs4 import BeautifulSoup as soup
from openpyxl import Workbook


def get_courses_list(number_courses=None):
    link_coursers = 'https://www.coursera.org/sitemap~www~courses.xml'
    response = requests.get(link_coursers)
    soup_courses = soup(response.text, 'xml')
    courses_list = soup_courses.text.split()
    if number_courses:
        return random.sample(courses_list, number_courses)
    return courses_list


def get_start_date(course):
    start_date = course.find(attrs={'class': 'startdate rc-StartDateString caption-text'})
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


def get_courses_info(urls_list):
    for url in urls_list:
        response = requests.get(url)
        response.encoding = 'utf8'
        course = soup(response.text, 'lxml')
        yield OrderedDict([
            ('name', get_name_course(course)),
            ('url', url),
            ('language', get_language_course(course)),
            ('start_date', get_start_date(course)),
            ('weeks_number', get_weeks_number(course)),
            ('user_rating', get_user_rating(course))
        ])


def output_courses_info_to_xlsx(directory, courses_info):
    file_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S.xlsx')
    file_path = os.path.join(directory, file_name)
    wb = Workbook()
    sheet = wb.active
    sheet.title = 'courses'
    title_row = [title for title in courses_info[0].keys()]
    sheet.append(title_row)
    for course in courses_info:
        sheet.append([value for value in course.values()])
    wb.save(file_path)


if __name__ == '__main__':
    try:
        path_for_save = sys.argv[1]
        if not os.path.isdir(path_for_save):
            exit('Directory for saving not found')
        number_courses = 20
        courses_list = get_courses_list(number_courses)
        # path_for_save = '.'
        # courses_list = ['https://www.coursera.org/learn/hanyu-yufa']
        courses_info = [info for info in get_courses_info(courses_list)]
        output_courses_info_to_xlsx(path_for_save, courses_info)
    except IndexError:
        exit('Path for saving not input')
    except requests.ConnectionError:
        exit('Check connection')
