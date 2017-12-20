#!/usr/bin/env python
import sys
from random import randint
import requests
from xml.etree import ElementTree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_list_of_20_random_courses():
    url = 'https://www.coursera.org/sitemap~www~courses.xml'
    courses_list = []
    root = ElementTree.fromstring(requests.get(url).text)
    namespace = root.tag[1:root.tag.index('}')]
    namespace_map = {"ns": namespace}
    for elem in root.findall('ns:url', namespace_map):
        courses_list.append(elem.find('ns:loc', namespace_map).text)

    courses_limit = 20
    start_index = 0
    random_course = randint(start_index, len(courses_list) - courses_limit)
    return courses_list[random_course: random_course + courses_limit]


def get_course_info(course_slug):
    request = requests.get(course_slug)
    request.encoding = 'utf-8'
    soup = BeautifulSoup(request.text, 'html.parser')

    def get_course_attribute(html_tag, html_class):
        try:
            course_attribute = soup.find(html_tag, class_=html_class).text
        except AttributeError:
            return None
        else:
            return course_attribute

    course_name = get_course_attribute('h1', 'title display-3-text')
    course_language = get_course_attribute('div', 'rc-Language')
    course_start_date = get_course_attribute(
                            'div',
                            'startdate rc-StartDateString caption-text')
    course_score = get_course_attribute('div', 'ratings-text bt3-visible-xs')

    try:
        course_weeks_number = len(soup.findAll('div', class_='week'))
    except AttributeError:
        return None

    return {'name': course_name, 'language': course_language,
            'start_date': course_start_date,
            'number_of_weeks': course_weeks_number, 'score': course_score}


def output_courses_info_to_xlsx(file_path, courses):
    workbook = Workbook()
    worksheet = workbook.active

    worksheet['A1'] = 'Url'
    worksheet['B1'] = 'Name'
    worksheet['C1'] = 'Language'
    worksheet['D1'] = 'Start date'
    worksheet['E1'] = 'Weeks'
    worksheet['F1'] = 'Score'

    offset = 2
    for row, course in enumerate(courses):
        row_number = row + offset
        worksheet['A{}'.format(row_number)] = course
        worksheet['B{}'.format(row_number)] = courses[course]['name']
        worksheet['C{}'.format(row_number)] = courses[course]['language']
        worksheet['D{}'.format(row_number)] = courses[course]['start_date']
        worksheet['E{}'.format(row_number)] = courses[course][
            'number_of_weeks']
        worksheet['F{}'.format(row_number)] = courses[course]['score']

    workbook.save(file_path)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        exit("Usage: python coursera.py path_for_saving_file")

    courses_dict = {}
    try:
        courses_urls_list = get_list_of_20_random_courses()
    except ConnectionResetError:
        exit('Could\'t connect to coursera. Try again later.')
    else:
        for course_url in courses_urls_list:
            courses_dict[course_url] = get_course_info(course_url)

        output_courses_info_to_xlsx(sys.argv[1], courses_dict)
