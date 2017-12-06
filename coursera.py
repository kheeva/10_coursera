#!/usr/bin/env python
import sys
import requests
from xml.etree import ElementTree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_courses_list():
    url = 'https://www.coursera.org/sitemap~www~courses.xml'
    xml_courses = requests.get(url).text
    courses_list = []
    # print(xml_courses)
    root = ElementTree.fromstring(xml_courses)
    # print(root.tag)
    namespace = root.tag[1:root.tag.index('}')]
    namespace_map = {"ns": namespace}
    for elem in root.findall('ns:url', namespace_map)[:3]:
        courses_list.append(elem.find('ns:loc', namespace_map).text)

    return courses_list


def get_course_info(course_slug):
    soup = BeautifulSoup(requests.get(course_slug).text, 'html.parser')
    course_name = soup.find('h1').text
    course_language = soup.find('div', class_='rc-Language').text
    course_start_date = soup.find(
        'div',
        class_='startdate rc-StartDateString caption-text').text
    course_weeks_number = len(soup.findAll('div', class_='week'))
    try:
        course_score = soup.find('div',
                                 class_='ratings-text bt3-visible-xs').text
    except AttributeError:
        course_score = 'not rated'

    return {'name': course_name,
            'language': course_language,
            'start_date': course_start_date,
            'number_of_weeks': course_weeks_number,
            'score': course_score}


def output_courses_info_to_xlsx(file_path, courses):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet['A1'] = 'name'
    worksheet['A2'] = 'language'
    workbook.save(file_path)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        exit("Usage: python coursera.py path_for_saving_file")

    courses_dict = {}
    courses = get_courses_list()
    for course in courses:
        courses_dict[course] = get_course_info(course)

    output_courses_info_to_xlsx(sys.argv[1], courses_dict)