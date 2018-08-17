from bs4 import BeautifulSoup
from openpyxl import Workbook
from lxml import etree
import requests
import os
import re


def get_courses_links(response_content, courses_count):
    courses_list = []
    xml_root = etree.XML(response_content)
    xml_locs = xml_root.findall(
        ".//{http://www.sitemaps.org/schemas/sitemap/0.9}loc")
    for course in xml_locs[:courses_count]:
        courses_list.append(course.text)
    return courses_list


def get_course_info(course_response):
    course_info = {}
    course_soup = BeautifulSoup(course_response, "html.parser")
    course_info["Name"] = course_soup.find("h1", {"class": "title"}).text
    course_info["Start date"] = course_soup.find("div", "startdate").text
    tbody = course_soup.find("table", {"class": "basic-info-table"})
    table_tds = tbody.find_all("td")
    table_titles = [title.text for title in table_tds[::2]]
    table_datas = [data.text for data in table_tds[1::2]]
    course_info.update(dict(zip(table_titles, table_datas)))
    if "User Ratings" in table_titles:
        course_info["User Ratings"] = course_soup.find(
            "div",
            {"class": "ratings-text"}).text
    return course_info


def output_courses_info_to_xlsx(worksheet, all_courses_info):
    first_row = []
    for course in all_courses_info:
        first_row += [key for key in course.keys() if key not in first_row]
    worksheet.append(first_row)
    for course in all_courses_info:
        course_row = []
        course_keys = list(course.keys())
        missing_keys = [key for key in first_row if key not in course_keys]
        course.update(dict([(key, None) for key in missing_keys]))
        course_row = [course[title] for title in first_row]
        worksheet.append(course_row)
    return worksheet


if __name__ == "__main__":
    filepath = os.getcwd()
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    courses_count = 20
    workbook = Workbook()
    worksheet = workbook.active
    response = requests.get(url)
    courses_list = get_courses_links(response.content, courses_count)
    all_courses_info = []
    for course in courses_list:
        course_response = requests.get(course)
        course_response.encoding = "utf-8"
        all_courses_info.append(get_course_info(course_response.text))
    output_courses_info_to_xlsx(worksheet, all_courses_info)
    workbook.save("{}/{}".format(filepath, "courses.xlsx"))
