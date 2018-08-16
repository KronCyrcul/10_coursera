from bs4 import BeautifulSoup
from openpyxl import Workbook
from lxml import etree
import requests
import os
import re


def encode_string(string):
    return string.encode("raw_unicode_escape").decode("utf-8")


def get_courses_list(response, courses_count):
    courses_list = []
    xml_root = etree.XML(response.text.encode('utf-8'))
    xml_locs = xml_root.findall(
        ".//{http://www.sitemaps.org/schemas/sitemap/0.9}loc")
    for course in xml_locs[:courses_count]:
        courses_list.append(course.text)
    return courses_list


def get_course_info(course_response):
    course_info = {}
    course_soup = BeautifulSoup(course_response.text, "html.parser")
    course_name = course_soup.find("h1", {"class": "title"}).text
    course_info["Name"] = encode_string(course_name)
    course_info["Start date"] = encode_string(
        course_soup.find("div", "startdate").text)
    tbody = course_soup.find("tbody")
    table_rows = tbody.find_all("tr")
    for row in table_rows:
        row_data = row.find_all("td")
        info_title = row_data[0].text
        if info_title == "User Ratings":
            course_info[info_title] = row_data[1].find(
                "div",
                {"class": "ratings-text"}).text
        else:
            course_info[info_title] = encode_string(row_data[1].text)
    return course_info


def output_courses_info_to_xlsx(worksheet, course_info):
    course_row = []
    try:
        first_row = worksheet.rows[0]
    except IndexError:
        worksheet.append(list(course_info.keys()))
        first_row = worksheet.rows[0]
    first_row_values = [cell.value for cell in first_row]
    for key in course_info.keys():
        if key not in first_row_values:
            worksheet.cell(
                row=1,
                column=len(first_row_values)+1).value = key
    for title in first_row_values:
        course_row += [course_info[title] if title in list(course_info.keys())
            else None]
    worksheet.append(course_row)
    return worksheet


if __name__ == "__main__":
    filepath = os.getcwd()
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    courses_count = 20
    workbook = Workbook()
    worksheet = workbook.active
    courses_url = "{}/{}".format(url)
    response = requests.get(courses_url)
    courses_list = get_courses_list(response, courses_count)
    for course in courses_list:
        course_response = requests.get(course)
        course_info = get_course_info(course_response)
        output_courses_info_to_xlsx(worksheet, course_info)
    workbook.save("{}/{}".format(filepath, "courses.xlsx"))
