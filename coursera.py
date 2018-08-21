from bs4 import BeautifulSoup
from openpyxl import Workbook
from lxml import etree
import requests
import sys
import os


def get_courses_links(xml, courses_count):
    courses_list = []
    xml_root = etree.XML(xml)
    xml_locs = xml_root.findall(
        ".//{http://www.sitemaps.org/schemas/sitemap/0.9}loc")
    for course in xml_locs[:courses_count]:
        courses_list.append(course.text)
    return courses_list


def get_course_info(html_feed, main_course_keys):
    course_info = dict.fromkeys(main_course_keys)
    course_soup = BeautifulSoup(html_feed, "html.parser")
    course_info["Name"] = course_soup.find("h1", {"class": "title"}).text
    course_info["Start date"] = course_soup.find("div", "startdate").text
    tbody = course_soup.find("table", {"class": "basic-info-table"})
    table_tds = tbody.find_all("td")
    table_titles = [title.text for title in table_tds[::2]]
    table_datas = [table_data.text for table_data in table_tds[1::2]]
    for title in table_titles:
        if title == "User Ratings":
            course_info["User Ratings"] = course_soup.find(
                "div",
                {"class": "ratings-text"}).text
        elif title in main_course_keys:
            course_info[title] = table_datas[table_titles.index(title)]
    return course_info


def output_courses_info_to_xlsx(worksheet, all_courses_info, main_course_keys):
    worksheet.append(main_course_keys)
    for course in all_courses_info:
        course_row = [course[title] for title in main_course_keys]
        worksheet.append(course_row)
    return worksheet


if __name__ == "__main__":
    filepath = os.getcwd()
    try:
        file_name = sys.argv[1]
    except IndexError:
        sys.exit("Enter file name")
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    courses_count = 20
    workbook = Workbook()
    worksheet = workbook.active
    response = requests.get(url)
    courses_list = get_courses_links(response.content, courses_count)
    main_course_keys = ["Name", "Start date",
                        "Language", "User Ratings", "Commitment"]
    all_courses_info = []
    for course in courses_list:
        html_feed = requests.get(course)
        html_feed.encoding = "utf-8"
        all_courses_info.append(
            get_course_info(html_feed.text, main_course_keys))
    output_courses_info_to_xlsx(worksheet, all_courses_info, main_course_keys)
    workbook.save(os.path.join(filepath, ".".join((file_name,"xlsx"))))
