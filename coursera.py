from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests
import os
import re


def get_courses_list(url, courses_count, page_start, courses_links):
    params = {"start": page_start}
    response = requests.get(url)
    response_soup = BeautifulSoup(response.text, "html.parser")
    research_result = response_soup.find_all(
        "div",
        {"data-courselenium": "catalog_search_result"})
    all_links = research_result[0].find_all("a", href=True)
    courses_links += [link["href"] for link in all_links if
        re.search(r"/learn/", link["href"])]
    if len(courses_links) <= courses_count:
        page_start += 20
        get_courses_list(url, courses_count, page_start, courses_links)
    return courses_links


def get_courses_info(url, course_link):
    course_info = {}
    course_url = "{}{}".format(url, course_link)
    response = requests.get(course_url)
    course_soup = BeautifulSoup(response.text, "html.parser")
    course_info["Name"] = course_soup.find("h1", {"class": "title"}).text
    course_info["Start date"] = course_soup.find("div", "startdate").text
    tbody = course_soup.find_all("tbody")[0]
    for tbody["tr"] in tbody:
        row_data = tbody["tr"].find_all("td")
        info_title = row_data[0].text
        if info_title == "User Ratings":
            course_info[info_title] = row_data[1].find(
                "div",
                {"class": "ratings-text"}).text
        else:
            course_info[info_title] = row_data[1].text
    return course_info


def output_courses_info_to_xlsx(filepath, courses_info):
    workbook = Workbook()
    worksheet = workbook.active
    course_row = 2
    for course in courses_info:
        try:
            first_row = worksheet.rows[0]
        except IndexError:
            worksheet.append(list(course.keys()))
            first_row = worksheet.rows[0]
        first_row_values = [cell.value for cell in first_row]
        for key, value in course.items():
            if key not in first_row_values:
                worksheet.cell(
                    row=1,
                    column=len(first_row_values)+1).value = key
                worksheet.cell(
                    row=course_row,
                    column=len(first_row_values)+1).value = value
            else:
                worksheet.cell(
                    row=course_row,
                    column=first_row_values.index(key)+1).value = value
        course_row += 1
    workbook.save("{}/{}".format(filepath, "courses.xlsx"))


if __name__ == "__main__":
    filepath = os.getcwd()
    url = "https://www.coursera.org"
    courses_url = "{}/{}".format(url, "courses")
    courses_count = 20
    page_start = 0
    courses_links = get_courses_list(courses_url, courses_count, page_start, list())
    courses_info = []
    for course in courses_links:
        courses_info.append(get_courses_info(url, course))
    output_courses_info_to_xlsx(filepath, courses_info)
