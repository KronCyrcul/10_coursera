from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests
import os
import re


def get_courses_list(url, courses_count, page_search, courses_links):
    params = {"start": page_search}
    response = requests.get(url)
    response_soup = BeautifulSoup(response.text, "html.parser")
    research_result = response_soup.find_all(
        "div",
        {"data-courselenium": "catalog_search_result"})
    courses_links += [link["href"] for link in research_result[0].find_all(
        "a", href=re.compile(r"/learn/"))]
    if len(courses_links) <= courses_count:
        page_search += 20
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


def output_courses_info_to_xlsx(workbook, course_info, course_row):
    for key, value in course_info.items():
        try:
            first_row = worksheet.rows[0]
            first_row_values = [cell.value for cell in first_row]
            worksheet.cell(
                row=course_row,
                column=first_row_values.index(key)+1).value = value
        except IndexError:
            worksheet.append(list(course_info.keys()))
            first_row = worksheet.rows[0]
            first_row_values = [cell.value for cell in first_row]
            worksheet.cell(
                row=course_row,
                column=first_row_values.index(key)+1).value = value
        except ValueError:
            worksheet.cell(
                row=1,
                column=len(first_row_values)+1).value = key
            worksheet.cell(
                row=course_row,
                column=len(first_row_values)+1).value = value
    return workbook


if __name__ == "__main__":
    filepath = os.getcwd()
    url = "https://www.coursera.org"
    courses_url = "{}/{}".format(url, "courses")
    courses_count = 20
    page_start = 0
    courses_links = get_courses_list(
        courses_url,
        courses_count,
        page_start,
        list())
    workbook = Workbook()
    worksheet = workbook.active
    course_row = 2
    for course in courses_links:
        course_info = get_courses_info(url, course)
        workbook = output_courses_info_to_xlsx(
            workbook,
            course_info,
            course_row)
        course_row += 1
    workbook.save("{}/{}".format(filepath, "courses.xlsx"))
