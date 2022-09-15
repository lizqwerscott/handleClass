import os
from sys import float_repr_style
import xlrd
import json
import math
import re

def split_name_number(value: str):
    result = re.search(r"\d+-\d+\S+", value)
    name = value[0:result.start()]
    number = value[result.start():result.end()]
    return [name, number]

def parse_value_info_l(value: str, result: dict):
    class_value_week = {"teacher": "", "room": "", "week": ""}
    data = value.split(" ")
    result["name"] = data[0]
    a = split_name_number(data[1])
    class_value_week["week"] = a[1]
    class_value_week["teacher"] = a[0]
    class_value_week["room"] = data[2]
    result["weeks"].append(class_value_week)

def parse_value_info(value: str | list):
    class_value = {"name": "", "weeks": []}

    if type(value) is str:
        if value == "":
            return ""
        parse_value_info_l(value, class_value)

    if type(value) is list:
        parse_value_info_l(value[1], class_value)
        parse_value_info_l(value[0], class_value)

    return class_value

def parse_class_value(value: str):
    class_value = {"signal": "", "double": ""}

    #解析单双周
    data = value.split(";")
    if len(data) > 1:
        if "双周" in data[0] or "单周" in data[0]:
            class_value["double"] = data[0]
            class_value["signal"] = data[1]
        else:
            class_value["double"] = data
            class_value["signal"] = data
    else:
        if "双周" in data[0]:
            class_value["double"] = data[0]
        else:
            if "单周" in data[0]:
                class_value["signal"] = data[0]
            else:
                class_value["signal"] = data[0]
                class_value["double"] = data[0]

    result = {}
    result["signal"] = parse_value_info(class_value["signal"])
    result["double"] = parse_value_info(class_value["double"])
    return result

def get_class(path: str, class_name: str, type: str = "jk") -> list:
    xlsx = xlrd.open_workbook(path)
    table = xlsx.sheet_by_index(0)

    #寻找班级所在的行
    rows = -1
    for i in range(table.nrows):
        class_line = table.cell_value(i, 0)
        print(class_line)
        if class_name in class_line:
            print("rows: ", i)
            rows = i
            break
    if rows == -1:
        print("没找到课表，请确认班级名!")
        return None

    #提取课表
    class_week = []
    for i in range(5):
        day = []
        if type == "jk":
            for j in range(1, 9, 2):
                value = table.cell_value(rows, i * 10 + j + 1)
                res = parse_class_value(value)
                res["index"] = int(j / 2) + 1
                day.append(res)
        if type == "zn":
            for j in range(1, 5):
                value = table.cell_value(rows, i * 4 + j)
                res = parse_class_value(value)
                res["index"] = j
                day.append(res)
        class_week.append(day)

    print(json.dumps(class_week, ensure_ascii=False, indent=4))
    return class_week

if __name__ == "__main__":
    # df = load_data("./execl/计算机系班级大课表2022-2023-1.xls"
    # data = df.head()
    class_name = "智能20-1"
    # class_name = "网工20-3"
    class_week = get_class("./execl/班级大课表2022-2023-1_添加实验课.xls", class_name, type="zn")
    # class_week = get_class("./execl/new.xls", class_name, type="jk")
    if class_week is not None:
        with open("./data/{}.json".format(class_name), "w") as f:
            json.dump(class_week, f, indent=4, ensure_ascii=False)

        print("单周:")
        for i in range(len(class_week)):
            print("星期{}".format(i + 1))
            for j in range(len(class_week[i])):
                print("第{}节课".format(j + 1), end=" ")
                print(class_week[i][j]["signal"])
        print("双周:")
        for i in range(len(class_week)):
            print("星期{}".format(i + 1))
            for j in range(len(class_week[i])):
                print("第{}节课".format(j + 1), end=" ")
                print(class_week[i][j]["double"])
