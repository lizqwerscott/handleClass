import os
import xlrd
import json

def parse_value_info(value: str):
    class_value = {"teacher": "", "room": "", "name": ""}
    if value == "":
        return ""
    data = value.split(" ")
    class_value["name"] = data[0]
    class_value["teacher"] = data[1]
    class_value["room"] = data[2]
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
            class_value["double"] = value
            class_value["signal"] = value

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

def get_class(path: str, class_name: str) -> list:
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
        # for j in range(1, 9, 2):
        for j in range(1, 5):
            # value = table.cell_value(rows, i * 10 + j + 1)
            value = table.cell_value(rows, i * 4 + j)
            day.append(parse_class_value(value))
            # print(value)
        class_week.append(day)
    # print(json.dumps(class_week, ensure_ascii=False, indent=4))
    return class_week

if __name__ == "__main__":
    # df = load_data("./execl/计算机系班级大课表2022-2023-1.xls"
    # data = df.head()
    class_week = get_class("./execl/班级大课表2022-2023-1_添加实验课.xls", "智能20-1")
    # class_week = get_class("./execl/new.xls", "网工20-3")
    if class_week != None:
        with open("./data/智能20-1.json", "w") as f:
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
