import csv
import math
from _datetime import datetime
import re
import matplotlib.pyplot as plt
import numpy as np

currency_to_rub = {
    "AZN": 35.68, "BYR": 23.91, "EUR": 59.90, "GEL": 21.74, "KGS": 0.76,
    "KZT": 0.13, "RUR": 1, "UAH": 1.64, "USD": 60.66, "UZS": 0.0055}


class Vacancy:
    def __init__(self, args):
        self.name = args[0]
        self.salary_from = float(args[1])
        self.salary_to = float(args[2])
        self.salary_currency = args[3]
        self.area_name = args[4]
        self.published_at = args[5]


class DataSet:
    def __init__(self, file_name):
        self.file_name = file_name
        self.vacancies_objects = list()

    @staticmethod
    def get_dataset(file_name):
        data = DataSet.csv_reader(file_name)
        dict_list = DataSet.csv_filter(data[0], data[1])
        dataset = DataSet(file_name)
        for item in dict_list:
            vacancy = Vacancy([f"{item['name']}", f"{item['salary_from']}", f"{item['salary_to']}",
                               f"{item['salary_currency']}", f"{item['area_name']}", f"{item['published_at']}"])
            vacancy.published_at = datetime.strptime(vacancy.published_at, "%Y-%m-%dT%H:%M:%S%z").year
            dataset.vacancies_objects.append(vacancy)
        return dataset

    @staticmethod
    def csv_reader(file_name):
        with open(file_name, "r", encoding="utf-8-sig", newline="") as file:
            data = [x for x in csv.reader(file)]
        columns = data[0]
        rows = [x for x in data[1:] if len(x) == len(columns) and not x.__contains__("")]
        return columns, rows

    @staticmethod
    def csv_filter(columns, rows):
        dic_list = list()
        for row in rows:
            dic_result = dict()
            for i in range(len(row)):
                items = DataSet.format_word(row[i].split('\n'))
                dic_result[columns[i]] = items[0] if len(items) == 1 else "; ".join(items)
            dic_list.append(dic_result)
        return dic_list

    @staticmethod
    def format_word(items):
        for i in range(len(items)):
            items[i] = " ".join(re.sub(r"\<[^>]*\>", "", items[i]).split())
        return items


class InputConnect:
    def __init__(self):
        self.params = InputConnect.get_params()

    @staticmethod
    def get_params():
        file_name = input("Введите название файла: ")
        profession_name = input("Введите название профессии: ")
        return file_name, profession_name

    @staticmethod
    def get_salary_by_city(data: DataSet):
        salary_by_city = dict()
        for vacancy in data.vacancies_objects:
            if math.floor(data.vacancy_rate_by_city[vacancy.area_name] / len(data.vacancies_objects) * 100) >= 1:
                if not salary_by_city.__contains__(vacancy.area_name):
                    salary_by_city[vacancy.area_name] = InputConnect.get_currency_to_rub(vacancy)
                else:
                    salary_by_city[vacancy.area_name] += InputConnect.get_currency_to_rub(vacancy)
        for key in salary_by_city:
            salary_by_city[key] = math.floor(salary_by_city[key] / data.vacancy_rate_by_city[key])
        return dict(sorted(salary_by_city.items(), key=lambda item: item[1], reverse=True))

    @staticmethod
    def get_vacancies_count_by_name(data: DataSet, name):
        vacancies_count = dict()
        for vacancy in data.vacancies_objects:
            if vacancy.name.__contains__(name) or name == "None":
                InputConnect.make_a_value_by_name(vacancies_count, vacancy.published_at)
        if len(vacancies_count) == 0:
            return {2022: 0}
        return vacancies_count

    @staticmethod
    def get_vacancy_rate_by_city(data: DataSet):
        vacancy_rate = dict()
        for vacancy in data.vacancies_objects:
            InputConnect.make_a_value_by_name(vacancy_rate, vacancy.area_name)
        return vacancy_rate

    @staticmethod
    def get_currency_to_rub(vacancy):
        course_money = currency_to_rub[vacancy.salary_currency]
        return int((vacancy.salary_from + vacancy.salary_to) * course_money / 2)

    @staticmethod
    def make_a_value_by_name(vacancy_dict: dict, name):
        if not vacancy_dict.__contains__(name):
            vacancy_dict[name] = 1
        else:
            vacancy_dict[name] = vacancy_dict[name] + 1

    @staticmethod
    def get_salary_by_name(data: DataSet, name):
        salary_by_name = dict()
        for vacancy in data.vacancies_objects:
            if vacancy.name.__contains__(name) or name == "None":
                if not salary_by_name.__contains__(vacancy.published_at):
                    salary_by_name[vacancy.published_at] = InputConnect.get_currency_to_rub(vacancy)
                else:
                    salary_by_name[vacancy.published_at] += InputConnect.get_currency_to_rub(vacancy)
        if len(salary_by_name) == 0:
            return {2022: 0}
        for key in salary_by_name.keys():
            if name == "None":
                salary_by_name[key] = math.floor(salary_by_name[key] / data.vacancies_count_by_year[key])
            else:
                salary_by_name[key] = math.floor(salary_by_name[key] / data.vacancies_count_by_profession_name[key])
        return salary_by_name

    @staticmethod
    def print_data_dictionary(self, data: DataSet):
        def get_correct_vacancy_rate(data: DataSet):
            data.vacancy_rate_by_city = {x: round(y / len(data.vacancies_objects), 4) for x, y in
                                         data.vacancy_rate_by_city.items()}
            # data.vacancy_rate_by_city = {k: v for k, v in data.vacancy_rate_by_city.items() if math.floor(v * 100 >= 1)}
            return dict(sorted(data.vacancy_rate_by_city.items(), key=lambda item: item[1], reverse=True))

        data.vacancies_count_by_year = InputConnect.get_vacancies_count_by_name(data, "None")
        data.salary_by_year = InputConnect.get_salary_by_name(data, "None")
        data.vacancies_count_by_profession_name = InputConnect.get_vacancies_count_by_name(data, self.params[1])
        data.salary_by_profession_name = InputConnect.get_salary_by_name(data, self.params[1])
        data.vacancy_rate_by_city = InputConnect.get_vacancy_rate_by_city(data)
        data.salary_by_city = InputConnect.get_salary_by_city(data)
        data.vacancy_rate_by_city = get_correct_vacancy_rate(data)
        data.dict_lict = [data.salary_by_year, data.salary_by_profession_name, data.vacancies_count_by_year,
                          data.vacancies_count_by_profession_name, dict(list(data.salary_by_city.items())[:10]),
                          data.vacancy_rate_by_city]
        print(f"Динамика уровня зарплат по годам: {data.salary_by_year}")
        print(f"Динамика количества вакансий по годам: {data.vacancies_count_by_year}")
        print(f"Динамика уровня зарплат по годам для выбранной профессии: {data.salary_by_profession_name}")
        print(
            f"Динамика количества вакансий по годам для выбранной профессии: {data.vacancies_count_by_profession_name}")
        print(f"Уровень зарплат по городам (в порядке убывания): {dict(list(data.salary_by_city.items())[:10])}")
        print(f"Доля вакансий по городам (в порядке убывания): {dict(list(data.vacancy_rate_by_city.items())[:10])}")


class Report:
    def __init__(self, dict_lict: list()):
        self.data = dict_lict

    def generate_img(self, profession_name):
        def myfunc(item):
            if item.__contains__(' '):
                return item[:item.index(' ')] + '\n' + item[item.index(' ') + 1:]
            elif item.__contains__('-'):
                return item[:item.index('-')] + '-\n' + item[item.index('-') + 1:]
            return item

        width = 0.3
        nums = np.arange(len(self.data[0].keys()))
        dx1 = nums - width / 2
        dx2 = nums + width / 2

        fig = plt.figure()
        ax = fig.add_subplot(221)
        ax.set_title("Уровень зарплат по годам")
        ax.bar(dx1, self.data[0].values(), width, label="средняя з/п")
        ax.bar(dx2, self.data[1].values(), width, label=f"з/п {profession_name.lower()}")
        ax.set_xticks(nums, self.data[0].keys(), rotation="vertical")
        ax.legend(fontsize=8)
        ax.tick_params(axis="both", labelsize=8)
        ax.grid(True, axis='y')

        ax = fig.add_subplot(222)
        ax.set_title("Количество вакансии по годам")
        ax.bar(dx1, self.data[2].values(), width, label="Количество вакансии")
        ax.bar(dx2, self.data[3].values(), width, label=f"Количество вакансии\n{profession_name.lower()}")
        ax.set_xticks(nums, self.data[0].keys(), rotation="vertical")
        ax.legend(fontsize=8)
        ax.tick_params(axis="both", labelsize=8)
        ax.grid(True, axis='y')

        ax = fig.add_subplot(223)
        ax.set_title("Уровень зарплат по городам")
        cities = list(map(myfunc, tuple(self.data[4].keys())))
        y_pos = np.arange(len(cities))
        ax.barh(y_pos, list(self.data[4].values()), align='center')
        ax.set_yticks(y_pos, labels=cities)
        ax.invert_yaxis()
        ax.grid(True, axis='x')

        ax = fig.add_subplot(224)
        ax.set_title("Доля вакансии по годам")
        labels = list(dict(list(self.data[5].items())[:10]).keys())
        labels.insert(0, "Другие")
        vals = list(dict(list(self.data[5].items())[:10]).values())
        vals.insert(0, 1 - sum(list(dict(list(self.data[5].items())[:10]).values())))
        ax.pie(vals, labels=labels, startangle=0, textprops={"fontsize": 6})
        plt.tight_layout()
        fig.set_size_inches(9.5, 7.5)
        plt.savefig("graph.png", dpi=120)
        return


inputParam = InputConnect()
dataSet = DataSet.get_dataset(inputParam.params[0])
InputConnect.print_data_dictionary(inputParam, dataSet)
report = Report(dataSet.dict_lict)
report.generate_img(inputParam.params[1])
