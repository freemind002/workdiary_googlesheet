import re
import string

import arrow
import pygsheets
import settings


class WorkdiaryGooglesheet(object):
    def __init__(self) -> None:
        self.gc = pygsheets.authorize(service_account_file=settings.json_path)
        self.red_num, self.green_num, self.blue_num = 1, 0, 0
        self.update_date = arrow.now().format('YYYY-MM-DD')

    def get_member_list(self):
        """
        取得組員名單
        """
        sheet = self.gc.open_by_url(settings.survey_url).worksheet_by_title(settings.member_list)
        member_list = sheet.get_all_records(numericise_data=False)
        member_list = [v for i in member_list for k, v in i.items()]

        return member_list

    def get_holiday_list(self):
        """
        取得假日名單
        """
        sheet = self.gc.open_by_url(settings.survey_url).worksheet_by_title(settings.holiday_list)
        holiday_list = sheet.get_all_records(numericise_data=False)
        year_month = re.findall(r'\d{4}-\d{2}', self.update_date)[0]
        holiday_list = [
            i[f'假日名單_{year_month.split("-")[0]}'] for i in holiday_list
            if year_month in i[f'假日名單_{year_month.split("-")[0]}']
        ]

        return holiday_list

    def get_date_list(self):
        """
        取得該月的所有日期的名單
        """
        update_date_list = self.update_date.split('-')
        first_day = arrow.Arrow(int(update_date_list[0]), int(update_date_list[1]), 1)
        last_day = first_day.shift(months=1, days=-1)
        date_list = [r.format('YYYY-MM-DD') for r in arrow.Arrow.range('day', first_day, last_day)]

        return date_list

    def add_conditional_formatting(self, sheet_updatedate, holiday_list):
        """
        使用條件式格式設定
        """
        total_cols, row_count = sheet_updatedate.cols, sheet_updatedate.rows
        column_num_list = [letter for letter in string.ascii_uppercase] + \
            [letter1 + letter2 for letter1 in string.ascii_uppercase for letter2 in string.ascii_uppercase]
        conditional_formatting_01, conditional_formatting_02 = column_num_list[0], column_num_list[total_cols-1]
        holiday_list = ['A$1=DATE(' + i.replace('-', ',') + ')' for i in holiday_list]
        conditional_formatting_formula = '=OR(' + ', '.join(holiday_list) + ')'
        sheet_updatedate.add_conditional_formatting(
            f'{conditional_formatting_01}1',
            f'{conditional_formatting_02}{row_count}',
            'CUSTOM_FORMULA',
            {'backgroundColor': {'red': self.red_num, 'green': self.green_num, 'blue': self.blue_num}},
            [f'{conditional_formatting_formula}']
        )

    def make_new_sheet(self, member_list, holiday_list, date_list):
        """
        新增新的工作表，或指定工作表，開始進行動作
        """
        sheet = self.gc.open_by_url(settings.survey_url)
        wks_list = sheet.worksheets()
        wks_title_list = [i.title for i in wks_list]
        update_date_list = self.update_date.split('-')
        # 查看該工作表是否存在，如不存在則新增並插入title
        new_tab = f"{update_date_list[0]}年{update_date_list[1]}月"
        if new_tab in wks_title_list:
            sheet_updatedate = sheet.worksheet_by_title(new_tab)
        else:
            sheet.add_worksheet(title=new_tab, rows=10000, cols=len(date_list)+10)
            sheet_updatedate = sheet.worksheet_by_title(new_tab)
            sheet_updatedate.update_values("B1", [date_list])

        # 插入組員名單
        exits_records_list = sheet_updatedate.get_all_records(numericise_data=False)
        if not exits_records_list:
            member_list = [[i] for i in member_list]
            sheet_updatedate.append_table(values=member_list, start='A2')
            # 進行條件式格式設定，將假日的部份用條件式格式設定變色
            self.add_conditional_formatting(sheet_updatedate, holiday_list)
            # 在「工作日誌列表」的工作表增加hyperlink，以倒序的方式排列
            sheet_updatedate_url = re.search(r"#gid=(\d+)", sheet_updatedate.url).group(0)
            sheet_datalist = sheet.worksheet_by_title('工作日誌列表')
            sheet_datalist.insert_rows(1)
            value = f'=HYPERLINK("{sheet_updatedate_url}", "{new_tab}")'
            sheet_datalist.update_value('A2', value)
        else:
            exists_member_list = [i[''] for i in exits_records_list]
            member_list = [[i] for i in member_list if i not in exists_member_list]
            sheet_updatedate.append_table(values=member_list, start='A2')

        # 凍結第一列
        sheet_updatedate.frozen_rows = 1

    def run_all(self):
        member_list = self.get_member_list()
        holiday_list = self.get_holiday_list()
        date_list = self.get_date_list()
        self.make_new_sheet(member_list, holiday_list, date_list)

    def main(self):
        try:
            self.run_all()
            print("程序順利執行完成")
        except Exception as e:
            print("發生錯誤")
            print(e)


if __name__ == "__main__":
    workdiarygooglesheet = WorkdiaryGooglesheet()
    workdiarygooglesheet.main()
