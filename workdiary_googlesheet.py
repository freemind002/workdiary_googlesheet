import re
import string
from pathlib import Path
from typing import List, Text

import arrow
import polars as pl
import pygsheets


class WorkdiaryGooglesheet(object):
    def __init__(self) -> None:
        self.json_path = ["src", "api_key.json"]
        self.survey_url = (
            "https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXX/"
        )
        self.gs_url = pygsheets.authorize(
            service_account_file=Path(__file__).parent.joinpath(*self.json_path)
        ).open_by_url(self.survey_url)
        self.red_num, self.green_num, self.blue_num = 1, 0, 0
        self.update_date = arrow.now().format("YYYY-MM-DD")
        self.member_sheet, self.holiday_sheet = "組員名單", "假日名單"

    def get_member_list(self):
        """取得組員名單

        Returns:
            >>> get_member_list():
            List[Text]
        """
        sheet = self.gs_url.worksheet_by_title(self.member_sheet)
        member_list = sheet.get_all_records(numericise_data=False)
        member_list = (
            pl.LazyFrame(member_list)
            .select(pl.col(self.member_sheet))
            .collect()
            .to_numpy()
            .flatten()
            .tolist()
        )

        return member_list

    def get_month_holiday_list(self):
        """取得該月的假日名單

        Returns:
            >>> get_month_holiday_list():
            List[Text]
        """
        sheet = self.gs_url.worksheet_by_title(self.holiday_sheet)
        holiday_list = sheet.get_all_records(numericise_data=False)
        year, year_month = (
            arrow.get(self.update_date).format("YYYY"),
            arrow.get(self.update_date).format("YYYY-MM"),
        )
        month_holiday_list = (
            pl.LazyFrame(holiday_list)
            .select(pl.col(f"{self.holiday_sheet}_{year}"))
            .filter(
                pl.col(f"{self.holiday_sheet}_{year}").str.extract(r"\d{4}-\d{2}", 0)
                == year_month
            )
            .collect()
            .to_numpy()
            .flatten()
            .tolist()
        )

        return month_holiday_list

    def get_month_date_list(self):
        """取得該月的所有日期名單

        Returns:
            >>> get_month_date_list():
            List[Text]
        """
        update_date_list = self.update_date.split("-")
        first_day = arrow.Arrow(int(update_date_list[0]), int(update_date_list[1]), 1)
        last_day = first_day.shift(months=1).shift(days=-1)
        month_date_list = [
            date.format("YYYY-MM-DD")
            for date in arrow.Arrow.range("day", first_day, last_day)
        ]

        return month_date_list

    def add_conditional_formatting(
        self, sheet_updatedate: pygsheets.Worksheet, month_holiday_list: List[Text]
    ):
        """使用條件式格式設定，將假日上顏色

        Args:
            sheet_updatedate (pygsheets.Worksheet): 指定要更新的worksheet
            month_holiday_list (List[Text]): 當月的假日名單
        """
        total_cols, row_count = sheet_updatedate.cols, sheet_updatedate.rows
        column_num_list = [letter for letter in string.ascii_uppercase] + [
            letter1 + letter2
            for letter1 in string.ascii_uppercase
            for letter2 in string.ascii_uppercase
        ]
        conditional_formatting_01, conditional_formatting_02 = (
            column_num_list[0],
            column_num_list[total_cols - 1],
        )
        month_holiday_list = [
            f"A$1=DATE({holiday.replace('-', ',')})" for holiday in month_holiday_list
        ]
        conditional_formatting_formula = "=OR({})".format(", ".join(month_holiday_list))
        sheet_updatedate.add_conditional_formatting(
            f"{conditional_formatting_01}1",
            f"{conditional_formatting_02}{row_count}",
            "CUSTOM_FORMULA",
            {
                "backgroundColor": {
                    "red": self.red_num,
                    "green": self.green_num,
                    "blue": self.blue_num,
                }
            },
            [f"{conditional_formatting_formula}"],
        )

    def make_new_sheet(
        self,
        member_list: List[Text],
        month_holiday_list: List[Text],
        month_date_list: List[Text],
    ):
        """新增新的工作表，或指定工作表，開始進行動作

        Args:
            member_list (List[Text]): 組員名單
            month_holiday_list (List[Text]): 該月的假日名單
            month_date_list (List[Text]): 該月的每一日的日期
        """
        wks_list = self.gs_url.worksheets()
        wks_title_dic = {wks.title: wks.title for wks in wks_list}
        update_date_list = self.update_date.split("-")
        # 查看該工作表是否存在，如不存在則新增並插入title
        new_tab = f"{update_date_list[0]}年{update_date_list[1]}月"
        if not wks_title_dic.get(new_tab):
            month_date_list.insert(0, "組員名單")
            self.gs_url.add_worksheet(
                title=new_tab,
                rows=len(member_list) + 10,
                cols=len(month_date_list) + 10,
            )
            sheet_updatedate = self.gs_url.worksheet_by_title(new_tab)
            sheet_updatedate.update_values(crange="A1", values=[month_date_list])
        else:
            sheet_updatedate = self.gs_url.worksheet_by_title(new_tab)

        # 插入組員名單
        exits_member_list = sheet_updatedate.get_all_records(numericise_data=False)
        # 如果組員名單不存在
        if not exits_member_list:
            member_list = [[member] for member in member_list]
            sheet_updatedate.update_values(crange="A2", values=member_list)
            # 進行條件式格式設定，將假日的部份用條件式格式設定變色
            self.add_conditional_formatting(sheet_updatedate, month_holiday_list)
            # 在「工作日誌列表」的工作表增加hyperlink，以倒序的方式排列
            sheet_updatedate_url = re.search(r"#gid=(\d+)", sheet_updatedate.url).group(
                0
            )
            sheet_datalist = self.gs_url.worksheet_by_title("工作日誌列表")
            value = f'=HYPERLINK("{sheet_updatedate_url}", "{new_tab}")'
            sheet_datalist.insert_rows(row=1, number=1, values=[[value]])
            # 凍結第一列
            sheet_updatedate.frozen_rows = 1
        # 如果組員名單存在，則將組員名單與組員名單的的那個sheet資料進行比較
        else:
            exits_member_list = (
                pl.LazyFrame(exits_member_list)
                .select(pl.col("組員名單"))
                .collect()
                .to_numpy()
                .flatten()
                .tolist()
            )
            not_exist_member_list = list(set(member_list) - set(exits_member_list))
            if not_exist_member_list:
                not_exist_member_list = [[member] for member in not_exist_member_list]
                sheet_updatedate.update_values(
                    crange=f"A{len(exits_member_list)+2}", values=not_exist_member_list
                )

    def run_all(self):
        member_list = self.get_member_list()
        month_holiday_list = self.get_month_holiday_list()
        month_date_list = self.get_month_date_list()
        self.make_new_sheet(member_list, month_holiday_list, month_date_list)

    def main(self):
        try:
            self.run_all()
        except Exception as e:
            print("發生錯誤")
            print(e)
        else:
            print("程序順利執行完成")


if __name__ == "__main__":
    WorkdiaryGooglesheet().main()
