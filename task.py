from pathlib import Path

from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF


class InsertDataAndExportPdf(object):

    def __init__(self):
        self.wd = Selenium()
        self.http = HTTP()
        self.excel = Files()
        self.pdf = PDF()
        self.output_dir = Path(Path.cwd(), 'output')

    def run_task(self):
        try:
            self.open_intranet_website()
            self.log_in()
            self.download_excel_file()
            self.fill_form_with_data_from_excel_file()
            self.collect_results()
            self.export_table_as_pdf()
        finally:
            self.logout_and_close_browser()
        return

    def open_intranet_website(self):
        self.wd.open_available_browser(url='https://robotsparebinindustries.com/')
        return

    def log_in(self):
        self.wd.input_text('username', 'maria')
        self.wd.input_text('password', 'thoushallnotpass')
        self.wd.submit_form()
        return

    def download_excel_file(self):
        self.http.download(url='https://robotsparebinindustries.com/SalesData.xlsx', overwrite=True)
        return

    def fill_form_with_data_from_excel_file(self):
        self.excel.open_workbook('SalesData.xlsx')
        sales_reps = self.excel.read_worksheet(header=True)
        self.excel.close_workbook()
        for person in sales_reps:
            self.fill_and_submit_form_for_one_person(person)
        return

    def fill_and_submit_form_for_one_person(self, person):
        try:
            self.wd.wait_until_page_contains_element('sales-form')
            self.wd.input_text('firstname', person['First Name'])
            self.wd.input_text('lastname', person['Last Name'])
            self.wd.input_text('salesresult', person['Sales'])
            self.wd.select_from_list_by_value('salestarget', str(person['Sales Target']))
            self.wd.submit_form('sales-form')
        except TypeError:
            pass
        return

    def collect_results(self):
        file_path = str(Path(self.output_dir, 'sales_summary.png'))
        self.wd.screenshot('css:div.sales-summary', file_path)
        return

    def export_table_as_pdf(self):
        self.wd.wait_until_element_is_visible('sales-results')
        sales_result_html = self.wd.get_element_attribute('sales-results', 'outerHTML')
        file_path = str(Path(self.output_dir, 'output.pdf'))
        self.pdf.html_to_pdf(sales_result_html, file_path)
        return

    def logout_and_close_browser(self):
        self.wd.click_button('logout')
        self.wd.wait_until_page_does_not_contain_element('logout')
        self.wd.close_browser()
        return


if __name__ == '__main__':
    task = InsertDataAndExportPdf()
    task.run_task()
