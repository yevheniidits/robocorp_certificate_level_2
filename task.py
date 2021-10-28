from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import XlsxWorkbook
from RPA.PDF import PDF
import os
import time
import re


excel_book = XlsxWorkbook('output/agencies.xlsx')
excel_book.create()


class PDFLib(object):

    def __init__(self, pdf_obj):
        self.pdf = pdf_obj

    def pdf_data(self, name: str) -> list:
        text_pdf = self.pdf.get_text_from_pdf('output/' + name)
        text = text_pdf[1].replace('\n', ' ')
        name_inv_re = re.search(r'Name of this Investment: (.*)2\. Unique', text)
        name_inv = name_inv_re.group(1) if name_inv_re else ''
        uii_re = re.search(r'Unique Investment Identifier \(UII\): (.*)Section B', text)
        uii = uii_re.group(1) if uii_re else ''
        return [uii, name_inv]


class Driver(object):

    def __init__(self, browser):
        self.browser = browser

    def wait_elements(self, css_selector: str, timeout: int = 10):
        """Wait and find elements with css selector.
        Wait timeout(def 10s)"""
        self.browser.wait_until_element_is_visible(f'css:{css_selector}', timeout)
        return self.browser.find_elements(f'css:{css_selector}')

    def parse_info(self):
        self.wait_elements('.dataTables_scrollHeadInner th')

        self.browser.select_from_list_by_value('css:select[name=investments-table-object_length]', '-1')
        self.browser.wait_until_element_is_enabled('css:.paginate_button.next.disabled', 20)
        all_records = self.wait_elements('#investments-table-object tbody tr', 60)

        details_investments = []
        excel_book.create_worksheet('investments')
        for tmp_record in all_records:
            _record = tmp_record.find_elements_by_css_selector('td')
            record = [i.text for i in _record]
            details_investments.append(record)
        excel_book.append_worksheet('investments', details_investments)

        selector_links = self.browser.find_elements('css:#investments-table-object td a')
        all_links = [i.get_attribute('href') for i in selector_links]
        for link in all_links:
            self.browser.go_to(link)
            self.browser.wait_until_element_is_visible(f'css:#business-case-pdf a', 30)
            self.browser.click_element('css:#business-case-pdf a')
            self.browser.wait_until_element_is_not_visible('css:#business-case-pdf span', 60)
            time.sleep(1)


def main():
    name_agencies = 'Department of Education'
    browser = Selenium()
    browser.set_download_directory(rf"{os.getcwd()}/output")
    web_driver = Driver(browser)
    try:
        browser.open_available_browser("http://itdashboard.gov/")
        browser.click_element(f'css:.tuck-7 .trend_sans_oneregular.btn-default.btn-lg-2x')
        agency_elements = web_driver.wait_elements('#agency-tiles-widget .tuck-5')
        excel_book.rename_worksheet('agencies')

        all_agencies = []
        for agency in agency_elements:
            agency_title = agency.find_element_by_css_selector('.h4.w200').text
            agency_amount = agency.find_element_by_css_selector('.h1.w900').text
            all_agencies.append([agency_title, agency_amount])

        excel_book.append_worksheet('agencies', all_agencies)

        for agency in agency_elements:
            title = agency.find_element_by_css_selector('.h4.w200').text
            if title == name_agencies:
                agency.find_element_by_css_selector('.col-sm-12 a').click()
                web_driver.parse_info()
                break

        excel_book.save()

        # compare pdf and excel files
        pdf = PDF()
        pdf_lib = PDFLib(pdf)
        all_investments = {}
        investments_data = excel_book.read_worksheet('investments')
        for investment in investments_data:
            if investment.get('A'):
                all_investments[investment['A']] = investment.get('C')

        all_pdf_files = []

        for f in os.listdir('output'):
            name, ext = os.path.splitext(f)
            if ext == '.pdf':
                all_pdf_files.append(f)

        for pdf in all_pdf_files:
            data = pdf_lib.pdf_data(pdf)
            if data[0] and data[1] == all_investments.get(data[0]):
                print(f'is true:\n{data[0]} - {data[1]}')
            else:
                print(
                    f"Does not match"
                    f"PDF:\nUII: {data[0]}\nInvestments name:{data[1]}\n\n"
                    f"Excel:\nInvestments name: {all_investments.get(data[0])}"
                )

    finally:
        browser.close_all_browsers()
        excel_book.close()


if __name__ == "__main__":
    main()