"""
OrderRobot for RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
"""
from pathlib import Path

from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive
from RPA.Dialogs import Dialogs


class OrderRobot(object):

    def __init__(self):
        self.wd = Selenium()
        self.http = HTTP()
        self.tables = Tables()
        self.pdf = PDF()
        self.fs = FileSystem()
        self.archive = Archive()
        self.dialogs = Dialogs()
        self.output_path = Path(Path.cwd(), 'output')

    def start_robot(self):
        self.open_robot_order_website()
        try:
            link_for_orders = self.get_download_link_from_user()
            orders = self.get_orders(link_for_orders)
            for order in orders:
                self.close_modal_window()
                self.fill_the_form(order)
                self.submit_the_order()
                pdf = self.save_order_receipt_as_pdf(order['Order number'])
                screenshot = self.save_robot_screenshot_to_pdf()
                self.create_receipt_pdf(pdf, screenshot)
                self.wd.click_button('order-another')
            self.create_archive()
            self.remove_temp_files()
        finally:
            self.wd.close_browser()
        return

    def open_robot_order_website(self):
        self.wd.open_available_browser('https://robotsparebinindustries.com/#/robot-order', headless=True)
        self.wd.wait_until_location_is('https://robotsparebinindustries.com/#/robot-order')
        return

    def close_modal_window(self):
        self.wd.wait_until_page_contains_element('css:div.modal')
        if self.wd.does_page_contain_element('css:.btn-danger'):
            self.wd.click_button_when_visible('css:.btn-danger')
        return

    def get_download_link_from_user(self):
        self.dialogs.add_text_input(
            name='csv_src',
            label='Source for .csv file download',
            placeholder='https://robotsparebinindustries.com/orders.csv'
        )
        self.dialogs.add_text('Use https://robotsparebinindustries.com/orders.csv but it`s a secret')
        submitted_result = self.dialogs.run_dialog()
        link = submitted_result['csv_src']
        return link

    def get_orders(self, link: str):
        self.http.download(url=link, overwrite=True)
        orders = self.tables.read_table_from_csv('orders.csv', header=True)
        return orders

    def fill_the_form(self, order: dict):
        self.wd.select_from_list_by_value('head', order['Head'])
        self.wd.select_radio_button('body', order['Body'])
        self.wd.input_text('css:input[type="number"]', order['Legs'])
        self.wd.input_text('address', order['Address'])
        self.wd.click_button('preview')
        return

    def submit_the_order(self):
        self.wd.click_button('order')
        while self.wd.does_page_contain_element('css:div.alert-danger'):
            self.wd.click_button('order')
        return

    def save_order_receipt_as_pdf(self, order_number: str):
        receipt_html = self.wd.get_element_attribute('receipt', 'outerHTML')
        pdf_file_path = str(Path(self.output_path, 'receipts', f'{order_number}.pdf'))
        self.pdf.html_to_pdf(receipt_html, pdf_file_path)
        return pdf_file_path

    def save_robot_screenshot_to_pdf(self):
        screenshot_path = str(Path(self.output_path, 'tmp', 'robot_screenshot.png'))
        self.wd.screenshot('robot-preview-image', screenshot_path)
        return screenshot_path

    def create_receipt_pdf(self, pdf: str, screenshot: str):
        self.pdf.add_files_to_pdf([screenshot], pdf, append=True)
        return

    def create_archive(self):
        receipts_path = str(Path(self.output_path, 'receipts'))
        self.archive.archive_folder_with_zip(receipts_path, str(Path(self.output_path, 'receipts.zip')))
        return

    def remove_temp_files(self):
        receipts_path = str(Path(self.output_path, 'receipts'))
        tmp_path = str(Path(self.output_path, 'tmp'))
        self.fs.remove_directory(receipts_path, recursive=True)
        self.fs.remove_directory(tmp_path, recursive=True)
        return


if __name__ == '__main__':
    robot = OrderRobot()
    robot.start_robot()
