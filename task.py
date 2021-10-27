from RPA.Browser.Selenium import Selenium


class InsertDataAndExportPdf(object):

    def __init__(self):
        self.wd = Selenium()

    def open_intranet_website(self):
        self.wd.open_available_browser(url='https://robotsparebinindustries.com/')
        print(self.wd.get_cookies())
        return


if __name__ == '__main__':
    task = InsertDataAndExportPdf()
    task.open_intranet_website()
