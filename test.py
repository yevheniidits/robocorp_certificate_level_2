from RPA.Browser.Selenium import Selenium
from time import sleep


def main():
    wd = Selenium()
    wd.set_screenshot_directory('output_test')
    wd.open_available_browser(url='https://google.com', headless=False)
    sleep(2)
    wd.capture_page_screenshot()
    wd.close_browser()
    return


if __name__ == '__main__':
    main()
