# +
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

class Title:

    def __init__(self):
        self.browser = Selenium()
        self.url_link = "https://robotsparebinindustries.com/"
        self.open_browser = self.browser.open_available_browser(self.url_link)
        self.http = HTTP()
        self.files = Files()
        self.pdf = PDF()
        self.vari = []

    def log_in(self):
        self.browser.input_text('id:username', 'maria')
        self.browser.input_text('password', 'thoushallnotpass')
        self.browser.submit_form()
        self.browser.wait_until_page_contains_element('id:sales-form')

    def fill_form(self):
        for var in self.vari:
            self.browser.input_text('firstname', f'{var["First Name"]}')
            self.browser.input_text('lastname', f'{var["Last Name"]}')
            self.browser.input_text('salesresult', f'{var["Sales"]}')
            self.browser.select_from_list_by_value('salestarget', f'{var["Sales Target"]}')
            self.browser.click_button('Submit')
        self.collect_data()
        self.pdf_file()

    def download_excel_file(self):
        self.http.download('https://robotsparebinindustries.com/SalesData.xlsx')

    def collect_data(self):
        while True:
            try:
                self.browser.screenshot('css:div.sales-summary', "output/sales_summary.png")
                break
            except:
                pass
    def pdf_file(self):
        html = self.browser.get_element_attribute('id:sales-results', 'outerHTML')
#         ele = self.browser.get_element_attribute('id:sales_results')
        self.pdf.html_to_pdf(html, 'output/sales_result.pdf')

    def workbook(self):
        self.files.open_workbook('SalesData.xlsx')
        var = self.files.read_worksheet_as_table(header=True)
        self.files.close_workbook()
        for ele in var:
            self.vari.append(ele)

    def close_browser(self):
        self.browser.click_button('logout')
        self.browser.close_browser


# -

if __name__ == "__main__":
    title_obj = Title()
    title_obj.log_in()
    title_obj.download_excel_file()
    title_obj.workbook()
    title_obj.fill_form()
    title_obj.close_browser()
