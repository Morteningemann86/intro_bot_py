from robocorp import browser
from robocorp.tasks import task

from RPA.Excel.Files import Files as Excel
from RPA.HTTP import HTTP
import os

dir_target_file = os.path.join(os.getcwd(), 'downloads', 'customers.xlsx')
dir_receipts = os.path.join(os.getcwd(), 'receipts')
url_rpachallenge = "http://rpachallenge.com/"
url_challenge_file = "http://rpachallenge.com/assets/downloadFiles/challenge.xlsx"

@task
def add_customers_from_list():
    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
    )

    HTTP().download(url_challenge_file, dir_target_file)

    excel = Excel()
    excel.open_workbook(dir_target_file)
    rows = excel.read_worksheet_as_table("Sheet1", header=True)

    page = browser.goto(url_rpachallenge)
    page.click("button:text('Start')")

    for row in rows:
        fill_and_submit_form(row)

    element = page.locator("css=div.congratulations")
    browser.screenshot(element)


def fill_and_submit_form(row):
    page = browser.page()
    page.fill("//input[@ng-reflect-name='labelFirstName']", str(row["First Name"]))
    page.fill("//input[@ng-reflect-name='labelLastName']", str(row["Last Name"]))
    page.fill("//input[@ng-reflect-name='labelCompanyName']", str(row["Company Name"]))
    page.fill("//input[@ng-reflect-name='labelRole']", str(row["Role in Company"]))
    page.fill("//input[@ng-reflect-name='labelAddress']", str(row["Address"]))
    page.fill("//input[@ng-reflect-name='labelEmail']", str(row["Email"]))
    page.fill("//input[@ng-reflect-name='labelPhone']", str(row["Phone Number"]))
    page.click("input:text('Submit')")
