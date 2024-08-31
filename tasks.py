from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from pathlib import Path
from PIL import Image
import os
import shutil

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    browser.configure(
        slowmo=100,
    )

    open_robot_order_website()
    close_annoying_modal()
    orders = get_orders()
    output_folder_path = 'output/receipts/'
    for order in orders:
        try:
            
            fill_the_form(order)
            store_receipt_as_pdf(order['Order number'])
            screenshot_robot(order['Order number'])
            embed_screenshot_to_receipt(output_folder_path+order['Order number']+".png", output_folder_path+order['Order number']+".pdf")
            do_another_order()
            close_annoying_modal()
            archive_receipts()

        except:
            
            continue
        


def open_robot_order_website():
    """ Launches a browser and Navigates to orders page """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def get_orders():
    """ Gets the orders.csv file """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
    """ Gets the data from excel and returns it """
    tables = Tables()
    orders = tables.read_table_from_csv(path="orders.csv", header=True)
    return orders


def close_annoying_modal():
    page = browser.page()
    page.locator("xpath=//button[contains(.,'OK')]").click()
    

def fill_the_form(order):
    page = browser.page()
    page.select_option("#head", str(order['Head']))
    page.locator("#id-body-"+order['Body']).click()
    page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").fill(order['Legs'])
    page.locator("#address").fill(order['Address'])
    page.locator("#preview").click()
    page.locator("xpath=//p[contains(.,'Admire your robot!')]").wait_for()
    page.locator("#order").click()

    page.wait_for_selector("#order-another")
    


def store_receipt_as_pdf(order_number):
    pdf = PDF()
    page = browser.page()

    receipt = page.locator("#order-completion").inner_html()

    folder_path = 'output/receipts'
    if not(os.path.exists(folder_path) and os.path.isdir(folder_path)):
        Path(folder_path).mkdir(parents=True, exist_ok=True)

    pdf.html_to_pdf(receipt+"<br><br><br>", folder_path + "/" + str(order_number) + ".pdf")
    

def screenshot_robot(order_number):
    image_path = 'output/receipts/'+order_number+".png"
    pdf = PDF()
    page = browser.page()
    page.screenshot(path=image_path)

    

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot, source_path=pdf_file, output_path=pdf_file, coverage=0.25)

def do_another_order():
    page = browser.page()
    page.click("button#order-another")

def archive_receipts():
    archive_path = shutil.make_archive("output", "zip", "output/receipts")
    shutil.rmtree("output/receipts")