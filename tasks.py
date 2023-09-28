from robocorp.tasks import task
from robocorp import browser, http, log
from RPA.Tables import Tables
import time
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)
    open_robot_order_website()
    close_annoying_modal()
    orders = get_orders()
    order_table = open_the_file_as_table()
    loop_over_table(order_table)
    archive_receipts()
    
    #below functions nested in others
    #store_receipt_as_pdf()
    #embed_screenshot_to_receipt(screenshot = r"output/receipts/1.png", pdf_file = r"output/receipts/1.pdf")

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def get_orders():
    """Downloads excel file from the given URL"""
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True,target_file="orders.csv")

def open_the_file_as_table():
    tables = Tables()
    table = tables.read_table_from_csv(path="orders.csv",columns=['Order number', 'Head', 'Body', 'Legs', 'Address'])    
    return table

def loop_over_table(table):
    for row in table:
        log.info(row)
        fill_the_form(row=row)
        
def close_annoying_modal():
    # APIs in robocorp.browser return the same browser instance, which is
    # automatically closed when the task finishes.
    page = browser.page() 
    page.click('text=OK')

def fill_the_form(row):
    #enter each of the below data points in the form
    page = browser.page()

    #Head
    page.select_option("select#head", value=row['Head'])
    
    #logic to select right radio button for body
    body_id = f"#id-body-{row['Body']}"
    page.click(body_id)
    
    #Legs
    page.fill('.form-control[type="number"]', row['Legs'])

    #Address
    page.fill('#address', row['Address'])
    
    #Order     
    page.click("#order")
    #check for error. Re-click if needed.
    error_order()
    
    #store order receipts as a pdf file.
    store_receipt_as_pdf(order_number=row['Order number'])
    
    #take screenshot
    screenshot_robot(order_number=row['Order number'])
    
    #embed screenshot to receipt
    embed_screenshot_to_receipt(screenshot=f"output/receipts/{row['Order number']}.png",
                                pdf_file=f"output/receipts/{row['Order number']}.pdf")
    

    #click on order another
    page.click('#order-another')
    #check for error. Re-click if needed.
    error_order_another()
    
    #call close to close pop-up
    close_annoying_modal()
    
    
def store_receipt_as_pdf(order_number):
    #extract div with id "Receipt" and store to a var
    page = browser.page()
    receipt_exists = page.query_selector('div#order-completion')
    if receipt_exists:    
        try:
            receipt = page.query_selector('div#order-completion')
            receipt_html =  receipt.evaluate('el => el.outerHTML')
        except:
            log.info("couldn't find receipt")
            time.sleep(2)
    
    #var to pdf
    pdf = PDF()
    try:
        pdf.html_to_pdf(receipt_html, f"output/receipts/{order_number}.pdf")
    except:
        log.info("could not find receipt")
    
def screenshot_robot(order_number):
    page = browser.page()
    page.screenshot(path=f"output/receipts/{order_number}.png",full_page=True)
    
    
def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf_lib = PDF()
    screenshot = [screenshot]
    pdf_lib.add_files_to_pdf(files = screenshot, target_document = pdf_file, append = True)


def error_order():
    #check for alert
    page = browser.page()
    element = page.query_selector('.alert.alert-danger')

    while element:
        log.info("Element with class 'alert alert-danger' is present on the page.")
        page.click("#order")
        # Recheck for the element after the click
        element = page.query_selector('.alert.alert-danger')
        

def error_order_another():
    #check for alert
    page = browser.page()
    element = page.query_selector('.alert.alert-danger')
    while element:
        log.info("Element with class 'alert alert-danger' is present on the page.")
        page.click("#order-another")
        # Recheck for the element after the click
        element = page.query_selector('.alert.alert-danger')
        
def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('./output/receipts', "receipts.zip")