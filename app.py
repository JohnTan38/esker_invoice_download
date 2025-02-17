from selenium import webdriver # 1 login 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
import pyautogui
import pandas as pd
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
import fitz
import glob, os, os.path
import pathlib
from pathlib import Path
import shutil

#get user input
import tkinter as tk

def get_username():
    root = tk.Tk()
    root.title("Esker Invoice Download")
    root.geometry("350x200")

    name_var = tk.StringVar()
    username=None

    def submit_username():
        nonlocal username
        username = name_var.get()
        print("Username:", username)  # Print username within the function
        root.destroy()  # Close the window after submission
        #return name  # Return the username

    name_label = tk.Label(root, text='Username', font=('calibre', 10, 'bold'))
    name_entry = tk.Entry(root, textvariable=name_var, font=('calibre', 10, 'normal'))

    submit_btn = tk.Button(root, text='Submit', command=submit_username)

    name_label.grid(row=0, column=0, padx=5, pady=5)
    name_entry.grid(row=0, column=1, padx=5, pady=5)
    submit_btn.grid(row=1, column=1, padx=5, pady=5)
    root.mainloop()  # Start the Tkinter event loop and store the returned value
    return username  # Return the username

global username
username=get_username()  # Call the function to get the username


driver = webdriver.Chrome()
driver.get("https://az3.ondemand.esker.com/ondemand/webaccess/asf/home.aspx")
driver.maximize_window()
time.sleep(1)

driver.find_element(By.XPATH, '//*[@id="ctl03_tbUser"]').send_keys("john.tan@sh-cogent.com.sg")
driver.find_element(By.XPATH, '//*[@id="ctl03_tbPassword"]').send_keys("PASSWORD")
driver.find_element(By.XPATH, '//*[@id="ctl03_btnSubmitLogin"]').click()
time.sleep(5) # login

path_page = r"C:/Users/"+ username + r"/Documents/esker_merged/"
df_page = pd.read_excel(path_page+ 'page.xlsx', sheet_name='page', engine='openpyxl')
pg = df_page.at[0, 'page']
pg_max = df_page.at[0, 'page_total']

def hover(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover.perform()

time.sleep(2)
x_path_hover = '//*[@id="mainMenuBar"]/td/table/tbody/tr/td[36]/a/div' #arrow
hover(driver, x_path_hover)
time.sleep(2)

try:
    #drop_down=driver.find_element(By.XPATH, '//*[@id="CUSTOMTABLE_TAB_100872176"]"]/a/div[2]').click()
    tables=driver.find_element(By.XPATH, '//*[@id="DOCUMENT_TAB_100872215"]').click()
    time.sleep(1)
except Exception as e:
    print(e) #VENDOR INVOICES (SUMMARY) #TABLES

pyautogui.moveTo(180,320, duration=1.5) #06 - Paid
pyautogui.click(button='left')
pyautogui.write('05 - Pending payment', interval=0.25)
pyautogui.press('enter')
time.sleep(1)

# if pg!=0 then page right
def main():
    path_page = r"C:/Users/"+ username + r"/Documents/esker_merged/"
    df_page = pd.read_excel(path_page+ 'page.xlsx', sheet_name='page', engine='openpyxl')
    global pg
    pg = df_page.at[0, 'page']
    pg_max = df_page.at[0, 'page_total']
    if pg > pg_max:
        print('max page reached')
        quit()
    else:
        print('continue rest of code')
        # Rest of the code will execute only if pg <= pg_max
        if pg !=0:
            print(f'click page right {pg} time')                          
            
            #btn_pg_right = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_NextPageLinkButtonTopBar"]')
            pyautogui.moveTo(1315,530, duration=1) #505
            for _ in range(pg):
                try:    
                    #btn_pg_right.click()
                    pyautogui.click(button='left')
                    time.sleep(1.5)
                except Exception as e:
                    print(f'button page right not click {e}')

            invoice_download_code()
        elif pg ==0:
            print(f'click page right {pg}')

            invoice_download_code()


def invoice_download_code():
    def get_index(n):
        """
        Returns the index based on the given value of n.
        Args:
        n: The input value.
        Returns:
        The corresponding index based on the following mapping:
            - n=1: index=1
            - For other values of n: index=n+2
        """
        if n == 1:
            return 1
        else:
            return 2* (n - 1) + 1

    def find_common_invoice_numbers(list_invoice_numbr, list_invoice_numbr_on_page):
        """
        Finds the common invoice numbers between two lists.

        Args:
        list_invoice_numbr: The first list of invoice numbers.
        list_invoice_numbr_on_page: The second list of invoice numbers.

        Returns:
        A list containing the common invoice numbers.
        """
        return list(set(list_invoice_numbr) & set(list_invoice_numbr_on_page))
    
    def get_invoice_number_indices(list_invoice_numbr_on_page, list_invoice_numbr):
        """
        Finds the indices of the given invoice numbers within a list of invoice numbers on a page.
        Args:
        list_invoice_numbr_on_page: A list of invoice numbers on the page.
        list_invoice_numbr: A list of invoice numbers to find the indices for.
        Returns:
        A dictionary where keys are invoice numbers and values are their corresponding indices 
        in the list_invoice_numbr_on_page.
        """
        index_dict = {}
        for invoice in list_invoice_numbr:
            try:
                    index_dict[invoice] = list_invoice_numbr_on_page.index(invoice)
            except ValueError:
                    index_dict[invoice] = -1  # Indicate that the invoice number was not found
        return index_dict

    import os.path
    from datetime import datetime

    date_today = datetime.now().strftime("%Y%m%d")
    file_log = 'esker_log_'+date_today+'.txt'
    path_invoice = "C:/Users/"+ username+ "/Downloads/" #
    path_log = path_invoice + file_log
    if os.path.exists(path_log) == False:
        open(path_log, "w").close #check log file exists, if not create

    import logging
    # Configure logging (do this once at the beginning of your script)
    logging.basicConfig(filename=path_log, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def write_log(log_message):
        if isinstance(log_message, list):
            for item in log_message:
                logging.info(item) # Log each item individually
        else:
            logging.info(log_message) # Log the message as is

    def get_pdf_files_with_invoice_number(path_pdf, invoice_numbr):
        """
        Gets a list of PDF files in the given directory that contain the specified invoice number in their filename.
        Args:
            path_pdf: The path to the directory containing the PDF files.
            invoice_numbr: The invoice number to search for.
        Returns:
            A list of PDF filenames that contain the invoice number.
        """
        if '/' in invoice_numbr:
            invoice_numbr = invoice_numbr.replace('/','_')
        try:
            pdf_files = [f for f in os.listdir(path_pdf) if invoice_numbr in f and f.endswith(".pdf")]
        except Exception as e:
            print(e)
        return pdf_files
    
    def merge_pdfs(pdf_files, path_pdf):
        """
        Merges multiple PDF files into a single PDF.
        Args:
            pdf_files: A list of PDF filenames.
            path_pdf: The path to the directory containing the PDF files.
        Returns:
            The merged PDF writer object.
        """
        pdf_writer = PdfWriter()

        for pdf in pdf_files:
            try:
                with open(path_pdf + pdf, 'rb') as pdf_file:
                    pdf_reader = PdfReader(pdf_file)
                    for page_num in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[page_num]
                        pdf_writer.add_page(page)
            except EOFError as e:
                print(f"Error reading PDF file '{pdf}': {e}")
                continue  # Skip the file that caused the error
        return pdf_writer

    def merge_xml_to_pdf(path_pdf, invoice_numbr):
        """
        Merges the XML/HTML file with the specified invoice number into the PDF file with the same invoice number.
        Args:
            path_pdf: The path to the directory containing the PDF and XML files.
            invoice_numbr: The invoice number to search for.
        """
        namedoc = invoice_numbr + "_merged.pdf"
        pathnamedoc = os.path.join(path_pdf, namedoc)
        print(pathnamedoc)

        doc = fitz.open(pathnamedoc)  # open main document
        count = doc.embfile_count()
        #print(f"number of embedded file {count}")  # shows number of embedded files
        namedata = "html_" + invoice_numbr + ".html"
        pathnamedata = os.path.join(path_pdf, namedata)
        print(pathnamedata)

        try:

            embedded_doc = fitz.open(pathnamedata)  # open document you want to embed
            embedded_data = pathlib.Path(pathnamedata).read_bytes()  # get the document byte data as a buffer
            doc.embfile_add(namedata, embedded_data)
            doc.saveIncr()
            print(f"XML/HTML file merged with {invoice_numbr} PDF successfully.")
        except Exception as e:
            print(e)
    
    def success_html():
        try:
            success=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_MainGrid"]/tbody/tr[2]/td[7]/a')#.click()
            if success.get_attribute('innerHTML') != 'Success':
                time.sleep(3)
                pyautogui.moveTo(370,120, duration=1.5)
                pyautogui.click(button='left')
                pyautogui.hotkey('ctrl', 'shift','r')
                time.sleep(2)
                pyautogui.hotkey('ctrl', 'shift','r') #refresh page 2 times
                print('processing Success...')
                time.sleep(2)
        except Exception as e:
                print(e)
        """    
        if success.get_attribute('innerHTML') != 'Success':
            time.sleep(3)
            pyautogui.moveTo(370,120, duration=1.5)
            pyautogui.click(button='left')
            pyautogui.hotkey('ctrl', 'shift','r')
            print('processing Success...')
            time.sleep(2)
        """
        try:
            success = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_MainGrid"]/tbody/tr[2]/td[7]/a')
            success.click()
        except Exception as e:
            print(e)
        time.sleep(3)

        def hover(driver, elem_to_hover):
            #elem_to_hover = driver.find_element(By.XPATH, x_path)
            elem_to_hover = pyautogui.moveTo(30,55, duration=2)
            hover = ActionChains(driver).move_to_element(elem_to_hover)
            hover.perform()
        try:

            elem_to_hover = pyautogui.moveTo(30,55, duration=2)
            hover(driver, elem_to_hover)
            print('hovering...')
            time.sleep(2)
        except Exception as e:
            print(e)
    
            pyautogui.moveTo(30,65, duration=1.5) #back arrow previous webpage
            time.sleep(1)
            pyautogui.click(button='left')
        
        time.sleep(5)
        try:
             
            html_2 = driver.find_element(By.XPATH, '//*[@id="ZipPane_eskCtrlBorder_content"]/div/div[2]/div[2]/div[3]/div[1]/div/span[1]')
            if html_2.get_attribute("innerHTML") =='2':
                    html_2.click() #download only details.html
                    time.sleep(3)
        except Exception as e:
            print(e)
            pyautogui.moveTo(330,290, duration=1.5)
            pyautogui.click(button='left')
            time.sleep(1.5)
        html_quit = driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a/span')
        html_quit.click()
        time.sleep(1) #pyautogui.moveTo(330,290, duration=2)

    # define create zip file
    def create_zip_file():
        time.sleep(0.5)
        try:
                checkbox = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_MainGrid_ctl03_MultiAction"]')
                checkbox.click()
                time.sleep(0.5)
        except Exception as e:
                pyautogui.moveTo(20,610, duration=1.5) #checkbox
                pyautogui.click(button='left')
                print(f' checkbox click with mouse movement {e}')
                time.sleep(0.5)
        try:
             
            download_docu = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_adminListActionsBar"]/div/div[1]/button[2]')
            download_docu.click()
            time.sleep(2.5)
        except Exception as e:
            pyautogui.moveTo(160,540, duration=1.5) #Download documents button
            pyautogui.click(button='left')
            print(e)
            time.sleep(3)
        try:
            audit_checkbox=driver.find_element(By.XPATH, '//*[@id="ZipDetailsPane_eskCtrlBorder_content"]/div/div/table/tbody/tr[4]/td[2]/div/label/input')
            audit_checkbox.click()
            time.sleep(1.5)
        except Exception as e:
            pyautogui.moveTo(600,390, duration=1.5) #audit checkbox
            pyautogui.click(button='left')
            print(e)
        try:
            create_zip = driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a[1]/span')
            create_zip.click()
            time.sleep(3)
        except Exception as e:
            pyautogui.moveTo(255,685, duration=1.5) # #create zip file button
            pyautogui.click(button='left')
            time.sleep(2)
            print(e)
        time.sleep(3)

        try:
            success_html()
        except Exception as e:
            print(e)
            """
            download_icon = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tpl_ih_adminList_MainGrid_ctl03_rcsa_a_firstAttach"]')))
            download_icon.click()
        except Exception as e:
            print(e)
            """
        time.sleep(6) #wait for zip file to download
        
        try:
            pyautogui.moveTo(25,60, duration=1.5)
            pyautogui.click(button='left') #back to previous page after zip download
            time.sleep(1)
        except Exception as e:
            print(e)
        try:
            btn_quit_after_download_zip = driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a/span') #Quit btn
            btn_quit_after_download_zip.click()
        except Exception as e:
            print(e)

    #extract from zip file and save as pdf
    import pdfkit
    import zipfile #glob,os

    config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')
    global new_path_name

    def extract_zip_to_pdf(invoice_numbr):
        out_path = os.path.splitext(file)[0]
        file_name = os.path.abspath(file) #full file path    
    
        fh = open(file_name, "rb")
        zip_ref = zipfile.ZipFile(fh)
        name = 'details.html'
        invoice_numbr = invoice_numbr #'132822841'
        zip_ref.extract(name, out_path)
        #parent_zip = os.path.basename(os.path.dirname(outpath)) + ".zip"
        try:
            new_file_name = os.path.splitext(os.path.basename(name))[0] + '_'+invoice_numbr
            
            new_path_name = os.path.dirname(out_path) + os.sep + new_file_name# + "_" + invoice_numbr
            print(f"extracted to {new_path_name}")
            os.rename(out_path, new_path_name)
        except KeyError:
            {}

        for file in glob.glob(new_path_name+'/*.html'):
            pdfkit.from_file(file, file[:-4]+'.pdf', configuration=config) #save html as pdf


    # define function to rename pdf file
    def rename_pdf_in_dir(file_path, file_name):
        """
        Renames a PDF file in the specified directory.
        Args:
            file_path (str): The full path to the directory containing the PDF file.
            file_name (str): The new name for the PDF file (without the .pdf extension).
        Returns:
            str: A message indicating success or failure.
        """
        try:
            # Check if the directory exists
            if not os.path.exists(file_path):
                return f"Error: The directory '{file_path}' does not exist."

            # Find the first PDF file in the directory
            pdf_files = [f for f in os.listdir(file_path) if f.lower().endswith('.pdf')]
            if not pdf_files:
                return f"Error: No PDF files found in the directory '{file_path}'."

            # Get the first PDF file (assuming there's only one PDF file to rename)
            old_file_name = pdf_files[0]
            old_file_path = os.path.join(file_path, old_file_name)

            # Create the new file name with .pdf extension
            new_file_name = f"{file_name}.pdf"
            new_file_path = os.path.join(file_path, new_file_name)

            # Check if a file with the new name already exists
            if os.path.exists(new_file_path):
                return f"Error: A file with the name '{new_file_name}' already exists in '{file_path}'."

            # Rename the file
            os.rename(old_file_path, new_file_path)
            return f"Success: File '{old_file_name}' renamed to '{new_file_name}' in '{file_path}'."

        except Exception as e:
            return f"Error: An unexpected error occurred - {str(e)}"

    def rename_pdf_files_by_name(file_path, invoice_number):
        if '/' in invoice_number:
            invoice_number = invoice_number.replace('/','_')
        # Iterate over all files in the given directory
        for filename in os.listdir(file_path):
            # Check if the file is a PDF and contains 'details' in its name
            if filename.endswith('.pdf') and 'details' in filename:
                new_filename = f"{invoice_number}_0.pdf" # Create the new filename 
                old_file = os.path.join(file_path, filename)
                new_file = os.path.join(file_path, new_filename) # Construct full file paths
                os.rename(old_file, new_file) # Rename the file
                print(f"Renamed '{filename}' to '{new_filename}'")


    # define function move pdf
    import shutil
    def move_pdf(src_dir, dest_dir):
        """
        Moves a PDF file from the source directory to the destination directory.
        Args:
            src_dir (str): The full path to the source directory containing the PDF file.
            dest_dir (str): The full path to the destination directory where the PDF file will be moved.
        Returns:
            str: A message indicating success or failure.
        """
        try:
            # Check if the source directory exists
            if not os.path.exists(src_dir):
                return f"Error: Source directory '{src_dir}' does not exist."

            # Check if the destination directory exists; if not, create it
            if not os.path.exists(dest_dir):
                os.makedirs(dest_dir)  # Create the destination directory if it doesn't exist
                print(f"Info: Destination directory '{dest_dir}' created.")

            # Find the first PDF file in the source directory
            pdf_files = [f for f in os.listdir(src_dir) if (f.lower().endswith('.pdf') and 'merged' in f.lower())]
            if not pdf_files:
                return f"Error: No PDF files found in the source directory '{src_dir}'."

            # Get the first PDF file (assuming there's only one PDF file to move)
            for pdf_file in pdf_files:
                pdf_file_name = pdf_files[0]
                src_file_path = os.path.join(src_dir, pdf_file_name)
                dest_file_path = os.path.join(dest_dir, pdf_file_name)

                # Check if a file with the same name already exists in the destination directory
                if os.path.exists(dest_file_path):
                    return f"Error: A file with the name '{pdf_file_name}' already exists in '{dest_dir}'."

                # Copy and Move the file
                shutil.copy(src_file_path, dest_file_path)
                shutil.move(src_file_path, dest_file_path)
                return f"Success: File '{pdf_file_name}' moved from '{src_dir}' to '{dest_dir}'."

        except PermissionError:
            return f"Error: Permission denied. Check if you have the necessary permissions to access '{src_dir}' or '{dest_dir}'."
        except FileNotFoundError:
            return f"Error: The file '{pdf_file_name}' was not found in '{src_dir}'."
        except Exception as e:
            return f"Error: An unexpected error occurred - {str(e)}"

    def reset_input_invoice_numbr(invoice_numbr):
        try:

                btn_reset=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_displayedFilters_resetBtn"]/span[2]/span[2]')
                time.sleep(0.5)
                if btn_reset.get_attribute("innerHTML") == 'Reset':
                    btn_reset.click()
        except Exception as e:
                print(e)
        time.sleep(0.5)    
        input_invoice_numbr=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_displayedFilters_ctl00_ctl05_ddl1_tagify"]/span')
        input_invoice_numbr.send_keys(invoice_numbr)
        input_invoice_numbr.send_keys(Keys.ENTER)
    
    def remove_html_files_by_name(file_path):
        # Iterate over all files in the given directory
        for filename in os.listdir(file_path):
            # Check if the file is a html and contains 'details' in its name
            if filename.endswith('.html') and 'details' in filename:
                
                try:
                    os.remove(file_path + filename)
                except Exception as e:
                    print(e)
                #print(f"Deleted html file '{filename}'.")
    
    def remove_zip_files(file_path):
            # Iterate over all files in the given directory
            for filename in os.listdir(file_path):
                # Check if the file is a zip and contains 'details' in its name
                if filename.endswith('.zip'):
                    try:
                        os.remove(file_path + filename)
                    except Exception as e:
                        print(e)
                    #print(f"Deleted zip file '{filename}'.")

    write_log((f'Process started at {datetime.now().strftime("%Y%m%d %H:%m")}'))
    #write_log(list_common_invoice_numbr)

    list_invoice_downloaded = []

    list_invoice_numbr_on_page=[]
    dct_invoice_indices_all ={}
    
    path_invoice = r'C:/Users/'+username+'/Downloads/'
    path_pdf = path_invoice
    df_invoice_user_input = pd.read_excel(path_invoice+'invoice.xlsx', sheet_name='invoice', engine='openpyxl')
    list_invoice_numbr_original = df_invoice_user_input['invoice'].tolist()
    # convert all elements in list to str
    list_invoice_numbr = [str(elem) for elem in list_invoice_numbr_original]
    #list_invoice_numbr = ['132822841'] ##user input
    for invoice_numbr in list_invoice_numbr:
            reset_input_invoice_numbr(invoice_numbr)
            

        #for i in range(2):
            i = 0 #line 1
            idx_invoice_numbr = str(i+2)
            invoice_numbr_on_page = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_MainGrid"]/tbody/tr['+idx_invoice_numbr+']/td[7]/a')
            print(invoice_numbr_on_page.get_attribute("innerHTML"))
            list_invoice_numbr_on_page.append(str(invoice_numbr_on_page.get_attribute("innerHTML")))
        #print(list_invoice_numbr_on_page)  
    
            list_common_invoice_numbr = find_common_invoice_numbers(list_invoice_numbr, list_invoice_numbr_on_page)
        #print(list_common_invoice_numbr)
            dct_invoice_indices = get_invoice_number_indices(list_invoice_numbr_on_page, list_invoice_numbr)
            #print(dct_invoice_indices)
        # append to dictionary
            dct_invoice_indices_all.update(dct_invoice_indices)
    print(f"dict all invoice indices {dct_invoice_indices_all}")

    def reset_values_to_zero(input_dict):
        # Create a new dictionary with the same keys but all values set to 0
        return {key: 0 for key in input_dict}

  
    #get dictionary of invoice number, index
    def get_invoice_number_indices(list_invoice_numbr_on_page, list_invoice_numbr):
        """
        Finds the indices of the given invoice numbers within a list of invoice numbers on a page.
        Args:
        list_invoice_numbr_on_page: A list of invoice numbers on the page.
        list_invoice_numbr: A list of invoice numbers to find the indices for.
        Returns:
        A dictionary where keys are invoice numbers and values are their corresponding indices 
        in the list_invoice_numbr_on_page.
        """
        index_dict = {}
        for invoice in list_invoice_numbr:
                try:
                    index_dict[invoice] = list_invoice_numbr_on_page.index(invoice)
                except ValueError:
                    index_dict[invoice] = -1  # Indicate that the invoice number was not found
        return index_dict

    #dct_invoice_indices = get_invoice_number_indices(list_invoice_numbr_on_page, list_invoice_numbr)
    #print(dct_invoice_indices)  # Output: {'SK026923': 8}

    new_dct_invoice_indices_all = reset_values_to_zero(dct_invoice_indices_all)
    print(f"dict all invoice indices {new_dct_invoice_indices_all}")
    for key,value in new_dct_invoice_indices_all.items():
        if value != -1:            
            #reset and input invoice number
            reset_input_invoice_numbr(key) #invoice_numbr
            print(key)
            time.sleep(1)
            """
            try:

                btn_reset=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_displayedFilters_resetBtn"]/span[2]/span[2]')
                time.sleep(0.5)
                if btn_reset.get_attribute("innerHTML") == 'Reset':
                    btn_reset.click()
            except Exception as e:
                print(e)
            time.sleep(0.5)    
            input_invoice_numbr=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_displayedFilters_ctl00_ctl05_ddl1_tagify"]/span')
            input_invoice_numbr.send_keys(key)
            input_invoice_numbr.send_keys(Keys.ENTER)
            """
            time.sleep(0.5)

            n_times = value
            idx_invoice_numbr = str(value+2)
            actions = ActionChains(driver)
            actions.send_keys(Keys.DOWN*n_times).perform()
            time.sleep(3.5)
            btn_invoice_numbr = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_MainGrid"]/tbody/tr['+idx_invoice_numbr+']/td[7]/a')
            time.sleep(2) #time delay
            btn_invoice_numbr.click()
            #print(f"invoice number '{key}' clicked")
        else:
            #print(f"Invoice number '{key}' not found on the page.")
            break

        #find all_docNum
        time.sleep(1.5)
        try:
            doc_num=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="DocumentsPanel_eskCtrlBorder_content"]/div/div[2]/div[2]/div[1]/div[1]/div/span[1]')))#.click()

            #docNum = driver.find_element(By.XPATH, '//*[@id="DocumentsPanel_eskCtrlBorder_content"]/div/div[2]/div[2]/div[1]/div[1]/div/span[1]')
        except Exception as e:
            print(e)
        try:
            docNum_class_all = driver.find_elements(By.XPATH, '//*[contains(@class,"docNum")]')
        except Exception as e:
            print(e)
        #docNum_class_all[0].get_attribute("innerHTML")


        def find_max_number(list_docNum_class_all):
            """
            Finds the maximum number in a list of strings.
            Args:
            list_docNum_class_all: A list of strings, where each string represents a number.
            Returns:
            The maximum number as an integer, or None if the list is empty.
            """
            if not list_docNum_class_all:
                return None  # Handle empty list

            max_number = int(list_docNum_class_all[0])  # Initialize with the first element

            for num_str in list_docNum_class_all[1:]:  # Start from the second element
                try:
                    num = int(num_str)
                    if num > max_number:
                        max_number = num
                except ValueError:
                    # Skip if the string cannot be converted to an integer
                    continue

            return max_number

        list_docNum_class_all = []
        #docNum_class[3].get_attribute("innerHTML")
        for docNum_class in docNum_class_all:
            list_docNum_class_all.append(docNum_class.get_attribute("innerHTML"))

        max_num= find_max_number(list_docNum_class_all)
        #print(f"Maximum number: {max_num}")  # Output: Maximum number: 2

        time.sleep(5)
        doc_num_pages = max_num #iterate through all pages, save pdf
        for p in range(1, doc_num_pages+1):
            indx = get_index(p)
            #print(f"Page {p} has index {indx}")
            try:

                doc_num_page = driver.find_element(By.XPATH, '//*[@id="DocumentsPanel_eskCtrlBorder_content"]/div/div[2]/div[2]/div['+str(indx)+']/div[1]/div/span[1]')
                doc_num_page.click()
            except Exception as e:
                print(e)
            #print(doc_num_page.get_attribute("innerHTML"))
            time.sleep(3)           

            try:
                btn_print_doc = driver.find_element(By.XPATH, "//*[contains(@class, 'printDocument')]")
                btn_print_doc.click()
                time.sleep(3)
                #print('print btn found')
            except Exception as e:
                print(e)
            
            pyautogui.moveTo(220,170, duration=1.5)
            time.sleep(1)
            pyautogui.click(button='right')
            pyautogui.keyDown('ctrl')
            pyautogui.press('s')
            pyautogui.keyUp('ctrl')
            time.sleep(1.5)

            invoice_numbr = key
            pyautogui.moveTo(350,385, duration=1.5) #620
            #pyautogui.click(button='right')
            ## get invoice_name by key
            time.sleep(0.5)
            #pyautogui.press('d')
            if '/' in invoice_numbr:
                    invoice_numbr = invoice_numbr.replace('/','_') #format invoice_numbr#
            if '/' in key:
                    key = key.replace('/','_') #format key#
            pyautogui.write(key+'_'+str(p), interval=0.25)
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(2)
            #look for right arrow multi page
            try:
                multi_page_right_arrow=driver.find_element(By.XPATH, "//*[contains(@title, 'Next')]")
                multi_page_right_arrow.click()
            except Exception as e:
                print(e)
            #doc_num_page_2=driver.find_element(By.XPATH,'//*[@id="DocumentsPanel_eskCtrlBorder_content"]/div/div[2]/div[2]/div[3]/div[1]/div/span[1]')
        
        all_tabs = driver.window_handles
        for handl in all_tabs:
                driver.switch_to.window(handl)
                if driver.title != 'Esker on Demand - Vendor Invoice':
                    driver.close()
                    time.sleep(0.5)

        try:

            pyautogui.moveTo(275,700, duration=1.5)
            pyautogui.click(button='left')
        except Exception as e:
            print(e)
        time.sleep(8)
        #WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form-footer"]/div[1]/a[8]/span'))).click() #Quit button
        #btn_invoice_numbr = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_MainGrid"]/tbody/tr[10]/td[7]/a')
        write_log(f'Download process completed at {datetime.now().strftime("%Y%m%d %H:%m")}')
        list_invoice_downloaded.append(invoice_numbr)
        

        ## create zip file, save html to pdf
        create_zip_file()
        dir = "C:/Users/" +username+ "/Downloads/"
        list_html_file = glob.glob(dir+'*.html')
        list_latest_html_file = [max(list_html_file, key=os.path.getctime)]

        for file in list_latest_html_file:
            if file.endswith(".html"):
                #extract_zip_to_pdf(invoice_numbr)
                pdfkit.from_file(file, file[:-5]+'_'+str(invoice_numbr)+'.pdf', configuration=config) #save html as pdf

        #rename 'details' pdf file in Downloads folder #zip folder
        directory_path = "C:/Users/" +username+ "/Downloads/"
        list_pdf_file = glob.glob(dir+'*.pdf')
        list_latest_pdf_file = [max(list_pdf_file, key=os.path.getctime)]
        for file in list_latest_pdf_file:
            if file.endswith(".pdf"):
                    print(file)

                    new_name = invoice_numbr# + '_0'
                    result = rename_pdf_files_by_name(directory_path, new_name)
                    print(result)

        #move pdf
        """
        src_directory = directory_path 
        dest_directory = r"C:/Users/" +username+ "/Downloads/esker_merged/"
        result = move_pdf(src_directory, dest_directory)
        print(result)
        """

        # merge pdf files, #attach xml to pdf

        path_pdf = r"C:/Users/"+username+r"/Downloads/"  # Replace with the actual path
        #list_invoice_numbr = ['SK026923'] #'24061193'

        #for invoice_numbr in list_common_invoice_numbr:
        pdf_files = get_pdf_files_with_invoice_number(path_pdf, invoice_numbr)
        print(pdf_files)

        merged_writer = merge_pdfs(pdf_files, path_pdf)

            # Save the merged PDF 
        with open(path_pdf+invoice_numbr+"_merged.pdf", "wb") as output_file:
                merged_writer.write(output_file)
        
        fldr = r"C:/Users/"+username+r"/Documents/esker_merged/"
        if os.path.isdir(fldr):
                print('folder exists')
        else:
                os.mkdir(fldr)

        print(f"Merged PDF {invoice_numbr} saved successfully.")
        write_log(f"Merged PDF {invoice_numbr} saved successfully.")

        src_directory = directory_path 
        dest_directory = r"C:/Users/" +username+ "/Downloads/esker_merged/"
        result = move_pdf(src_directory, dest_directory)
        print(result)

        #attach xml to 
        #try:

                #merge_xml_to_pdf(path_pdf, invoice_numbr)
        #except Exception as e:
                #print(e)
        keywords = ['_merged']
            #dir = r"C:/Users/"+username+r"/Downloads/"
        for file in os.listdir(path_invoice):
                filename = os.fsdecode(file)
                ext = Path(file).suffix
                for keyword in keywords:
                    if keyword in filename: # this tests for substrings
                        file_path = os.path.join(path_invoice, filename)
                        print(file_path)

                        if os.path.isfile(os.path.join(fldr, filename)):
                            os.remove(os.path.join(fldr, filename))
                            shutil.move(file_path, fldr)

        # Delete the original PDF files
        for pdf in pdf_files:
                try:
                    os.remove(path_pdf + pdf)
                except Exception as e:
                    print(e)
                #print(f"Deleted PDF file '{pdf}'.")
        #remove 'details.html' and '.zip' files
        remove_html_files_by_name(path_pdf)
        remove_zip_files(path_pdf)

        # write to excel
        fldr_invoice_downloaded = r"C:/Users/"+ username + r"/Documents/esker_merged/"
        df_invoice_downloaded = pd.DataFrame(list_invoice_downloaded, columns=['invoice_downloaded'])
        with pd.ExcelWriter(fldr_invoice_downloaded+'invoice_downloaded.xlsx', mode="a", if_sheet_exists='overlay', engine='openpyxl') as writer_downloaded:
                    df_invoice_downloaded.to_excel(writer_downloaded, sheet_name='invoice_downloaded', engine='openpyxl')
        
        """
        path_page = r"C:/Users/"+ username + r"/Documents/esker_merged/"
        df_page = pd.read_excel(path_page+ 'page.xlsx', sheet_name='page', engine='openpyxl')
        pg = df_page.at[0, 'page']
        pg +=1
        if pg > df_page.at[0, 'page_total']:
                print('current_page exceeds total pages')
    
        else:
                df_page['page'] = pg
                with pd.ExcelWriter(path_page+r'page.xlsx', mode="a", if_sheet_exists='replace', engine='openpyxl') as writer_current_page:
                    df_page.to_excel(writer_current_page, sheet_name='page', index=False)
        """
        write_log(f'Process completed at {datetime.now().strftime("%Y%m%d %H:%m")}')
        #done

if __name__ == "__main__":
    main()

