import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service  
from selenium.webdriver.support.ui import Select
import pandas as pd
import requests
import os

print("*** AUTOMATION STARTING ***")

#current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

#file paths
CSV_PATH = os.path.join(current_dir, "ParaBank users.csv")
DRIVER_PATH = os.path.join(current_dir, "chromedriver")
REPORT_PATH = os.path.join(current_dir, "Parabank_Report.xlsx")


#Read the customers from CSV
try:
    df_customers = pd.read_csv(CSV_PATH)
    print(f"successfully loaded customers")
except Exception as e:
    print(f"Error reading CSV: {e}")
    exit()

#function to get USDtoEUR rate(might api not work)
def UsdToEur_rate():
    try:
        response = requests.get("https://api.exchangerate-api.com/v4/latest/USD", timeout=10)
        data = response.json()
        rate = data['rates']['EUR']
        print(f"1 USD = {rate} EUR")
        return rate
    except requests.exceptions.RequestException as e:
        print(f"Error fetching exchange rate: {e}. Using custom rate of 0.93.")
        return 0.93

exchange_rate = UsdToEur_rate()





#initialize the browser
try:
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    
    
    service = Service(executable_path=DRIVER_PATH)  #service class for newer Selenium versions
    driver = webdriver.Chrome(service=service, options=options)
    
    driver.maximize_window()
    wait = WebDriverWait(driver, 10)
except Exception as e:
    print(f"error: {e}")
    exit()

#results for each customer
results = []


for index, customer in df_customers.iterrows():
    customer_name = f"{customer['First Name']} {customer['Last Name']}"
    
    customer_result = customer.to_dict()
    customer_result['Status'] = "Not Started"
    customer_result['Error'] = ""

    try:
        #open Parabank
        driver.get("https://parabank.parasoft.com/parabank/index.htm")
        
        #click Register
        register_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Register")))
        register_link.click()
        time.sleep(1)
        
        #fill Registration Form        

        def fill_field(field_id, value, field_name):
            try:
                if pd.notna(value):
                    element = driver.find_element(By.ID, field_id)
                    element.clear()
                    element.send_keys(str(value))
                return True
            except Exception as e:
                print(f"Error  {field_name}: {e}")
                return False

        fill_field("customer.firstName", customer['First Name'], "First Name")
        fill_field("customer.lastName", customer['Last Name'], "Last Name")
        fill_field("customer.address.street", customer['Address'], "Address")
        fill_field("customer.address.city", customer['City'], "City")
        fill_field("customer.address.state", customer['State'], "State")
        fill_field("customer.address.zipCode", customer['Zip Code'], "Zip Code")
        fill_field("customer.phoneNumber", customer['Phone Number'], "Phone")
        fill_field("customer.ssn", customer['SSN'], "SSN")
        
        #credentials
        username = customer['Username']
        password = customer['Password']
        fill_field("customer.username", username, "Username")
        fill_field("customer.password", password, "Password")
        fill_field("repeatedPassword", password, "Password Confirm")

        #submit Registration
        driver.find_element(By.XPATH, "//input[@value='Register']").click()
        time.sleep(2)

        #Check if registration successful
        try:
            #success message
            success_element = driver.find_element(By.XPATH, "//h1[contains(text(), 'Welcome')] | //p[contains(text(), 'success')]")
            print(f"registration successful for {username}")
            customer_result['Generated_Username'] = username
            customer_result['Generated_Password'] = password
            customer_result['Status'] = "Registration Successful"
            
            #request Loan only if registration successful
            try:
                wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Request Loan"))).click()
                time.sleep(1)
                print(f"Navigated to Request Loan for {username}")
                
                #calculate 20%
                initial_deposit = float(customer['Initial Deposit']) if pd.notna(customer['Initial Deposit']) else 0
                down_payment = round(initial_deposit * 0.2, 2)
                loan_amount = 10000
                
                # loan form
                amount_field = wait.until(EC.presence_of_element_located((By.ID, "amount")))
                amount_field.send_keys(str(loan_amount))
                
                down_payment_field = wait.until(EC.presence_of_element_located((By.ID, "downPayment")))
                down_payment_field.send_keys(str(down_payment))
                
                #first available account
                from_account_select = Select(wait.until(EC.presence_of_element_located((By.ID, "fromAccountId"))))
                from_account_select.select_by_index(0)
                

                submit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.button[value='Apply Now']")))
                submit_button.click()
                time.sleep(2)

                
                #loan status
                try:
                    loan_status_element = wait.until(EC.presence_of_element_located((By.ID, "loanStatus")))
                    loan_status = loan_status_element.text
                    customer_result['Loan_Status'] = "Successful"
                    customer_result['Down_Payment'] = down_payment
                    customer_result['Loan_Error'] = ""
                    print(f"loan request for {username}: {loan_status}")
                except Exception as status_e:
                    customer_result['Loan_Status'] = "Failed"
                    customer_result['Down_Payment'] = down_payment
                    customer_result['Loan_Error'] = f"Failed to retrieve loan status: {str(status_e)}"
                    print(f"loan request for {username}: {str(status_e)}")
            except Exception as loan_e:
                customer_result['Loan_Status'] = "Failed"
                customer_result['Down_Payment'] = down_payment if 'down_payment' in locals() else 0
                customer_result['Loan_Error'] = str(loan_e)
                print(f"loan request for {username}: {loan_e}")
            
        except NoSuchElementException:
            #error message
            try:
                error_element = driver.find_element(By.CLASS_NAME, "error")
                error_msg = error_element.text
                print(f"registration failed: {error_msg}")
                customer_result['Error'] = error_msg
                customer_result['Status'] = "Registration Failed"
                results.append(customer_result)
                continue
            except NoSuchElementException:
                customer_result['Error'] = "Unknown registration status"
                customer_result['Status'] = "Registration Unknown"
                results.append(customer_result)
                continue

        #Log Out
        try:
            logout_link = driver.find_element(By.LINK_TEXT, "Log Out")
            logout_link.click()
            time.sleep(1)
        except:
            print("fail")

        # store successful registration
        results.append(customer_result)
        

    except Exception as e:
        print(f"error: {e}")
        customer_result['Error'] = str(e)
        customer_result['Status'] = "Error"
        results.append(customer_result)
        continue



#create final report
try:
    df_report = pd.DataFrame(results)
    
    #select only the most important columns for the report
    important_columns = ['First Name', 'Last Name', 'Username', 'Status', 'Error', 
                        'Generated_Username', 'Generated_Password', 'DOB', 'Debit Card', 'CVV',
                        'Loan_Status', 'Down_Payment', 'Loan_Error']
    
    #filter to only include columns that actually exist in our data
    available_columns = [col for col in important_columns if col in df_report.columns]
    final_report = df_report[available_columns]
    
    final_report.to_excel(REPORT_PATH, index=False)
    print(f"report saved in: {REPORT_PATH}")

        
except Exception as e:
    print(f"{e}")

# close browser
driver.quit()