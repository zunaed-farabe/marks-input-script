from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.service import Service
from selenium.webdriver.chrome.options import Options
import webbrowser
import time
import pandas as pd
import traceback
import os 
def autoInputMarks(filename, username, password, semester, section):

 # Get the path to the directory containing the script
    current_dir = os.path.dirname(os.path.abspath(__file__))

    print(current_dir)
    # # Construct the path to the bundled WebDriver executable
    webdriver_path = os.path.join(current_dir, 'webdriver', 'msedgedriver.exe')
    print(webdriver_path)
    service = Service(webdriver_path)
    
    # # os.environ['EDGE_DRIVER_PATH'] = webdriver_path
    driver = webdriver.Edge(service=service)
 
    driver.maximize_window()
    
    # Navigate to the URL of the marks entry form
    url = "http://182.160.97.196:8088/NUBERP/courseResult/marksEntry"
    driver.get(url)
    
    driver.implicitly_wait(10)
    # Find the username and password input fields using XPath
    username_input = driver.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div/div/form/div[1]/input')
    password_input = driver.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div/div/form/div[2]/input')
    
    # Enter your credentials
    # username = "NUB9221995"
    # password = "nub3523"
    
    username_input.send_keys(username)
    password_input.send_keys(password)
    
    # Find and click the login button using XPath
    login_button = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div[2]/div/div/div/form/button")
    login_button.click()
    
    
    # Wait for the login process to complete (adjust the delay as needed)
    driver.implicitly_wait(5)
    
    
    url = 'http://182.160.97.196:8088/NUBERP/courseResult/marksEntry'
    
    driver.get(url)
    driver.implicitly_wait(5)
    
    
    # Find and select the semester dropdown
    semester_dropdown = Select(driver.find_element(By.ID, "semesterId"))
    # semester_dropdown.select_by_visible_text("Spring 2024")
    semester_dropdown.select_by_visible_text(semester)
    
    # Find and select the course and section dropdown
    section_dropdown = Select(driver.find_element(By.ID, "sectionId"))
    # section_dropdown.select_by_visible_text("CSE 1360 - A")
    section_dropdown.select_by_visible_text(section)
    
    # Find and click the "Populate Students" button
    populate_button = driver.find_element(By.ID, "loadStudents")
    populate_button.click()
    
    driver.implicitly_wait(5)
    
    # Fill in the marks for each student
    # marks = pd.read_excel("testdata.xlsx")
    marks = pd.read_excel(filename)
    
    bodypath = "/html/body/div[3]/div[2]/div/div[2]/div/div/form/div[2]/div[1]/div[1]/table/tbody"
               
    tablebody = element_outer_html = driver.find_element(By.XPATH, bodypath)
    rows = tablebody.find_elements(By.TAG_NAME, 'tr')
    
    for row in marks.iterrows():
    
        student_id = row[1]["Student ID"]
        class_attendance_marks = row[1]["Class Attendance"]
        continuous_assessment_marks = row[1]["Continuous Assessment"]
        mid_term_marks = row[1]["Mid Term"]
        final_marks = row[1]["Final"]
        
        for i in range(1,len(rows)+1):
            
            path = f"/html/body/div[3]/div[2]/div/div[2]/div/div/form/div[2]/div[1]/div[1]/table/tbody/tr[{i}]/td[2]"
            row_id = tablebody.find_element(By.XPATH, path)
            # print(id[-4:], rowdata.text[-4:]) 
            if student_id[-4:] == row_id.text[-4:]:
    
                # row_path = "/html/body/div[3]/div[2]/div/div[2]/div/div/form/div[2]/div[1]/div[1]/table/tbody"
                # row_element = tablebody.find_element(By.XPATH, path)
                class_attendance_input = tablebody.find_element(By.CSS_SELECTOR, f"#classAttendanceMarks{i}")
                continuous_assessment_input = tablebody.find_element(By.CSS_SELECTOR, f"#continuousAssessmentMarks{i}")
                mid_term_input = tablebody.find_element(By.CSS_SELECTOR, f"#midTermMarks{i}")
                final_input = tablebody.find_element(By.CSS_SELECTOR,  f"#finalMarks{i}")
                
                # Fill in the marks for the current student
                class_attendance_input.clear()
                class_attendance_input.send_keys(class_attendance_marks)
                continuous_assessment_input.clear()
                continuous_assessment_input.send_keys(continuous_assessment_marks)
                mid_term_input.clear()
                mid_term_input.send_keys(mid_term_marks)
                final_input.clear()
                final_input.send_keys(final_marks)
                driver.implicitly_wait(2)
                
                break
            else:
                print(student_id, "Not found")
    

def get_user_inputs():
    
    filename = input("Enter Filename: ")
    time.sleep(2)
    username = input("Enter ERP username: ")
    time.sleep(2)
    password = input("Enter ERP password: ")
    time.sleep(2)
    semester = input("Enter Semester: ")
    time.sleep(2)
    section = input("Enter section: ")
    time.sleep(2)
    return filename, username, password, semester, section


if __name__ == "__main__":
    
    

    filename, username, password, semester, section = get_user_inputs()
    time.sleep(5)
    print("Input taked Successfully, Loading webpage")
    try:
        autoInputMarks(filename, username, password, semester, section)
        
        pass
    except Exception as e:
        traceback.print_exc()
        input("Please Enter to exit.....")
        input("Please Enter to exit.....")
    