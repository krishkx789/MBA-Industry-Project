import json
import time, os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities


def init_webdriver():
    # edge_driver_path = os.getcwd()+'\edgedriver_win64\msedgedriver.exe'
    driver_path = os.getcwd()+'\edgedriver_win64\geckodriver.exe'
    options = Options()
    # options.binary_location = 'C:\\Users\\Karna\\AppData\\Local\\Microsoft\\WindowsApps\\firefox.exe'
    # options.add_argument('--headless')  # Enable headless mode
    options.add_argument('--disable-gpu')  # Disable GPU acceleration
    # options.add_argument('--window-size=1920x1080')

    return driver_path, options


def load_config():
    with open('./config.json', 'r') as config_file:
        config = json.load(config_file)
        return config


def __is_html_ele_exists__(driver, element):
    try:
        return True if driver.find_element(By.XPATH, value=(element)).text else False
    except NoSuchElementException:
        return False


def run_scrapper():
    config = load_config()
    driver_path, options = init_webdriver()
    service = Service(executable_path=driver_path)
    driver = webdriver.Firefox(options=options,
                            #    firefox_profile=profile,
                               service=service)
    try:
        url = config["url"]
        out_lst = []
        driver.get(url)
        for company in config['ind_500']:
            out_d = {}
            print(f"Running for company -> {company}")
            time.sleep(3)
            if __is_html_ele_exists__(driver, config['elements']['bottom_ok_btn_ele']):
                ok_btn = driver.find_element(By.XPATH, config['elements']['bottom_ok_btn_ele'])
                ok_btn.click()
            search_box = driver.find_element(By.XPATH, config['elements']['search_box_ele'])
            search_box.clear()
            for char in company:
                time.sleep(1)
                search_box.send_keys(char)
            # search_box.send_keys(Keys.RETURN)
            time.sleep(3)
            if __is_html_ele_exists__(driver, config['elements']['searchResults']):
                # print("In SearchResults")
                driver.find_element(By.XPATH, config['elements']['searchResults']).click()
                time.sleep(1)
            else:
                print(f'skipped running for company -> {company}')
                continue
            if __is_html_ele_exists__(driver, config['elements']['new_page_title_ele']):
                # print("In new_page_title_ele")
                title = driver.find_element(By.XPATH, config['elements']['new_page_title_ele']).text
            else:
                print(f'skipped running for company -> {company}')
                continue
            # time.sleep(1)
            if company.lower() in title.lower():
                out_d['title'] = title.lower()
                if __is_html_ele_exists__(driver, config['elements']['esg_risk_score_ele']):
                    # print("In esg_risk_score_ele")
                    esg_risk_score = driver.find_element(By.XPATH, config['elements']['esg_risk_score_ele']).text
                    out_d['esg_risk_score'] = esg_risk_score
                # time.sleep(1)
                if __is_html_ele_exists__(driver, config['elements']['esg_category_div_ele']):
                    # print("In esg_category_div_ele")
                    esg_risk_category = driver.find_element(By.XPATH, config['elements']['esg_category_div_ele']).text
                    out_d['esg_risk_category'] = esg_risk_category
                # time.sleep(1)
                if __is_html_ele_exists__(driver, config['elements']['ind_group_ele']):
                    # print("In ind_group_ele")
                    industry = driver.find_element(By.XPATH, config['elements']['ind_group_ele']).text
                    out_d['industry'] = industry
                # time.sleep(1)
                out_lst.append(out_d)
            else:
                print(f'Failed fetching relevant company info for {company}')
            # time.sleep(1)
        df_out = pd.DataFrame(out_lst)
        df_out.to_csv(r'esg_risk_ratings.csv', index=False)
    except Exception as e:
        print(f"error -> {e}")
        df_out = pd.DataFrame(out_lst)
        df_out.to_csv(r'esg_risk_ratings.csv', index=False)
    finally:
        driver.quit()
        df_out = pd.DataFrame(out_lst)
        df_out.to_csv(r'esg_risk_ratings.csv', index=False)

if __name__ == "__main__":
    run_scrapper()