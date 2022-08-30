from selenium import webdriver
from selenium.webdriver.chrome import options
from webdriver_manager.chrome import ChromeDriverManager
import time, csv, requests, shutil, os
from requests import get
from io import BytesIO
from PIL import Image
import pandas as pd
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("window-size=1200x600")
options.add_argument("--disable-blink-features=AutomationControlled")
options.headless = False
options.add_argument("--disable-infobars")
options.add_argument("--disable-extensions")
options.add_argument("test-type")
options.add_argument('--disable-useAutomationExtension')
options.add_argument("--disable-xss-auditor")
options.add_argument("--disable-web-security")
options.add_argument("--allow-running-insecure-content")
options.add_argument("--disable-setuid-sandbox")
options.add_argument("--disable-webgl")
options.add_argument("--disable-popup-blocking")
options.add_argument("no-default-browser-check")
options.add_argument("--disable-notifications")

def check_for_internet():
    print('Checking for Internet connection...')
    while True:
        try:
            request = requests.get('https://www.google.com/', timeout=30)
            time.sleep(5)
            print("Connected to the Internet")
            break
        except (requests.ConnectionError, requests.Timeout) as exception:
            pass
            try:
                driver.refresh()
            except:
                break


while True:
    try:
        user_input = int(input("Enter the number of images per reference : "))
        if user_input < 0:
            raise Exception("Sorry, negative numbers not allowed !")

        break
    except:
        print("Please enter a valid number")

driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
driver.implicitly_wait(1)
df = pd.read_excel('WatchList.xlsx')

try:
    os.mkdir('images')
    os.chdir(f"{format(os.getcwd())}//images")
except:
    os.chdir(f"{format(os.getcwd())}//images")

check_for_internet()
links = {}
numb = 1
for i in df.iloc():
    while True:
        try:
            if numb % 50 == 0:
                driver.quit()
                time.sleep(5)
                driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
                
            driver.get('https://www.google.com.pk/imghp?hl=en&authuser=0&ogbl')
            numb+=1
            print(numb)
            # cookies window click
            time.sleep(3)
            try:
                driver.execute_script("document.getElementById('L2AGLb').click()")
            except:
                pass
            driver.find_element(by=By.XPATH, value='//input[@title="Search"]').send_keys(
                f'{i["Brand"]} {i["Model"]} {i["Reference Number"]}' + Keys.ENTER)
            all_results = driver.find_elements(By.XPATH, '//div[@id="islrg"]/div/div')

            lst, count = [], 1
            for j in (all_results):
                try:
                    if len(lst) == user_input:
                        break
                    ele = j
                    desired_y = (ele.size['height'] / 2) + ele.location['y']
                    window_h = driver.execute_script('return window.innerHeight')
                    window_y = driver.execute_script('return window.pageYOffset')
                    current_y = (window_h / 2) + window_y
                    scroll_y_by = desired_y - current_y
                    driver.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by)
                    j.click()
                    time.sleep(5)

                    url = driver.find_element(by=By.XPATH,
                                              value='//div[@id="Sva75c"]/div/div/div[3]/div[2]/c-wiz/div/div/div/div[3]/div/a/img').get_attribute(
                        'src')
                    print('\n', '-' * 120, '\n' * 2, f"Link: {url}")
                    if len(url.split('data:')) > 1:
                        continue

                    image = Image.open(BytesIO(get(url, timeout=10).content))
                    width, height = image.size

                    print(f"Resolution: {image.size[0]} x {image.size[1]}")
                    if width < 500 or height < 500:
                        print('Low resolution image, rejecting...')
                        continue

                    file_name = f'{i["Brand"]}_{i["Model"]}_{i["Reference Number"]}_{count}.jpg'

                    res = get(url, stream=True, timeout=10)
                    if res.status_code == 200:
                        with open(file_name, 'wb') as f:
                            shutil.copyfileobj(res.raw, f)
                        print('Image sucessfully Downloaded:', file_name)
                        lst.append(url)
                        count += 1
                    else:
                        print('Image Couldn\'t be retrieved')

                except Exception as e:
                    check_for_internet()
                    print(e)
                    continue

            links[f'{i["Brand"]} {i["Model"]} {i["Reference Number"]}'] = lst
            break

        except:
            check_for_internet()
            print('Retrying the current reference....')

try:
    with open('google_result.csv', 'w', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file)
        for i in links:
            csv_writer.writerow([i, ' '.join(links[i])])
    driver.close()
except:
    pass