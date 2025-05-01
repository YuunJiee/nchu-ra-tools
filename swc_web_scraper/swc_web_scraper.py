from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import pandas as pd

# **初始化 Selenium**
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

# 🔹 登入資訊 (請修改成你的帳號密碼)
USERNAME = ""
PASSWORD = ""

# ✅ **讀取工程編號**
def read_project_ids(file_path):
    """讀取 .txt 檔案內的工程編號"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            project_ids = [line.strip() for line in f if line.strip()]
        print(f"✅ 讀取 {len(project_ids)} 筆工程編號")
        return project_ids
    except Exception as e:
        print(f"❌ 無法讀取 {file_path}，錯誤: {e}")
        return []
    

# ✅ **存入 Excel**
def save_to_excel(data_list, output_file="output.xlsx"):
    """將所有工程資料存入 Excel"""
    df = pd.DataFrame(data_list)
    df.to_excel(output_file, index=False)
    print(f"✅ 已將資料存入 `{output_file}`")

# ✅ **切換 frame 並點擊連結**
def switch_and_click(link_text):
    """切換到指定 frame 並點擊連結"""
    driver.switch_to.default_content()
    driver.switch_to.frame("top_frame")
    #print(f"✅ 已切換到 top_frame")

    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, link_text)))
    link.click()
    #print(f"✅ 成功點擊「{link_text}」")

    driver.switch_to.default_content()
    driver.switch_to.frame("body_frame")
    #print("✅ 已切回 body_frame")

def login():
    """登入系統"""
    
    # 填寫帳號與密碼
    username_input = wait.until(EC.presence_of_element_located((By.ID, "txbUserId")))
    password_input = driver.find_element(By.ID, "txbPassword")
    captcha_input = driver.find_element(By.ID, "txbValidateCode")
    login_button = driver.find_element(By.ID, "btnLogin")

    username_input.clear()
    password_input.clear()
    captcha_input.clear()

    username_input.send_keys(USERNAME)
    password_input.send_keys(PASSWORD)

    # **手動輸入驗證碼**
    captcha_text = input("請查看驗證碼圖片，輸入驗證碼後按 Enter：")
    captcha_input.send_keys(captcha_text)

    # 點擊登入
    login_button.click()

    print("⏳ 登入中，請稍後！")

    # 等待登入完成
    time.sleep(5)

    # **關閉 `display_message.asp` 分頁**
    if len(driver.window_handles) > 1:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if "display_message.asp" in driver.current_url:
                #print("✅ 發現 `display_message.asp`，即將關閉...")
                driver.close()  # 關閉公告頁面
                break

    # **切換回主頁 (`Home2015.aspx`)**
    driver.switch_to.window(driver.window_handles[0])
    print("✅ 登入成功！回到主頁:", driver.current_url)

    driver.get("https://mis.ardswc.gov.tw/Swcbtaojr/MainSearch.aspx")
    time.sleep(3)  # 確保頁面加載
    print("✅ 已轉跳 `MainSearch.aspx`")

    return True

def get_project_data(project_id):

    project_data = {
        "工程編號": project_id,
        "坐標位置(TWD97x)": "",
        "坐標位置(TWD97y)": "",
        "核定經費": "",
        "工程結算金額": "",
        "工程位置": "",
        "鄉鎮區": "",
        "村里": "",
        "竣工日期": "",
    }

    # 🔍 **輸入工程編號搜尋**
    search_box = wait.until(EC.presence_of_element_located((By.ID, "txtKeyWord")))
    search_box.clear()
    search_box.send_keys(project_id)

    # **點擊搜尋按鈕**
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@onclick, 'goSearch')]")))
    search_button.click()

    # ✅ **等待「搜尋結果表格」出現**
    try:
        wait.until(EC.presence_of_element_located((By.ID, "tabSEDT01")))  # 確保表格出現
        wait.until(EC.presence_of_element_located((By.ID, "listview1")))  # 確保 `tbody` 加載
        #print("✅ 搜尋結果表格已載入")
    except:
        #print("❌ 搜尋結果未載入，請檢查工程編號是否正確")
        return project_data

    # ✅ **找到搜尋結果，點擊工程名稱**
    result_link = wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"//tbody[@id='listview1']//a[contains(@onclick, \"goEnginDetail('{project_id}')\")]")
    ))

    # **滾動頁面，確保 Selenium 找得到**
    driver.execute_script("arguments[0].scrollIntoView();", result_link)
    time.sleep(1)  # 等待滾動完成

    # **等待 `a` 可點擊**
    result_link.click()
    print("✅ 成功點擊搜尋結果！")

    # ✅ **等待新分頁開啟，切換到 `bframeset.aspx`**
    time.sleep(3)  # 確保新分頁打開
    new_tab = driver.window_handles[-1]  # 獲取最新開啟的分頁
    driver.switch_to.window(new_tab)  # 切換到新分頁
    #print("✅ 已切換到新分頁:", driver.current_url)

    # **切換到 `body_frame`**
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "body_frame")))
    #print("✅ 已切換到 `body_frame`")
    
    # 🔍 取得 `竣工日期`**
    completion_date_element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'竣工日期')]/following-sibling::td"))
        )
    project_data["竣工日期"] = completion_date_element.text.strip()


    switch_and_click("基本資料")
    
    # **取得坐標**
    coord_element = wait.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'坐標')]/following-sibling::td")))
    coord_text = coord_element.text
    project_data["坐標位置(TWD97x)"] = coord_text.split("TWD97 X：")[1].split("　Y：")[0]
    project_data["坐標位置(TWD97y)"] = coord_text.split("Y：")[1].split("　　WGS84")[0]
    
    # **取得 `工程位置` (縣市)、`鄉鎮區`、`村里`**
    project_data["工程位置"] = wait.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'縣市')]/following-sibling::td"))).text.strip()
    project_data["鄉鎮區"] = wait.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'鄉鎮')]/following-sibling::td"))).text.strip()
    project_data["村里"] = wait.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'村里')]/following-sibling::td"))).text.strip()


    # **切換到 `top_frame` 並點擊「預算額度」**
    switch_and_click("預算額度")

    # **取得核定經費與結算金額**
    project_data["核定經費"] = wait.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'工程核定金額')]/following-sibling::td"))).text.strip().replace(",", "").replace("元", "")
    project_data["工程結算金額"] = wait.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'工程結算淨額')]/following-sibling::td"))).text.strip().replace(",", "").replace("元", "")

    print(project_data)

    return project_data


# ✅ **主程式**
def main():
    # 讀取工程編號
    project_ids = read_project_ids("input.txt")
    if not project_ids:
        print("❌ 沒有可搜尋的工程編號")
        return

    # **登入**
    login_url = "https://mis.ardswc.gov.tw/newmis/xtaojr.aspx"
    driver.get(login_url)

    if not login():
        print("❌ 登入失敗，結束程式")
        driver.quit()
        return

    # **結果列表**
    results = []

    # **打開 `MainSearch.aspx`**
    driver.get("https://mis.ardswc.gov.tw/Swcbtaojr/MainSearch.aspx")
    print("✅ 進入 `MainSearch.aspx`")

    for project_id in project_ids:
        print(f"\n🔍 正在搜尋: {project_id}")
        try:
            # **執行搜尋**
            data = get_project_data(project_id)
            if data:
                results.append(data)
                print(f"✅ 成功獲取 {project_id} 資料")
            else:
                print(f"❌ {project_id} 找不到，跳過...")
        except Exception as e:
            print(f"❌ {project_id} 發生錯誤，跳過... 錯誤: {e}")

        # **關閉 `bframeset.aspx`，回到 `MainSearch.aspx`**
        if len(driver.window_handles) > 1:
            driver.close()  # 關閉 `bframeset.aspx`
            driver.switch_to.window(driver.window_handles[0])  # 回到 `MainSearch`
            #print("✅ 關閉 `bframeset.aspx`，回到 `MainSearch.aspx`")

        # **等待 2 秒，避免觸發網站風控**
        time.sleep(2)

    # **存入 Excel**
    if results:
        save_to_excel(results)
    else:
        print("❌ 沒有成功獲取任何資料")

    # **關閉瀏覽器**
    driver.quit()
    print("✅ 爬取完成！")

# ✅ **執行程式**
if __name__ == "__main__":
    main()