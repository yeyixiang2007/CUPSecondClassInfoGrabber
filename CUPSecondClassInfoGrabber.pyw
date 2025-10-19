import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter import ttk
import threading
from bs4 import BeautifulSoup
import re
import time
import webbrowser
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    NoSuchElementException,
)

# ====================================================================
# 全局变量和配置
# ====================================================================
URL_LOGIN = "https://sct.cup.edu.cn/mucenter/index/index"
BASE_URL = "https://sct.cup.edu.cn"
WAIT_TIMEOUT = 6000  # Selenium等待登录成功的最大时间（秒）
LOGIN_SUCCESS_SELECTOR = (
    "div.my_name > div.name"  # 登录成功后页面上表示用户名的CSS选择器
)
COOKIE_DOMAIN = "sct.cup.edu.cn"
DEFAULT_FETCH_COUNT = 20  # 默认获取详细数据的活动数量

# 全局变量
USER_NAME = "同学"  # 存储登录后的用户名
GLOBAL_COOKIES = {}  # 存储登录后获取的Cookie


def extract_data_from_text(detail_text, activity_type):
    """
    根据活动的 InnerText 内容和活动类型，使用正则表达式提取详细数据。

    Args:
        detail_text (str): 活动详情页的全部文本内容。
        activity_type (str): 活动类型 ("Activity", "Organization", "Other")。

    Returns:
        dict: 包含 time_duration, tags, status_combined 的字典。
    """
    data = {
        "time_duration": "N/A",
        "tags": "N/A",
        "status_combined": "N/A",
    }

    # 提取活动时间/持续时长
    # 匹配 "活动时间：..." 直到下一个换行，并捕获括号内的时长
    match_time = re.search(
        r"活动时间：(.*?)\s*（(\d+\.?\d*小时)）", detail_text, re.DOTALL
    )
    if match_time:
        data["time_duration"] = f"{match_time.group(1).strip()} ({match_time.group(2)})"

    if activity_type == "Activity":
        # 提取活动标签 (适用于第二课堂活动)
        # 匹配 "（标签类别） 标签描述 \n活动时间："
        match_tags = re.search(r"\n（(.*?)）(.*?)\n活动时间：", detail_text, re.DOTALL)
        if match_tags:
            category = match_tags.group(1).strip()
            desc_raw = match_tags.group(2).strip()
            # 移除描述中的积分信息
            desc = re.sub(r"积分[+-]\d+\.?\d*", "", desc_raw).strip()
            data["tags"] = category + " / " + desc.replace(" ", " / ")

        # 提取签到签退情况
        check_in_status = "N/A"
        check_out_status = "N/A"

        # 匹配 "签到情况：" 后的状态，忽略括号内的时间戳
        match_in = re.search(
            r"签到情况：(.*?)(?:\s*[\(（].*?[\)）])?\n", detail_text, re.DOTALL
        )
        if match_in:
            check_in_status = match_in.group(1).strip()

        # 匹配 "签退情况：" 后的状态，忽略括号内的时间戳
        match_out = re.search(
            r"签退情况：(.*?)(?:\s*[\(（].*?[\)）])?\n", detail_text, re.DOTALL
        )
        if match_out:
            check_out_status = match_out.group(1).strip()

        data["status_combined"] = f"{check_in_status} | {check_out_status}"

    return data


def fetch_detail_data(driver, link, activity_type):
    """
    使用 Selenium 访问活动详情页，提取页面 Body Text 以获取详细数据。

    Args:
        driver (webdriver.Chrome): Selenium WebDriver实例。
        link (str): 活动详情页的URL。
        activity_type (str): 活动类型。

    Returns:
        dict: 包含详细信息的字典。
    """
    data = {
        "time_duration": "N/A (请求失败)",
        "tags": "N/A",
        "status_combined": "N/A (请求失败)",
    }

    try:
        driver.get(link)
        # 简单等待页面内容加载
        time.sleep(1)
        # 获取整个 body 的 InnerText
        body_text = driver.find_element(By.TAG_NAME, "body").text
        extracted_data = extract_data_from_text(body_text, activity_type)
        return extracted_data
    except NoSuchElementException:
        # 如果找不到 body 元素
        data["time_duration"] = "访问详情页出错: 元素未找到"
    except Exception as e:
        # 其他访问错误
        data["time_duration"] = f"访问详情页出错: {e.__class__.__name__}"
    return data


def extract_activity_data(html_content):
    """
    解析用户中心首页 HTML 内容，提取已报名活动的基本信息。

    Args:
        html_content (str): 用户中心首页的 HTML 源码。

    Returns:
        list: 包含每个活动信息的字典列表。
    """
    soup = BeautifulSoup(html_content, "html.parser")
    activity_list = []

    global USER_NAME
    # 提取用户名
    name_div = soup.select_one("div.my_name > div.name")
    if name_div and name_div.text:
        USER_NAME = name_div.text.strip()

    # 用户的报名信息通常在第二个 ul.events_list 中
    events_ul_list = soup.select("div.common_block > div.events_box > ul.events_list")
    if len(events_ul_list) > 1:
        events_ul = events_ul_list[1]
        for li in events_ul.find_all("li", recursive=False):
            data = {
                "name": "N/A",
                "kind": "N/A",
                "id": "N/A",
                "link": "",
                "type": "Other",
                "time_duration": "N/A (等待获取...)",
                "tags": "N/A",
                "status_combined": "N/A (等待获取...)",
            }

            a_tag = li.find("a", href=True)
            if a_tag:
                link = a_tag["href"]
                # 完善相对链接
                if not link.startswith("http"):
                    link = BASE_URL + link
                data["link"] = link

                # 排除志愿者活动 (通常是链接中包含 volunteer)
                if "volunteer" in link:
                    continue

                # 确定活动类型
                if "activitynew" in link:
                    data["type"] = "Activity"
                    data["kind"] = "第二课堂活动"
                elif "association" in link:
                    data["type"] = "Organization"
                    data["kind"] = "学生组织"

                # 提取活动ID (根据链接中的 actid, aid, 或 id 参数)
                match_actid = re.search(r"actid=(\d+)", data["link"])
                match_aid = re.search(r"aid=(\d+)", data["link"])
                if match_actid:
                    data["id"] = match_actid.group(1)
                elif match_aid:
                    data["id"] = match_aid.group(1)
                else:
                    match_id = re.search(r"id=(\d+)", data["link"])
                    if match_id:
                        data["id"] = match_id.group(1)

                # 提取活动名称
                name_div = li.find("div", class_="course_name")
                if name_div and name_div.text:
                    data["name"] = name_div.text.strip().replace("\n", " ")

                activity_list.append(data)

    return activity_list


def get_cookie_and_activities(fetch_count):
    """
    主控制流程：通过 Selenium 登录、获取首页数据、并按需获取详情页数据。

    Args:
        fetch_count (int): 需要获取详情的活动数量。
    """
    global GLOBAL_COOKIES
    driver = None
    parsed_activities = []

    try:
        # 配置 Chrome 选项
        options = Options()
        # 排除不必要的日志输出
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        # 初始化 WebDriver
        driver = webdriver.Chrome(options=options)
        driver.get(URL_LOGIN)

        log_message(
            f"请在弹出的浏览器窗口中完成登录。您有最多 {WAIT_TIMEOUT} 秒时间..."
        )

        # 等待用户登录成功
        WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, LOGIN_SUCCESS_SELECTOR))
        )
        log_message("已检测到登录成功！正在获取 Cookie 和页面内容...")

        # 1. 获取 Cookie
        GLOBAL_COOKIES = {c["name"]: c["value"] for c in driver.get_cookies()}
        # 2. 获取页面源码
        page_source = driver.page_source

        # 3. 隐藏浏览器窗口
        driver.minimize_window()

        log_message("页面内容获取成功，浏览器已隐藏。正在解析首页数据...")

        # 解析首页，获取活动基本信息
        parsed_activities = extract_activity_data(page_source)

        # 对活动进行排序（确保获取的是最新的 N 条）
        sorted_activities = sort_activities(parsed_activities)

        # 立即在界面上显示基本信息 (状态为 "等待获取...")
        display_results(sorted_activities)

        # 筛选出需要获取详情的活动（即状态仍为“等待获取...”的活动）
        details_to_fetch = [
            a for a in sorted_activities if a["status_combined"] == "N/A (等待获取...)"
        ][
            :fetch_count
        ]  # 只获取前 N 条

        log_message(
            f"首页解析完成，共找到 {len(parsed_activities)} 条报名信息。开始获取前 {len(details_to_fetch)} 条活动的详细数据..."
        )

        # 逐个获取活动详情并更新界面
        for i, activity in enumerate(details_to_fetch):
            log_message(f"正在获取第 {i+1}/{len(details_to_fetch)} 条活动的详细数据...")

            detail_data = fetch_detail_data(driver, activity["link"], activity["type"])
            # 更新活动字典中的详情信息
            activity.update(detail_data)

            # 即时更新界面
            display_results(sorted_activities)

        log_message(f"{USER_NAME}同学你好！所有活动详细数据获取完成。")

    except TimeoutException:
        log_message(
            f"等待登录超时 ({WAIT_TIMEOUT}秒)。请重新尝试并确保在时限内完成登录。"
        )
    except WebDriverException as e:
        log_message(f"WebDriver 错误: {e}")
        messagebox.showerror(
            "错误", f"浏览器驱动错误，请确保已安装 ChromeDriver 且版本匹配。"
        )
    except Exception as e:
        log_message(f"程序发生未知错误: {e}")
        messagebox.showerror("错误", f"操作失败: {e}")
    finally:
        # 确保关闭浏览器驱动
        if driver:
            driver.quit()


def sort_activities(activities_list):
    """
    按活动开始时间倒序排序的辅助函数。

    Args:
        activities_list (list): 包含活动信息的字典列表。

    Returns:
        list: 按时间倒序排序后的活动列表。
    """

    def get_sort_key(activity):
        time_str = activity["time_duration"]
        try:
            # 尝试从时间格式 "YYYY-MM-DD HH:MM..." 中提取日期
            match = re.match(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2})", time_str)
            if match:
                # 返回时间戳的负值，实现时间倒序排序（最近的在前）
                return -datetime.strptime(match.group(1), "%Y-%m-%d %H:%M").timestamp()
            return float("inf")  # 无法解析日期的排到最后
        except Exception:
            return float("inf")

    return sorted(activities_list, key=get_sort_key)


def display_results(activities_list):
    """
    在 Treeview 表格和状态栏中显示结果，并根据活动状态设置行背景颜色。

    Args:
        activities_list (list): 包含活动信息的字典列表。
    """
    log_message(f"{USER_NAME}同学你好！ 共找到 {len(activities_list)} 条报名信息。")

    # 清空表格
    for item in tree.get_children():
        tree.delete(item)

    # 确保活动列表已排序
    sorted_activities = sort_activities(activities_list)

    if not sorted_activities:
        tree.insert(
            "", tk.END, values=("未找到任何报名活动。", "N/A", "N/A", "N/A", "N/A")
        )
        return

    # 填充表格
    for activity in sorted_activities:
        kind_part = (
            f"({activity['kind']})"
            if activity["kind"] and activity["kind"] != "N/A"
            else ""
        )
        full_name = f"{activity['name']} {kind_part}".strip()

        row_tag = "default"

        # 设置行背景颜色逻辑
        if activity["status_combined"].endswith("(等待获取...)") or activity[
            "status_combined"
        ].endswith("(请求失败)"):
            # 未获取详情的活动设为浅灰色
            row_tag = "unfetched"
        elif activity["type"] == "Activity":
            if (
                "未签到" in activity["status_combined"]
                or "未签退" in activity["status_combined"]
                or "N/A" in activity["status_combined"]  # 签到信息未知的也设为浅灰色
            ):
                row_tag = "incomplete"
            elif (
                "已签到" in activity["status_combined"]
                and "已签退" in activity["status_combined"]
            ):
                row_tag = "complete"

        tree.insert(
            "",
            tk.END,
            values=(
                full_name,
                activity["time_duration"],
                activity["tags"],
                activity["status_combined"],
                activity["id"],
            ),
            # tags 用于存储链接和颜色信息
            tags=("link", activity["link"], row_tag),
        )

    # 填充链接列表
    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, "所有报名链接 (双击表格行可打开)：\n" + "=" * 50 + "\n")
    for activity in sorted_activities:
        result_text.insert(
            tk.END, f"{activity['name']} (ID: {activity['id']}): {activity['link']}\n"
        )


def start_process_thread():
    """
    读取用户输入的数量，并在单独的线程中运行主要逻辑，避免Tkinter界面卡死。
    """
    try:
        # 尝试将输入转换为整数
        fetch_count = int(num_entries_var.get())
        if fetch_count < 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("输入错误", "请输入一个有效的非负整数作为获取数量。")
        return

    log_message("正在启动浏览器...")
    # 清空现有显示
    for item in tree.get_children():
        tree.delete(item)
    result_text.delete(1.0, tk.END)
    # 禁用按钮，防止重复点击
    start_button.config(state=tk.DISABLED, text="正在运行...")

    def run_wrapper():
        """线程执行的主函数包装器"""
        try:
            get_cookie_and_activities(fetch_count)
        finally:
            # 无论成功失败，恢复按钮状态
            start_button.config(state=tk.NORMAL, text="获取活动信息")

    # 启动新线程
    thread = threading.Thread(target=run_wrapper)
    thread.start()


def log_message(message):
    """
    用于在界面底部状态栏显示运行状态信息。

    Args:
        message (str): 要显示的状态信息。
    """
    status_var.set(f"状态: {message}")
    root.update_idletasks()


def open_link(event):
    """
    处理 Treeview 的双击事件，使用默认浏览器打开对应的活动链接。

    Args:
        event (tk.Event): Tkinter事件对象。
    """
    try:
        # 获取被选中的行
        selected_item = tree.selection()[0]
        # 获取该行的 tags
        tags = tree.item(selected_item, "tags")
        # 标签格式为 ("link", actual_link, row_tag)
        if tags and len(tags) > 1:
            link = tags[1]
            webbrowser.open(link)
            log_message(f"已在默认浏览器中打开链接: {link}")
        else:
            log_message("选择行不包含有效链接。")
    except IndexError:
        # 未选中任何行
        pass
    except Exception as e:
        log_message(f"打开链接失败: {e}")


def open_author_link(event):
    """
    打开作者的 GitHub 链接。
    event=None 允许它可以被 Button 或 bind 调用。
    """
    try:
        AUTHOR_URL = "https://github.com/yeyixiang2007"  # 替换为你的 GitHub 链接
        webbrowser.open_new(AUTHOR_URL)
    except Exception as e:
        messagebox.showerror("错误", f"无法打开链接: {e}")


# ====================================================================
# Tkinter 界面设置
# ====================================================================
root = tk.Tk()
root.title("CUPSecondClassInfoGrabber")

# 配置 Frame for inputs and button
frame_config = ttk.Frame(root)
frame_config.pack(pady=10, padx=10, fill=tk.X)

# 获取数量输入框
tk.Label(
    frame_config,
    text=f"获取活动信息数量:",
    font=("Microsoft YaHei", 10),
).pack(side=tk.LEFT, padx=(0, 5))
num_entries_var = tk.StringVar(value=str(DEFAULT_FETCH_COUNT))
num_entries_entry = ttk.Entry(
    frame_config, textvariable=num_entries_var, width=5, font=("Microsoft YaHei", 10)
)
num_entries_entry.pack(side=tk.LEFT, padx=(0, 20))

# 启动按钮
start_button = tk.Button(
    frame_config,
    text="获取活动信息",
    command=start_process_thread,
    font=("Microsoft YaHei", 14),
)
start_button.pack(side=tk.LEFT, fill=tk.X, expand=True)


# Treeview 表格设置
frame_table = ttk.Frame(root)
frame_table.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

columns = ("#1", "#2", "#3", "#4", "#5")
tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=20)

style = ttk.Style()
style.configure("Treeview", font=("Microsoft YaHei", 10))
style.configure("Treeview.Heading", font=("Microsoft YaHei", 10))

# 表格列标题
tree.heading("#1", text="活动名称", anchor=tk.W)
tree.heading("#2", text="时间", anchor=tk.W)
tree.heading("#3", text="标签", anchor=tk.W)
tree.heading("#4", text="状态", anchor=tk.W)
tree.heading("#5", text="活动ID", anchor=tk.W)

# 表格列宽度
tree.column("#1", width=300, stretch=tk.YES)
tree.column("#2", width=180, stretch=tk.NO)
tree.column("#3", width=180, stretch=tk.NO)
tree.column("#4", width=120, stretch=tk.NO)
tree.column("#5", width=80, stretch=tk.NO)

# 表格行颜色配置
tree.tag_configure("incomplete", background="#FFE6E6")  # 签到未完成
tree.tag_configure("complete", background="#E6FFE6")  # 签到已完成
tree.tag_configure("unfetched", background="#F0F0F0")  # 未获取详情/状态未知（浅灰色）

# 绑定双击事件
tree.bind("<Double-1>", open_link)

# 滚动条
vsb = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
hsb = ttk.Scrollbar(frame_table, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

# 布局
tree.grid(row=0, column=0, sticky="nsew")
vsb.grid(row=0, column=1, sticky="ns")
hsb.grid(row=1, column=0, sticky="ew")

frame_table.grid_columnconfigure(0, weight=1)
frame_table.grid_rowconfigure(0, weight=1)

# 链接列表区域
tk.Label(root, text="链接列表 (双击表格行可打开):", font=("Microsoft YaHei", 10)).pack(
    anchor="w", padx=10, pady=(10, 0)
)
result_text = scrolledtext.ScrolledText(
    root, wrap=tk.WORD, height=10, width=70, font=("Microsoft YaHei", 9)
)
result_text.pack(pady=10, padx=10, fill=tk.X)
result_text.insert(tk.END, "表格数据抓取成功后，链接将在此处显示。")

# 状态栏
status_var = tk.StringVar()
status_var.set("状态: 准备就绪")
status_label = tk.Label(
    root,
    textvariable=status_var,
    bd=1,
    relief=tk.SUNKEN,
    anchor=tk.W,
    font=("Microsoft YaHei", 10),
)
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# 作者水印
author_label = ttk.Label(
    root,
    text="github.com/yeyixiang2007",
    foreground="blue",
    font=("Microsoft YaHei", 10),
    cursor="hand2",
)
author_label.pack(side=tk.BOTTOM, anchor=tk.SE, padx=10, pady=2)  # 放置在右下角
author_label.bind("<Button-1>", open_author_link)

# 启动 Tkinter 主循环
root.mainloop()
