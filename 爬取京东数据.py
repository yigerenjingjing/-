import requests
from DrissionPage import Chromium,ChromiumOptions
from DrissionPage.errors import PageDisconnectedError, ElementNotFoundError
from bs4 import BeautifulSoup
import time
import json
import pandas as pd
import os

# 目标网址（京东宠物玩具搜索结果页）
url = "https://search.jd.com/Search?keyword=%E5%AE%A0%E7%89%A9%E7%8E%A9%E5%85%B7&enc=utf-8&wq=%E5%AE%A0%E7%89%A9wan%27jv&pvid=a31833d4f6ba4a038b51e97e9aa91ab7"

headers = {
    # 模拟浏览器头部，防止被反爬
    "User-Agent": "自己的User-Agent",
}

#保存文件地址
output_path = r"D:\jd_result.xlsx"

# 浏览器管理函数
def initialize_browser():
    """初始化浏览器、打开页面、点击排序，并返回实例和主标签页。"""
    print("\n[系统] 正在初始化新的浏览器实例...")
    dp = Chromium()
    tab = dp.new_tab()
    tab.get(url)
    tab.wait.load_start()
    print("[系统] 正在等待并点击销量排序...")
    tab.wait.ele_displayed('text:销量', timeout=15)
    tab.ele('text:销量').click()
    time.sleep(2)
    print("[系统] 浏览器初始化完成，已进入商品列表页。")
    return dp, tab

# 主程序
# 1. 初始化浏览器并获取商品列表
dp, tab = initialize_browser()

print("\n[系统] 正在获取商品列表...")
tab.wait.ele_displayed('css:._text_1x4i2_30', timeout=15)
eles = tab.eles('css:._text_1x4i2_30')[:4]
content = [e.text for e in eles]
eles = tab.eles('css:._price_1tn4o_13')[:4]
price = [e.text for e in eles]
eles = tab.eles('css:._tags_hzhkm_2')
huodong = [(e.text or '').strip().replace('\n', '/') for e in eles[:4]]
eles = tab.eles('css:._name_d19t5_35')[:4]
store_name = [e.text for e in eles]
eles = tab.eles('css:._goods_volume_1xkku_1')[:4]
goods_volume = [e.text for e in eles]
print("[系统] 商品列表获取完成 (仅获取前4个)。")

# 2. 创建总览DataFrame和用于存储评论的字典
all_products_data = {
    '商品名称': content, '价格': price, '活动': huodong,
    '店铺名称': store_name, '销量': goods_volume
}
all_products_df = pd.DataFrame(all_products_data)
all_comments_data = {} # 用于在内存中存储所有评论

# 3. 使用while循环爬取每个商品的评论
i = 0
while i < len(content):
    product_name = content[i]
    print(f"\n--- 开始处理第 {i + 1}/{len(content)} 个商品: {product_name} ---")
    
    try:
        # 等待并点击商品进入详情页
        tab.wait.ele_displayed(f'text:{product_name}', timeout=15)
        tab.ele(f'text:{product_name}').click()
        new_tab = dp.latest_tab
        print("已进入商品详情页...")
        time.sleep(2)

        # 等待并点击“全部评价”
        new_tab.wait.ele_displayed('css:.all-btn', timeout=15)
        new_tab.ele('css:.all-btn').click()
        time.sleep(2)
        comments_page = dp.latest_tab
        comments_page.wait.load_start()

        # 等待并爬取评论
        comments_page.wait.ele_displayed('.jdc-pc-rate-card-main-desc', timeout=15)
        all_comments_texts = set()
        retries = 5
        last_comment_count = 0
        while retries > 0:
            current_comments_eles = comments_page.eles('css:.jdc-pc-rate-card-main-desc')
            for comment in current_comments_eles:
                all_comments_texts.add(comment.text)

            if len(all_comments_texts) > last_comment_count:
                last_comment_count = len(all_comments_texts)
                retries = 5
            else:
                retries -= 1
            
            if len(all_comments_texts) >= 20:
                break

            if current_comments_eles:
                current_comments_eles[-1].scroll.to_see()
            else:
                break
            time.sleep(1.5)

        # 数据清洗并存入内存字典
        if all_comments_texts:
            cleaned_texts = [
                text.strip() for text in all_comments_texts 
                if text.strip() != "此用户未填写评价内容" and len(text.strip()) >= 5
            ]
            if cleaned_texts:
                all_comments_data[product_name] = cleaned_texts
                print(f"成功采集 {len(cleaned_texts)} 条评论到内存。")

        # 任务成功，返回列表页准备下一个
        print("正在返回商品搜索列表页...")
        all_tabs = dp.get_tabs()
        if len(all_tabs) > 1:
            for t in all_tabs[1:]:
                t.close()
        tab.get(url)
        tab.wait.load_start()
        tab.ele('text:销量').click()
        time.sleep(2)
        
        i += 1

    except PageDisconnectedError as e:
        print(f"[严重错误] 与浏览器连接已断开: {e}")
        print("[恢复流程] 正在重启浏览器并跳过当前商品...")
        try:
            dp.quit()
        except:
            pass
        dp, tab = initialize_browser()
        i += 1

    except Exception as e:
        print(f"[未知错误] 处理 '{product_name}' 时发生错误: {e}")
        print("[恢复流程] 正在重启浏览器并跳过当前商品...")
        try:
            dp.quit()
        except:
            pass
        dp, tab = initialize_browser()
        i += 1

print("\n--- 所有商品评论爬取任务已尝试执行完毕 ---")
dp.quit()

# 4. 所有爬取任务结束后，一次性写入Excel文件
print("\n[系统] 开始将所有采集数据一次性写入Excel文件...")
try:
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        all_products_df.to_excel(writer, sheet_name='商品总览', index=False)
        print("已写入 '商品总览' 工作表。")
        
        for product_name, comments in all_comments_data.items():
            df_comments = pd.DataFrame(comments, columns=['评论'])
            sheet_name = product_name.replace('/', '-').replace('\\', '-').replace('?', '').replace('*', '').replace('[', '').replace(']', '')[:31]
            df_comments.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"已写入 '{sheet_name}' 工作表。")
    print(f"[成功] 数据已全部保存至 '{output_path}'")
except Exception as e:
    print(f"[失败] 写入Excel文件时发生错误: {e}")
