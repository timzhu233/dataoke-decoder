import os
import re
import sys
import time
import random
import requests
import pandas as pd

# ================= 配置区域 =================
# 自动获取程序所在目录（兼容打包后的 exe 和直接运行 py 脚本）
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)      # 打包后的 exe
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 直接运行 py

INPUT_FILE  = os.path.join(BASE_DIR, "待解密数据.xlsx")   # 输入文件名
OUTPUT_FILE = os.path.join(BASE_DIR, "解密结果.xlsx")     # 输出文件名

# 淘宝客解密接口（来自截图）
API_URL = "https://uland.taobao.com/item/edetail"

# 请求头：模拟浏览器，防止被拦截
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

MIN_DELAY = 3   # 最小间隔时间（秒）
MAX_DELAY = 5   # 最大间隔时间（秒）
# ============================================


def extract_id_from_response(response):
    """
    基于截图逻辑：从 302 响应中提取真实商品 ID
    优先从 response.text 匹配，其次从 headers 的 Location 中匹配
    """
    if response.status_code == 302:
        pattern = r'id=(\d+)'

        # 方案 1：从响应文本里正则匹配（和截图一致）
        match = re.search(pattern, response.text)
        if match:
            return match.group(1)

        # 方案 2：从 headers 的 Location 里匹配（更严谨）
        location = response.headers.get('Location', '')
        match_loc = re.search(pattern, location)
        if match_loc:
            return match_loc.group(1)

    return None


def decode_single_id(taobao_id):
    """
    解密单个 ID，返回 (真实ID, 状态信息)
    """
    params = {"id": taobao_id}
    try:
        # allow_redirects=False 阻止自动跳转，方便我们捕获 302
        response = requests.get(
            API_URL,
            params=params,
            headers=HEADERS,
            allow_redirects=False,
            timeout=10
        )
        real_id = extract_id_from_response(response)

        if real_id:
            return real_id, "成功"
        else:
            return "", f"失败(状态码:{response.status_code})"

    except requests.exceptions.RequestException as e:
        return "", f"网络错误: {str(e)}"
    except Exception as e:
        return "", f"未知错误: {str(e)}"


def main():
    # 设置窗口标题（仅 Windows 有效，Mac 会自动跳过）
    if os.name == 'nt':
        os.system("title 淘宝客 ID 自动解密工具")

    print("=" * 50)
    print("淘宝客 ID 自动解密工具 (基于 302 重定向逻辑)")
    print("请勿关闭此窗口，等待执行完毕...")
    print("=" * 50)

    # 1. 检查文件是否存在
    if not os.path.exists(INPUT_FILE):
        print(f"❌ 找不到文件: {INPUT_FILE}")
        print("请确保 Excel 文件与程序在同一目录下，且文件名正确。")
        input("\n按回车键退出...")
        return

    # 2. 读取 Excel
    try:
        df = pd.read_excel(INPUT_FILE)
        id_column = df.columns[0]  # 默认读取第一列
        print(f"✅ 成功读取 Excel，共 {len(df)} 条数据。")
        print(f"   正在处理第一列: '{id_column}'")
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        input("\n按回车键退出...")
        return

    # 准备结果列
    real_ids = []
    statuses = []

    # 3. 循环处理
    for index, row in df.iterrows():
        raw_id = str(row[id_column]).strip()

        # 跳过空值
        if raw_id == 'nan' or raw_id == '':
            real_ids.append("")
            statuses.append("跳过(空值)")
            continue

        print(f"[{index + 1}/{len(df)}] 正在解密: {raw_id} ...", end=" ", flush=True)

        real_id, status = decode_single_id(raw_id)
        real_ids.append(real_id)
        statuses.append(status)

        if real_id:
            print(f"成功 -> {real_id}")
        else:
            print(f"失败 -> {status}")

        # 防风控随机等待
        wait_time = random.uniform(MIN_DELAY, MAX_DELAY)
        time.sleep(wait_time)

    # 4. 写入结果并保存
    df['解密后ID'] = real_ids
    df['解密状态'] = statuses

    try:
        df.to_excel(OUTPUT_FILE, index=False)
        print("=" * 50)
        print(f"🎉 所有任务执行完毕！")
        print(f"📁 解密结果已保存为: {OUTPUT_FILE}")
        print("=" * 50)
    except Exception as e:
        print(f"❌ 保存结果文件失败，请检查文件是否被占用: {e}")

    input("\n按回车键退出...")
    


if __name__ == "__main__":
    main()