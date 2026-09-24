# -*- coding: utf-8 -*-
"""
大淘客加密商品ID解密工具 - GUI版

功能：
1. 选择 Excel 文件
2. 按 A 列日期筛选（可选）
3. B 列作为加密商品 ID
4. 请求淘宝 uland 商品接口，禁止自动跟随 30x
5. 从 Location / 响应文本中提取真实商品 ID
6. 按设置的随机时间间隔逐条处理
7. 自动导出“原文件名_解密结果.xlsx”
8. GUI 日志实时显示，处理期间不会卡死界面

依赖：
    pip install pandas openpyxl requests

打包 EXE（Windows）：
    pip install pyinstaller
    pyinstaller --onefile --windowed --name "大淘客加密商品ID解密工具" taoke_id_decrypt_gui.py

Excel格式：
    A列：日期
    B列：加密商品ID

如果没有日期筛选，可把“目标筛选日期”留空。
"""

import os
import re
import sys
import time
import random
import threading
import traceback
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

import requests
import pandas as pd


API_URL = "https://uland.taobao.com/item/edetail"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}


def app_base_dir():
    """获取程序所在目录，兼容 .py 和 PyInstaller EXE。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def normalize_id(value):
    """把 Excel 中的商品 ID 尽量转换成干净的字符串。"""
    if pd.isna(value):
        return ""

    # Excel 如果把纯数字读成 12345.0
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]

    return text


def extract_id_from_response(response):
    """
    从 30x 响应中提取真实商品 ID。
    优先 Location，再检查响应正文。
    """
    pattern = re.compile(r"(?:[?&])id=(\d+)", re.I)

    location = response.headers.get("Location", "")
    match = pattern.search(location)
    if match:
        return match.group(1)

    # 某些环境下接口可能把跳转信息放在正文中
    try:
        match = pattern.search(response.text or "")
        if match:
            return match.group(1)
    except Exception:
        pass

    return None


def decode_single_id(taobao_id, session=None):
    """
    解密单个 ID。
    返回：(真实ID, 状态文本)
    """
    if session is None:
        session = requests.Session()

    try:
        response = session.get(
            API_URL,
            params={"id": taobao_id},
            headers=HEADERS,
            allow_redirects=False,
            timeout=15,
        )

        real_id = extract_id_from_response(response)

        if real_id:
            return real_id, f"成功(HTTP {response.status_code})"

        return "", f"失败(HTTP {response.status_code})"

    except requests.exceptions.Timeout:
        return "", "网络错误(请求超时)"
    except requests.exceptions.ConnectionError as e:
        return "", f"网络错误(连接失败): {e}"
    except requests.exceptions.RequestException as e:
        return "", f"网络错误: {e}"
    except Exception as e:
        return "", f"未知错误: {e}"


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("大淘客加密商品ID解密工具")
        self.root.geometry("1000x720")
        self.root.minsize(900, 620)

        self.file_path = ""
        self.running = False
        self.stop_requested = False

        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
        self.min_delay_var = tk.StringVar(value="10")
        self.max_delay_var = tk.StringVar(value="15")
        self.file_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="请选择 Excel 文件")

        self.build_ui()

    # ---------------- UI ----------------

    def build_ui(self):
        root = self.root

        outer = ttk.Frame(root, padding=18)
        outer.pack(fill="both", expand=True)

        title = tk.Label(
            outer,
            text="🚀  大淘客加密商品ID解密工具",
            font=("Microsoft YaHei UI", 18, "bold"),
            anchor="w",
        )
        title.pack(fill="x", pady=(0, 8))

        tip = tk.Label(
            outer,
            text="✅ 工作流：加密商品；忽略表头规则：A列=日期、B列=加密商品id；全部处理完成统一导出文件",
            font=("Microsoft YaHei UI", 10),
            fg="#e88b20",
            anchor="w",
        )
        tip.pack(fill="x", pady=(0, 16))

        form = ttk.Frame(outer)
        form.pack(fill="x")

        # 日期
        ttk.Label(form, text="目标筛选日期：", font=("Microsoft YaHei UI", 11)).grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=7
        )

        self.date_entry = ttk.Entry(
            form, textvariable=self.date_var, width=24, font=("Microsoft YaHei UI", 11)
        )
        self.date_entry.grid(row=0, column=1, sticky="w", pady=7)

        ttk.Label(
            form,
            text="留空默认为昨天，统一格式 yyyy/mm/dd",
            foreground="#e88b20",
            font=("Microsoft YaHei UI", 9),
        ).grid(row=0, column=2, sticky="w", padx=14)

        # 最小间隔
        ttk.Label(form, text="最小间隔(秒)：", font=("Microsoft YaHei UI", 11)).grid(
            row=1, column=0, sticky="w", padx=(0, 12), pady=7
        )
        ttk.Entry(
            form, textvariable=self.min_delay_var, width=10,
            font=("Microsoft YaHei UI", 11)
        ).grid(row=1, column=1, sticky="w", pady=7)

        ttk.Label(form, text="最大间隔(秒)：", font=("Microsoft YaHei UI", 11)).grid(
            row=1, column=2, sticky="w", padx=(70, 12), pady=7
        )
        ttk.Entry(
            form, textvariable=self.max_delay_var, width=10,
            font=("Microsoft YaHei UI", 11)
        ).grid(row=1, column=3, sticky="w", pady=7)

        ttk.Label(
            form,
            text="风控严格可加大数值",
            foreground="#e88b20",
            font=("Microsoft YaHei UI", 9),
        ).grid(row=1, column=4, sticky="w", padx=14)

        # 文件
        ttk.Label(form, text="选择Excel文件：", font=("Microsoft YaHei UI", 11)).grid(
            row=2, column=0, sticky="w", padx=(0, 12), pady=7
        )

        file_frame = ttk.Frame(form)
        file_frame.grid(row=2, column=1, columnspan=4, sticky="ew", pady=7)
        form.columnconfigure(4, weight=1)

        ttk.Button(
            file_frame, text="选择文件", command=self.choose_file
        ).pack(side="left")

        ttk.Label(
            file_frame,
            textvariable=self.file_var,
            font=("Microsoft YaHei UI", 10),
        ).pack(side="left", padx=10)

        # 按钮
        btn_frame = ttk.Frame(outer)
        btn_frame.pack(fill="x", pady=(12, 18))

        self.start_btn = ttk.Button(
            btn_frame, text="开始解密处理", command=self.start_processing
        )
        self.start_btn.pack(side="left", ipadx=10, ipady=4)

        self.clear_btn = ttk.Button(
            btn_frame, text="清空日志", command=self.clear_log
        )
        self.clear_btn.pack(side="left", padx=12, ipadx=10, ipady=4)

        self.stop_btn = ttk.Button(
            btn_frame, text="停止处理", command=self.request_stop, state="disabled"
        )
        self.stop_btn.pack(side="left", ipadx=10, ipady=4)

        ttk.Label(
            btn_frame,
            textvariable=self.status_var,
            foreground="#666666",
        ).pack(side="left", padx=18)

        # 日志
        ttk.Label(
            outer, text="运行日志：", font=("Microsoft YaHei UI", 11)
        ).pack(anchor="w")

        log_frame = ttk.Frame(outer)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            log_frame,
            bg="#1f1f1f",
            fg="#f2f2f2",
            insertbackground="white",
            font=("Consolas", 10),
            wrap="word",
            relief="flat",
            padx=10,
            pady=8,
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.tag_configure("ok", foreground="#66d17a")
        self.log_text.tag_configure("error", foreground="#ff6b6b")
        self.log_text.tag_configure("warn", foreground="#f4c95d")
        self.log_text.tag_configure("info", foreground="#e6e6e6")

    # ---------------- 日志 ----------------

    def log(self, message, level="info"):
        def _write():
            now = time.strftime("%H:%M:%S")
            self.log_text.insert("end", f"[{now}] {message}\n", level)
            self.log_text.see("end")

        self.root.after(0, _write)

    def clear_log(self):
        self.log_text.delete("1.0", "end")

    # ---------------- 文件 ----------------

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls"),
                ("XLSX 文件", "*.xlsx"),
                ("XLS 文件", "*.xls"),
                ("所有文件", "*.*"),
            ],
        )

        if not path:
            return

        self.file_path = path
        self.file_var.set(os.path.basename(path))
        self.status_var.set("文件已选择")

        self.log(f"已选择文件：{os.path.basename(path)}", "ok")

        # 先快速检查 Excel 是否能读取。
        # 这里只检查，不自动开始解密。
        try:
            preview = pd.read_excel(path, nrows=5)
            if preview.shape[1] < 2:
                self.log("⚠️ Excel列数少于2列：预期 A列=日期、B列=加密商品ID", "warn")
            else:
                self.log(
                    f"Excel预检查成功：共检测到 {preview.shape[1]} 列，"
                    f"前两列为“{preview.columns[0]} / {preview.columns[1]}”",
                    "ok",
                )
        except Exception as e:
            self.log(f"❌ Excel预检查失败：{e}", "error")
            self.status_var.set("Excel无法读取")

    # ---------------- 参数 ----------------

    def get_parameters(self):
        date_text = self.date_var.get().strip()

        # 留空时默认昨天
        if not date_text:
            from datetime import timedelta
            date_value = (datetime.now() - timedelta(days=1)).date()
        else:
            date_text = date_text.replace("-", "/").replace(".", "/")
            try:
                date_value = datetime.strptime(date_text, "%Y/%m/%d").date()
            except ValueError:
                raise ValueError("目标筛选日期格式错误，请填写 yyyy/mm/dd，例如 2026/08/26")

        try:
            min_delay = float(self.min_delay_var.get().strip())
            max_delay = float(self.max_delay_var.get().strip())
        except ValueError:
            raise ValueError("最小/最大间隔必须是数字")

        if min_delay < 0 or max_delay < 0:
            raise ValueError("时间间隔不能小于 0")

        if min_delay > max_delay:
            raise ValueError("最小间隔不能大于最大间隔")

        return date_value, min_delay, max_delay

    # ---------------- 处理 ----------------

    def start_processing(self):
        if self.running:
            return

        if not self.file_path:
            messagebox.showwarning("提示", "请先选择 Excel 文件。")
            return

        try:
            target_date, min_delay, max_delay = self.get_parameters()
        except ValueError as e:
            messagebox.showerror("参数错误", str(e))
            return

        self.running = True
        self.stop_requested = False
        self.start_btn.configure(state="disabled")
        self.clear_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set("正在处理...")

        thread = threading.Thread(
            target=self.worker,
            args=(target_date, min_delay, max_delay),
            daemon=True,
        )
        thread.start()

    def request_stop(self):
        if self.running:
            self.stop_requested = True
            self.stop_btn.configure(state="disabled")
            self.log("⚠️ 已请求停止，将在当前请求结束后停止。", "warn")

    def worker(self, target_date, min_delay, max_delay):
        try:
            self.log(f"正在读取文件：{os.path.basename(self.file_path)}")

            df = pd.read_excel(self.file_path)

            if df.shape[1] < 2:
                raise ValueError(
                    "Excel至少需要两列：A列=日期，B列=加密商品ID。"
                )

            date_col = df.columns[0]
            id_col = df.columns[1]

            self.log(
                f"读取成功：共 {len(df)} 行；日期列：{date_col}；ID列：{id_col}",
                "ok",
            )

            # 日期筛选
            date_series = pd.to_datetime(df[date_col], errors="coerce").dt.date

            mask = date_series == target_date
            work_df = df.loc[mask].copy()

            self.log(
                f"目标日期：{target_date.strftime('%Y/%m/%d')}；"
                f"筛选出 {len(work_df)} 条数据",
                "ok",
            )

            if len(work_df) == 0:
                self.log("⚠️ 目标日期没有匹配数据，未发起接口请求。", "warn")
                self.finish(False)
                return

            # 在原表中建立结果列，保持原始所有数据
            if "解密后ID" not in df.columns:
                df["解密后ID"] = ""
            if "解密状态" not in df.columns:
                df["解密状态"] = ""

            session = requests.Session()
            session.headers.update(HEADERS)

            total = len(work_df)
            success_count = 0
            fail_count = 0

            for n, (index, row) in enumerate(work_df.iterrows(), start=1):
                if self.stop_requested:
                    self.log("⏹ 已停止处理。", "warn")
                    break

                raw_id = normalize_id(row[id_col])

                if not raw_id:
                    df.at[index, "解密后ID"] = ""
                    df.at[index, "解密状态"] = "跳过(空值)"
                    self.log(f"[{n}/{total}] 跳过空ID", "warn")
                    continue

                self.log(f"[{n}/{total}] 正在解密：{raw_id} ...")

                real_id, status = decode_single_id(raw_id, session)

                df.at[index, "解密后ID"] = real_id
                df.at[index, "解密状态"] = status

                if real_id:
                    success_count += 1
                    self.log(f"    成功 → {real_id}", "ok")
                else:
                    fail_count += 1
                    self.log(f"    失败 → {status}", "error")

                # 最后一条不需要等待
                if n < total and not self.stop_requested:
                    wait_time = random.uniform(min_delay, max_delay)
                    self.log(f"    等待 {wait_time:.1f} 秒...")
                    # 分段 sleep，让停止按钮响应更快
                    end_time = time.time() + wait_time
                    while time.time() < end_time:
                        if self.stop_requested:
                            break
                        time.sleep(min(0.2, end_time - time.time()))

            # 输出文件
            base, ext = os.path.splitext(self.file_path)
            output_file = f"{base}_解密结果.xlsx"

            df.to_excel(output_file, index=False)

            self.log("=" * 60)
            self.log("🎉 处理完成！", "ok")
            self.log(f"成功：{success_count} 条；失败：{fail_count} 条", "ok")
            self.log(f"📁 结果文件：{output_file}", "ok")
            self.log("=" * 60)

            self.finish(True)

        except Exception as e:
            self.log(f"❌ 程序发生异常：{e}", "error")
            self.log(traceback.format_exc(), "error")
            self.finish(False)

    def finish(self, success):
        def _finish():
            self.running = False
            self.start_btn.configure(state="normal")
            self.clear_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            self.status_var.set("处理完成" if success else "处理结束/失败")

        self.root.after(0, _finish)


def main():
    root = tk.Tk()

    # Windows 下尽量使用系统默认字体
    try:
        root.option_add("*Font", ("Microsoft YaHei UI", 10))
    except Exception:
        pass

    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
