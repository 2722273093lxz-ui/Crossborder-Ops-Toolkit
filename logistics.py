import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import requests
import json
import time

# ===================== 配置区（请在这里填写你的API凭证）=====================
# FedEx 配置
FEDEX_CLIENT_ID = "你的FedEx Client ID"
FEDEX_CLIENT_SECRET = "你的FedEx Client Secret"
FEDEX_TOKEN_URL = "https://apis.fedex.com/oauth/token"
FEDEX_TRACK_URL = "https://apis.fedex.com/track/v1/trackingnumbers"

# USPS 配置
USPS_CONSUMER_KEY = "你的USPS Consumer Key"
USPS_CONSUMER_SECRET = "你的USPS Consumer Secret"
USPS_TOKEN_URL = "https://apis.usps.com/oauth2/token"
USPS_TRACK_URL = "https://apis.usps.com/track/v3/trackingnumbers"

# 请求配置
REQUEST_TIMEOUT = 10  # 单次请求超时时间（秒）
REQUEST_INTERVAL = 1  # 两次请求间隔（秒，避免限流）
# ==========================================================================

class LogisticsBatchChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("FedEx / USPS 物流批量查询工具")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 美化主题
        style = ttk.Style()
        style.theme_use("clam")
        
        # 全局变量
        self.df = None
        self.file_path = ""
        self.fedex_token = ""
        self.usps_token = ""
        self.fedex_token_expire = 0
        self.usps_token_expire = 0

        # 标题
        header_label = ttk.Label(
            root, 
            text="📦 跨境物流批量查询工具（FedEx / USPS）", 
            font=("Microsoft YaHei", 16, "bold")
        )
        header_label.pack(pady=12)

        # 1. 文件导入区
        frame_file = ttk.LabelFrame(root, text="📁 Excel文件导入")
        frame_file.pack(fill="x", padx=16, pady=8)

        ttk.Button(frame_file, text="选择Excel文件", command=self.select_file).grid(row=0, column=0, padx=8, pady=8)
        self.file_label = ttk.Label(frame_file, text="未选择文件", foreground="gray")
        self.file_label.grid(row=0, column=1, padx=8, pady=8)

        # 2. 列选择区
        frame_col = ttk.LabelFrame(root, text="⚙️ 查询配置")
        frame_col.pack(fill="x", padx=16, pady=8)

        ttk.Label(frame_col, text="单号所在列：").grid(row=0, column=0, padx=8, pady=8)
        self.col_combo = ttk.Combobox(frame_col, width=20, state="disabled")
        self.col_combo.grid(row=0, column=1, padx=8, pady=8)

        ttk.Label(frame_col, text="快递公司：").grid(row=0, column=2, padx=8, pady=8)
        self.carrier_combo = ttk.Combobox(frame_col, width=15, state="readonly")
        self.carrier_combo["values"] = ["FedEx", "USPS"]
        self.carrier_combo.current(0)
        self.carrier_combo.grid(row=0, column=3, padx=8, pady=8)

        # 3. 操作按钮区
        frame_btn = ttk.Frame(root)
        frame_btn.pack(pady=10)
        self.start_btn = ttk.Button(frame_btn, text="开始批量查询", command=self.start_batch_query, state="disabled")
        self.start_btn.grid(row=0, column=0, padx=6)
        ttk.Button(frame_btn, text="清空结果", command=self.clear_result).grid(row=0, column=1, padx=6)
        ttk.Button(frame_btn, text="导出结果Excel", command=self.export_result).grid(row=0, column=2, padx=6)

        # 4. 进度条
        self.progress = ttk.Progressbar(root, orient="horizontal", length=800, mode="determinate")
        self.progress.pack(pady=8)

        # 5. 结果输出区
        frame_result = ttk.LabelFrame(root, text="📋 查询结果")
        frame_result.pack(fill="both", expand=True, padx=16, pady=8)
        self.result_text = scrolledtext.ScrolledText(frame_result, font=("Microsoft YaHei", 10))
        self.result_text.pack(padx=8, pady=8, fill="both", expand=True)

        # 状态提示
        self.status_label = ttk.Label(root, text="就绪", foreground="green")
        self.status_label.pack(pady=4)

    # ---------------------- 文件处理 ----------------------
    def select_file(self):
        """选择Excel文件"""
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=(("Excel文件", "*.xlsx"), ("所有文件", "*.*"))
        )
        if not self.file_path:
            return
        
        try:
            # 读取Excel
            self.df = pd.read_excel(self.file_path)
            # 更新列下拉框
            self.col_combo["values"] = list(self.df.columns)
            self.col_combo.state(["!disabled"])
            self.col_combo.current(0)
            # 更新文件标签
            self.file_label.config(text=f"已选择：{self.file_path.split('/')[-1]}", foreground="black")
            self.status_label.config(text="文件读取成功，请选择单号列和快递公司", foreground="green")
            # 清空结果
            self.clear_result()
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件失败：{str(e)}")
            self.status_label.config(text="文件读取失败", foreground="red")

    # ---------------------- API Token 获取 ----------------------
    def get_fedex_token(self):
        """获取FedEx OAuth Token"""
        # 检查token是否未过期
        if self.fedex_token and time.time() < self.fedex_token_expire:
            return True
        
        try:
            payload = {
                "grant_type": "client_credentials",
                "client_id": FEDEX_CLIENT_ID,
                "client_secret": FEDEX_CLIENT_SECRET
            }
            response = requests.post(FEDEX_TOKEN_URL, data=payload, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            token_data = response.json()
            self.fedex_token = token_data["access_token"]
            # token有效期1小时，提前5分钟过期
            self.fedex_token_expire = time.time() + token_data["expires_in"] - 300
            return True
        except Exception as e:
            self.result_text.insert(tk.END, f"❌ FedEx Token获取失败：{str(e)}\n")
            return False

    def get_usps_token(self):
        """获取USPS OAuth Token"""
        # 检查token是否未过期
        if self.usps_token and time.time() < self.usps_token_expire:
            return True
        
        try:
            payload = {
                "grant_type": "client_credentials",
                "client_id": USPS_CONSUMER_KEY,
                "client_secret": USPS_CONSUMER_SECRET
            }
            response = requests.post(USPS_TOKEN_URL, data=payload, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            token_data = response.json()
            self.usps_token = token_data["access_token"]
            # token有效期1小时，提前5分钟过期
            self.usps_token_expire = time.time() + token_data["expires_in"] - 300
            return True
        except Exception as e:
            self.result_text.insert(tk.END, f"❌ USPS Token获取失败：{str(e)}\n")
            return False

    # ---------------------- 单条物流查询 ----------------------
    def query_fedex_track(self, tracking_num):
        """查询FedEx物流信息"""
        if not self.get_fedex_token():
            return {"error": "Token获取失败"}
        
        try:
            headers = {
                "Authorization": f"Bearer {self.fedex_token}",
                "Content-Type": "application/json"
            }
            payload = {
                "trackingNumbers": [tracking_num],
                "includeDetailedScans": True
            }
            response = requests.post(
                FEDEX_TRACK_URL, 
                headers=headers, 
                json=payload, 
                timeout=REQUEST_TIMEOUT
            )
            response.raise_for_status()
            data = response.json()
            
            # 解析返回结果
            if "trackResults" in data and len(data["trackResults"]) > 0:
                track_result = data["trackResults"][0]
                latest_status = track_result.get("latestStatusDetail", {}).get("statusDescription", "无状态")
                latest_location = track_result.get("latestStatusDetail", {}).get("location", {}).get("city", "无位置")
                scan_list = track_result.get("scanDetails", [])
                
                # 整理轨迹
                track_info = []
                for scan in scan_list:
                    scan_time = scan.get("date", "") + " " + scan.get("time", "")
                    scan_status = scan.get("statusDescription", "")
                    scan_location = scan.get("location", {}).get("city", "")
                    track_info.append(f"[{scan_time}] {scan_status} - {scan_location}")
                
                return {
                    "status": latest_status,
                    "location": latest_location,
                    "track": "\n".join(track_info),
                    "error": ""
                }
            else:
                return {"error": "未找到该单号的物流信息"}
                
        except Exception as e:
            return {"error": f"查询失败：{str(e)}"}

    def query_usps_track(self, tracking_num):
        """查询USPS物流信息"""
        if not self.get_usps_token():
            return {"error": "Token获取失败"}
        
        try:
            headers = {
                "Authorization": f"Bearer {self.usps_token}",
                "Content-Type": "application/json"
            }
            payload = {
                "trackingNumbers": [tracking_num],
                "includeDetailedScans": True
            }
            response = requests.post(
                USPS_TRACK_URL, 
                headers=headers, 
                json=payload, 
                timeout=REQUEST_TIMEOUT
            )
            response.raise_for_status()
            data = response.json()
            
            # 解析返回结果
            if "trackResults" in data and len(data["trackResults"]) > 0:
                track_result = data["trackResults"][0]
                latest_status = track_result.get("latestStatusDetail", {}).get("statusDescription", "无状态")
                latest_location = track_result.get("latestStatusDetail", {}).get("location", {}).get("city", "无位置")
                scan_list = track_result.get("scanDetails", [])
                
                # 整理轨迹
                track_info = []
                for scan in scan_list:
                    scan_time = scan.get("date", "") + " " + scan.get("time", "")
                    scan_status = scan.get("statusDescription", "")
                    scan_location = scan.get("location", {}).get("city", "")
                    track_info.append(f"[{scan_time}] {scan_status} - {scan_location}")
                
                return {
                    "status": latest_status,
                    "location": latest_location,
                    "track": "\n".join(track_info),
                    "error": ""
                }
            else:
                return {"error": "未找到该单号的物流信息"}
                
        except Exception as e:
            return {"error": f"查询失败：{str(e)}"}

    # ---------------------- 批量查询 ----------------------
    def start_batch_query(self):
        """开始批量查询"""
        if self.df is None:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return
        
        selected_col = self.col_combo.get()
        if not selected_col:
            messagebox.showwarning("提示", "请选择单号所在列")
            return
        
        carrier = self.carrier_combo.get()
        tracking_list = self.df[selected_col].dropna().astype(str).tolist()
        
        if not tracking_list:
            messagebox.showwarning("提示", "选中的列没有有效单号")
            return
        
        # 确认查询
        if not messagebox.askyesno("确认", f"即将查询 {len(tracking_list)} 个{carrier}单号，是否继续？"):
            return
        
        # 初始化进度
        self.progress["maximum"] = len(tracking_list)
        self.progress["value"] = 0
        self.start_btn.state(["disabled"])
        self.status_label.config(text=f"正在批量查询{carrier}物流...", foreground="blue")
        
        # 清空结果
        self.clear_result()
        
        # 批量查询
        result_list = []
        for i, tracking_num in enumerate(tracking_list):
            self.result_text.insert(tk.END, f"🔍 正在查询第 {i+1}/{len(tracking_list)} 个单号：{tracking_num}\n")
            self.result_text.see(tk.END)
            
            # 查询物流
            if carrier == "FedEx":
                result = self.query_fedex_track(tracking_num)
            else:
                result = self.query_usps_track(tracking_num)
            
            # 整理结果
            if result["error"]:
                result_list.append({
                    "单号": tracking_num,
                    "最新状态": "查询失败",
                    "最新位置": "",
                    "详细轨迹": result["error"]
                })
                self.result_text.insert(tk.END, f"❌ {tracking_num} 查询失败：{result['error']}\n\n")
            else:
                result_list.append({
                    "单号": tracking_num,
                    "最新状态": result["status"],
                    "最新位置": result["location"],
                    "详细轨迹": result["track"]
                })
                self.result_text.insert(tk.END, f"✅ {tracking_num} 查询完成，最新状态：{result['status']}\n\n")
            
            # 更新进度
            self.progress["value"] = i + 1
            self.root.update_idletasks()
            
            time.sleep(REQUEST_INTERVAL)
        
        # 保存结果到全局
        self.result_df = pd.DataFrame(result_list)
        
        # 完成提示
        self.start_btn.state(["!disabled"])
        self.status_label.config(text=f"批量查询完成，共查询 {len(result_list)} 个单号", foreground="green")
        messagebox.showinfo("完成", f"批量查询完成！成功查询 {len(result_list)} 个单号，可点击导出结果Excel")

    # ---------------------- 结果处理 ----------------------
    def clear_result(self):
        """清空结果"""
        self.result_text.delete("1.0", tk.END)
        self.progress["value"] = 0
        self.result_df = None

    def export_result(self):
        """导出结果到Excel"""
        if not hasattr(self, "result_df") or self.result_df is None:
            messagebox.showwarning("提示", "暂无查询结果可导出")
            return
        
        save_path = filedialog.asksaveasfilename(
            title="保存结果Excel",
            defaultextension=".xlsx",
            filetypes=(("Excel文件", "*.xlsx"), ("所有文件", "*.*"))
        )
        
        if not save_path:
            return
        
        try:
            self.result_df.to_excel(save_path, index=False)
            messagebox.showinfo("成功", f"结果已导出到：{save_path}")
            self.status_label.config(text="结果导出成功", foreground="green")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")
            self.status_label.config(text="导出失败", foreground="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = LogisticsBatchChecker(root)
    root.mainloop()
