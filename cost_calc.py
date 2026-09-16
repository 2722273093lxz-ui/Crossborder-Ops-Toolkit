import tkinter as tk
from tkinter import messagebox

def calculate_cost():
    try:
        # 获取输入的售价
        price = float(entry_price.get())
        
        # 汇率
        exchange_rate = 6.9
        
        # 物流费用
        logistics_cost = 800
        
        # 国际费率
        international_rate = 0.0135
        
        # 折扣
        discount = 0.05
        
        # 平台扣点
        platform_cut = 0.085
        
        # 流量
        traffic = 0.1
        
        # 计算投流前的总成本
        total_cost_no_traffic = price * (1 - discount - international_rate - platform_cut) + logistics_cost * exchange_rate
        
        # 计算投流后的总成本
        total_cost_with_traffic = total_cost_no_traffic + price * traffic
        
        # 显示结果
        label_no_traffic.config(text=f"投流前的总成本: {total_cost_no_traffic:.2f} 元")
        label_with_traffic.config(text=f"投流后的总成本: {total_cost_with_traffic:.2f} 元")
    except ValueError:
        messagebox.showerror("输入错误", "请输入有效的数字")

# 创建主窗口
root = tk.Tk()
root.title("竞品成本分析")
root.geometry("400x300")

# 设置淡蓝色背景
root.configure(bg="#ADD8E6")

# 创建标签和输入框
label_price = tk.Label(root, text="售价:", bg="#ADD8E6")
label_price.pack(pady=5)
entry_price = tk.Entry(root, width=20)
entry_price.pack(pady=5)

# 创建计算按钮
button_calculate = tk.Button(root, text="计算成本", command=calculate_cost)
button_calculate.pack(pady=10)

# 创建显示标签
label_no_traffic = tk.Label(root, text="投流前的总成本:", bg="#ADD8E6")
label_no_traffic.pack(pady=5)
label_with_traffic = tk.Label(root, text="投流后的总成本:", bg="#ADD8E6")
label_with_traffic.pack(pady=5)

# 运行程序
root.mainloop()