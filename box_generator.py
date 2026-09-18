import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess
import os

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("收纳盒参数SCAD生成器")
        self.geometry("720x620")

        frame_input = ctk.CTkFrame(self)
        frame_input.pack(padx=10, pady=10, fill="x")

        self.label_len = ctk.CTkLabel(frame_input, text="外长度(mm):")
        self.label_len.grid(row=0, column=0, padx=8, pady=8)
        self.entry_len = ctk.CTkEntry(frame_input)
        self.entry_len.insert(0, "120")
        self.entry_len.grid(row=0, column=1, padx=8, pady=8)

        self.label_wid = ctk.CTkLabel(frame_input, text="外宽度(mm):")
        self.label_wid.grid(row=0, column=2, padx=8, pady=8)
        self.entry_wid = ctk.CTkEntry(frame_input)
        self.entry_wid.insert(0, "80")
        self.entry_wid.grid(row=0, column=3, padx=8, pady=8)

        self.label_h = ctk.CTkLabel(frame_input, text="盒子高度(mm):")
        self.label_h.grid(row=1, column=0, padx=8, pady=8)
        self.entry_h = ctk.CTkEntry(frame_input)
        self.entry_h.insert(0, "25")
        self.entry_h.grid(row=1, column=1, padx=8, pady=8)

        self.label_row = ctk.CTkLabel(frame_input, text="横向隔板行数:")
        self.label_row.grid(row=1, column=2, padx=8, pady=8)
        self.entry_row = ctk.CTkEntry(frame_input)
        self.entry_row.insert(0, "2")
        self.entry_row.grid(row=1, column=3, padx=8, pady=8)

        self.label_col = ctk.CTkLabel(frame_input, text="纵向隔板列数:")
        self.label_col.grid(row=2, column=0, padx=8, pady=8)
        self.entry_col = ctk.CTkEntry(frame_input)
        self.entry_col.insert(0, "2")
        self.entry_col.grid(row=2, column=1, padx=8, pady=8)

        frame_btn = ctk.CTkFrame(self)
        frame_btn.pack(padx=10, pady=5, fill="x")
        self.btn_gen = ctk.CTkButton(frame_btn, text="生成OpenSCAD脚本", command=self.generate_scad)
        self.btn_gen.pack(side="left", padx=6, pady=8)
        self.btn_save_scad = ctk.CTkButton(frame_btn, text="保存SCAD文件", command=self.save_scad, state="disabled")
        self.btn_save_scad.pack(side="left", padx=6, pady=8)
        self.btn_export_stl = ctk.CTkButton(frame_btn, text="后台导出STL", command=self.export_stl, state="disabled")
        self.btn_export_stl.pack(side="left", padx=6, pady=8)

        self.text_out = ctk.CTkTextbox(self)
        self.text_out.pack(padx=10, pady=10, fill="both", expand=True)
        self.generated_scad = ""
        self.OPENSCAD_PATH = r"C:\Program Files\OpenSCAD\openscad.exe"

    def generate_scad(self):
        try:
            L = float(self.entry_len.get())
            W = float(self.entry_wid.get())
            H = float(self.entry_h.get())
            R = int(self.entry_row.get())
            C = int(self.entry_col.get())
            wall_t = 1.8
            div_t = 1.6
            inner_L = L - 2*wall_t
            inner_W = W - 2*wall_t
            inner_H = H - wall_t

            scad_text = f'''
$fn=50;
wall_t = {wall_t};
div_t = {div_t};
L={L}; W={W}; H={H};
inner_L = L - 2*wall_t;
inner_W = W - 2*wall_t;
inner_H = H - wall_t;

//底板
translate([wall_t,wall_t,0])
cube([inner_L, inner_W, wall_t]);

//四周外壳侧壁
difference(){{
cube([L,W,H]);
translate([wall_t,wall_t,wall_t])
cube([inner_L, inner_W, H]);
}}

translate([wall_t,wall_t,wall_t])
{{
//横向隔板
for(i = [1:{R}]){{
y_pos = i * inner_W / ({R}+1) - div_t/2;
translate([0, y_pos, 0])
cube([inner_L, div_t, inner_H]);
}}
//纵向隔板
for(j = [1:{C}]){{
x_pos = j * inner_L / ({C}+1) - div_t/2;
translate([x_pos,0,0])
cube([div_t, inner_W, inner_H]);
}}
}}
'''
            self.generated_scad = scad_text
            self.text_out.delete("0.0", tk.END)
            self.text_out.insert(tk.END, self.generated_scad)
            self.btn_save_scad.configure(state="normal")
            self.btn_export_stl.configure(state="normal")
            messagebox.showinfo("完成","参数化SCAD脚本生成完毕！")
        except Exception as e:
            messagebox.showerror("错误",str(e))

    def save_scad(self):
        fp = filedialog.asksaveasfilename(defaultextension=".scad", filetypes=[("OpenSCAD","*.scad")])
        if fp:
            with open(fp,"w",encoding="utf-8") as f:
                f.write(self.generated_scad)

    def export_stl(self):
        if not os.path.exists(self.OPENSCAD_PATH):
            messagebox.showerror("路径错误","openscad路径不对，请修改代码")
            return
        scad_tmp="_temp_box.scad"
        stl_out = filedialog.asksaveasfilename(defaultextension=".stl", filetypes=[("STL model","*.stl")])
        if not stl_out:
            return
        with open(scad_tmp,"w",encoding="utf-8") as f:
            f.write(self.generated_scad)
        try:
            subprocess.run([self.OPENSCAD_PATH,"-o",stl_out,scad_tmp],check=True)
            messagebox.showinfo("导出成功",f"STL输出完成：{stl_out}")
        except Exception as e:
            messagebox.showerror("导出失败",str(e))
        finally:
            if os.path.exists(scad_tmp):
                os.remove(scad_tmp)


if __name__ == "__main__":
    app = App()
    app.mainloop()