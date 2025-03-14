import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
import colour
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

class ColorCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("颜色计算器")
        self.root.geometry("800x800")
        
        # 创建Notebook实现多页面
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 第一页：计算页面
        self.calc_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.calc_frame, text="计算")
        
        # 第二页：绘图页面
        self.plot_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.plot_frame, text="绘图分析")
        
        # 第一页组件
        self.create_file_selection(self.calc_frame)
        self.create_result_display(self.calc_frame)
        self.create_color_preview(self.calc_frame)
        self.create_buttons(self.calc_frame)
        
        # 第二页组件
        self.create_plot_area(self.plot_frame)
        
    def create_file_selection(self, parent):
        """创建文件选择区域"""
        file_frame = ttk.LabelFrame(parent, text="文件选择")
        file_frame.pack(fill=tk.X, pady=5)
        
        # Excel文件选择
        ttk.Label(file_frame, text="Excel文件:").grid(row=0, column=0, padx=5, pady=5)
        self.excel_entry = ttk.Entry(file_frame, width=50)
        self.excel_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="浏览...", command=self.select_excel).grid(row=0, column=2, padx=5, pady=5)
        
        # 输出目录选择
        ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, padx=5, pady=5)
        self.output_entry = ttk.Entry(file_frame, width=50)
        self.output_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="浏览...", command=self.select_output_dir).grid(row=1, column=2, padx=5, pady=5)
        
    def create_result_display(self, parent):
        """创建结果显示区域"""
        result_frame = ttk.LabelFrame(parent, text="计算结果")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def create_color_preview(self, parent):
        """创建颜色预览区域"""
        preview_frame = ttk.LabelFrame(parent, text="颜色预览")
        preview_frame.pack(fill=tk.BOTH, pady=5)
        
        # 创建5个颜色显示区域
        self.color_frames = []
        self.color_labels = []
        self.color_canvases = []
        
        for i in range(5):
            color_frame = ttk.Frame(preview_frame)
            color_frame.pack(fill=tk.X, pady=2)
            
            # 颜色显示
            fig = plt.Figure(figsize=(2, 0.5))
            canvas = FigureCanvasTkAgg(fig, master=color_frame)
            canvas.get_tk_widget().pack(side=tk.LEFT, padx=5)
            self.color_canvases.append(canvas)
            
            # Lab值显示
            lab_label = ttk.Label(color_frame, text="L: -, a: -, b: -\nHue: -°")
            lab_label.pack(side=tk.LEFT, padx=5)
            self.color_labels.append(lab_label)
            
            self.color_frames.append(color_frame)
        
    def create_buttons(self, parent):
        """创建操作按钮"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="开始计算", command=self.calculate).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存结果", command=self.save_results).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=tk.RIGHT, padx=5)
        
    def select_excel(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls")]
        )
        if file_path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)
            
    def select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, dir_path)
            
    def calculate(self):
        """执行计算"""
        excel_path = self.excel_entry.get()
        if not excel_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
            
        try:
            # 清空之前的结果
            self.result_text.delete(1.0, tk.END)
            
            # 读取并处理Excel文件
            sheet_names = ["1.0", "1.1", "1.2", "1.3", "1.4"]
            results = {}
            
            for i, sheet in enumerate(sheet_names):
                try:
                    df = pd.read_excel(excel_path, sheet_name=sheet)
                    df.columns = df.columns.str.strip()
                    
                    if "Wavelength" not in df.columns or "Reflectance" not in df.columns:
                        self.result_text.insert(tk.END, f"⚠️ {sheet} 缺少必要的列，跳过处理。\n")
                        continue
                        
                    # 转换数据格式
                    wavelengths = df["Wavelength"].astype(float).values
                    reflectance = df["Reflectance"].astype(float).values
                    
                    # 组织光谱数据
                    spectrum_data = dict(zip(wavelengths, reflectance))
                    
                    # 创建 SpectralDistribution 并插值
                    sd = colour.SpectralDistribution(spectrum_data)
                    regular_wavelengths = SliceWrapper(380, 781, 5)  # 5nm 采样
                    sd = sd.interpolate(regular_wavelengths)
                    
                    # 计算 CIE XYZ 值
                    illuminant = colour.SDS_ILLUMINANTS["D65"]
                    XYZ = colour.sd_to_XYZ(sd, illuminant=illuminant)
                    
                    # 获取 D65 白点
                    whitepoint = colour.CCS_ILLUMINANTS["CIE 1931 2 Degree Standard Observer"]["D65"]
                    
                    # 计算 CIELab
                    Lab = colour.XYZ_to_Lab(XYZ, illuminant=whitepoint)
                    L_val, a_val, b_val = Lab
                    
                    # 计算 sRGB 颜色
                    rgb = colour.XYZ_to_sRGB(XYZ / 100)
                    rgb = np.clip(rgb, 0, 1)
                    hex_color = rgb_to_hex(rgb)
                    
                    # 计算HSV值
                    hsv = colour.RGB_to_HSV(rgb)
                    hue = round(hsv[0] * 360, 2)  # 转换为0-360度
                    
                    # 计算ΔE值（与第一个样本比较）
                    if i > 0:
                        ref_Lab = results[sheet_names[0]]["CIELab"]
                        delta_E = colour.delta_E(Lab, ref_Lab)
                    else:
                        delta_E = 0.0
                    
                    # 更新颜色显示
                    self.update_color_display(i, rgb, Lab, hue)
                    
                    # 存储结果
                    results[sheet] = {
                        "CIELab": Lab,
                        "sRGB": rgb,
                        "Hex Color": hex_color,
                        "Hue": hue,
                        "ΔE": delta_E
                    }
                    
                    # 显示结果
                    self.result_text.insert(tk.END, f"\n--- {sheet} 计算结果 ---\n")
                    self.result_text.insert(tk.END, f"CIELab: {Lab}\n")
                    self.result_text.insert(tk.END, f"十六进制颜色: {hex_color}\n")
                    self.result_text.insert(tk.END, f"HSV Hue: {hue}°\n")
                    self.result_text.insert(tk.END, f"ΔE: {delta_E:.2f}\n")
                    
                except Exception as e:
                    self.result_text.insert(tk.END, f"❌ 读取 {sheet} 时发生错误: {e}\n")
                    
            # 存储原始光谱数据
            self.spectra_data = {}
            for sheet in sheet_names:
                if sheet in results:
                    df = pd.read_excel(excel_path, sheet_name=sheet)
                    self.spectra_data[sheet] = {
                        'wavelengths': df["Wavelength"].astype(float).values,
                        'reflectance': df["Reflectance"].astype(float).values
                    }
            
            self.results = results
            
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件时发生错误: {e}")
            
    def update_color_display(self, index, rgb, Lab, hue):
        """更新颜色显示区域"""
        # 更新颜色显示
        fig = self.color_canvases[index].figure
        fig.clear()
        ax = fig.add_subplot(111)
        color = np.array([[rgb]])
        ax.imshow(color)
        ax.axis("off")
        self.color_canvases[index].draw()
        
        # 更新Lab值和Hue值显示
        L_val, a_val, b_val = Lab
        self.color_labels[index].config(text=f"L: {L_val:.2f}, a: {a_val:.2f}, b: {b_val:.2f}\nHue: {hue}°")
        
    def create_plot_area(self, parent):
        """创建绘图分析区域"""
        plot_frame = ttk.LabelFrame(parent, text="颜色分析图")
        plot_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建绘图区域
        self.plot_fig = plt.Figure(figsize=(8, 4))
        self.plot_ax = self.plot_fig.add_subplot(111)
        self.plot_canvas = FigureCanvasTkAgg(self.plot_fig, master=plot_frame)
        self.plot_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 创建控制面板
        control_frame = ttk.Frame(plot_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        # 添加绘图类型选择
        ttk.Label(control_frame, text="绘图类型:").pack(side=tk.LEFT, padx=5)
        self.plot_type = tk.StringVar(value="Lab")
        plot_types = ["Lab", "RGB", "HSV", "ΔE", "反射光谱"]
        for ptype in plot_types:
            ttk.Radiobutton(control_frame, text=ptype, variable=self.plot_type, 
                           value=ptype).pack(side=tk.LEFT, padx=2)
            
        # 添加绘图按钮
        ttk.Button(control_frame, text="更新图表", 
                  command=self.update_plot).pack(side=tk.RIGHT, padx=5)
        
    def update_plot(self):
        """更新绘图"""
        if not hasattr(self, 'results'):
            messagebox.showwarning("警告", "请先进行计算")
            return
            
        self.plot_ax.clear()
        plot_type = self.plot_type.get()
        
        # 根据选择类型绘制图表
        if plot_type == "Lab":
            self.plot_lab_values()
        elif plot_type == "RGB":
            self.plot_rgb_values()
        elif plot_type == "HSV":
            self.plot_hsv_values()
        elif plot_type == "ΔE":
            self.plot_delta_e_values()
        elif plot_type == "反射光谱":
            self.plot_reflectance_spectra()
            
        self.plot_canvas.draw()

    def plot_reflectance_spectra(self):
        """绘制反射光谱曲线"""
        if not hasattr(self, 'spectra_data'):
            messagebox.showwarning("警告", "请先进行计算")
            return
            
        self.plot_ax.clear()
        
        # 绘制所有sheet的反射光谱
        for sheet, data in self.spectra_data.items():
            wavelengths = data['wavelengths']
            reflectance = data['reflectance']
            self.plot_ax.plot(wavelengths, reflectance, label=sheet)
            
        self.plot_ax.set_xlabel("Wavelength (nm)")
        self.plot_ax.set_ylabel("Reflectance")
        self.plot_ax.set_title("Reflection spectrum")
        self.plot_ax.legend()
        self.plot_ax.grid(True)
        
    def plot_lab_values(self):
        """绘制Lab值图表"""
        sheets = list(self.results.keys())
        L_values = [self.results[sheet]["CIELab"][0] for sheet in sheets]
        a_values = [self.results[sheet]["CIELab"][1] for sheet in sheets]
        b_values = [self.results[sheet]["CIELab"][2] for sheet in sheets]
        
        x = np.arange(len(sheets))
        width = 0.2
        
        self.plot_ax.bar(x - width, L_values, width, label='L')
        self.plot_ax.bar(x, a_values, width, label='a')
        self.plot_ax.bar(x + width, b_values, width, label='b')
        
        self.plot_ax.set_xticks(x)
        self.plot_ax.set_xticklabels(sheets)
        self.plot_ax.set_title("Lab Value")
        self.plot_ax.legend()
        
    def plot_rgb_values(self):
        """绘制RGB值图表"""
        sheets = list(self.results.keys())
        rgb_values = [self.results[sheet]["sRGB"] for sheet in sheets]
        
        # 将RGB值转换为0-255范围
        rgb_values = np.array(rgb_values) * 255
        
        x = np.arange(len(sheets))
        width = 0.2
        
        self.plot_ax.bar(x - width, rgb_values[:, 0], width, label='R')
        self.plot_ax.bar(x, rgb_values[:, 1], width, label='G')
        self.plot_ax.bar(x + width, rgb_values[:, 2], width, label='B')
        
        self.plot_ax.set_xticks(x)
        self.plot_ax.set_xticklabels(sheets)
        self.plot_ax.set_title("RGB Value")
        self.plot_ax.legend()
        
    def plot_hsv_values(self):
        """绘制HSV值图表"""
        sheets = list(self.results.keys())
        h_values = [self.results[sheet]["Hue"] for sheet in sheets]
        
        self.plot_ax.plot(sheets, h_values, marker='o')
        self.plot_ax.set_title("change of Hue")
        self.plot_ax.set_ylabel("Hue")
        
    def save_results(self):
        """保存结果"""
        if not hasattr(self, 'results'):
            messagebox.showwarning("警告", "请先进行计算")
            return
            
        output_dir = self.output_entry.get()
        if not output_dir:
            messagebox.showerror("错误", "请先选择输出目录")
            return
            
        try:
            os.makedirs(output_dir, exist_ok=True)
            
            # 保存颜色图片
            for sheet, data in self.results.items():
                plt.figure(figsize=(2, 2))
                color = np.array([[data["sRGB"]]])
                plt.imshow(color)
                plt.axis("off")
                plt.title(sheet)
                
                save_path = os.path.join(output_dir, f"{sheet}_color.png")
                plt.savefig(save_path, dpi=300, bbox_inches='tight')
                plt.close()
                
            # 保存文本结果
            text_path = os.path.join(output_dir, "results.txt")
            with open(text_path, 'w', encoding='utf-8') as f:
                for sheet, data in self.results.items():
                    f.write(f"\n--- {sheet} 计算结果 ---\n")
                    f.write(f"CIELab: {data['CIELab']}\n")
                    f.write(f"十六进制颜色: {data['Hex Color']}\n")
                    f.write(f"HSV Hue: {data['Hue']}°\n")
                    f.write(f"ΔE: {data['ΔE']:.2f}\n")
                    
            messagebox.showinfo("成功", f"结果已保存到：\n{output_dir}")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存结果时发生错误: {e}")

def rgb_to_hex(rgb):
    """将RGB颜色转换为十六进制"""
    return '#{:02x}{:02x}{:02x}'.format(int(rgb[0]*255),
                                        int(rgb[1]*255),
                                        int(rgb[2]*255))

class SliceWrapper:
    """用于为插值提供采样区间信息"""
    def __init__(self, start, stop, step):
        self.start = start
        self.stop = stop
        self.step = step
        self.end = stop
        self._interval = step

    @property
    def interval(self):
        return self._interval

    def __iter__(self):
        return iter(np.arange(self.start, self.stop, self.step))

if __name__ == "__main__":
    root = tk.Tk()
    app = ColorCalculatorApp(root)
    root.mainloop()
