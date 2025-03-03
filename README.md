# 颜色计算器

这是一个基于Python的颜色计算器程序，可以将lumerical FDTD光谱数据转换为颜色值。（反射スペクトルを色へ変換する）

## 功能
- 读取Excel表格中的光谱数据
- 计算CIELab色彩空间的颜色值（L，a，b）、sRGB颜色值和HSV色彩空间的Hue值
- 显示颜色预览
- 保存计算结果

## 使用方法
1. 安装依赖：`pip install -r requirements.txt`
2. 运行程序：`python color_calculator_gui.py`
3. 选择包含光谱数据的Excel文件
4. 点击"开始计算"查看结果
5. 点击"保存结果"保存计算结果

## Excel文件格式要求
Excel文件应包含以下列：
- Wavelength: 波长(nm)
- Reflectance: 反射率(0-1)

Excel文件应包含以下工作表：（分别表示对应的周围环境曲折率，模拟从1到1.4的曲折率变化）
- 1.0
- 1.1
- 1.2
- 1.3
- 1.4

程序文件路径：
- 主程序：code/color_calculator_gui.py
- 依赖文件：requirements.txt
