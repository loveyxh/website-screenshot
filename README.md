# 网站自动截图及报告文档生成工具

## 项目简介

网站自动截图及报告文档生成工具是一个自动化工具，用于批量获取网站截图并生成报告文档。该系统可以读取Excel文件中的网站信息（包括序号、网站名称和网站域名），自动访问这些网站并进行截图，最后将所有信息整合到Word文档中，形成一份完整的网站截图报告。

### 主要功能

- **Excel数据读取**：自动读取指定Excel文件中的网站信息
- **网站访问与截图**：自动访问网站并进行截图，支持HTTP/HTTPS协议自动切换
- **异常处理**：对无法访问的网站生成404错误图片
- **并发处理**：使用多线程技术提高批量处理效率
- **Word文档生成**：将网站信息和截图整合到Word文档中
- **分页处理**：支持大批量数据的分页文档生成

## 项目部署说明

### 环境要求

- Python 3.8或更高版本
- Chrome浏览器（用于网站截图）

### 安装步骤

1. 克隆仓库到本地

```bash
git clone https://github.com/loveyxh/website-screenshot.git
cd website-screenshot
```

2. 安装依赖包

```bash
pip install -r requirements.txt
```

依赖包包括：
- pandas>=2.0.0：用于Excel数据处理
- selenium>=4.0.0：用于网站访问和截图
- python-docx>=1.0.0：用于Word文档生成
- openpyxl>=3.0.0：用于Excel文件读取
- webdriver-manager>=4.0.0：用于Chrome驱动管理


## 项目结构

```
├── main.py              # 主程序
├── requirements.txt     # 依赖包列表
├── list.xlsx           # 网站数据文件
├── screenshots/        # 截图保存目录
├── website_screenshot.log  # 运行日志
└── 网站截图报告(0-100).docx  # 生成的报告文档
```

### 自定义配置

如需修改程序配置，可编辑`main.py`文件：

- 修改并发线程数：调整`ThreadPoolExecutor`的`max_workers`参数
- 修改截图分辨率：调整`set_window_size`的参数
- 修改每页记录数：调整`page_size`变量

## 项目运行说明

### 准备数据

1. 准备一个名为`list.xlsx`的Excel文件，放在项目根目录下
2. Excel文件需包含名为`sheet1`的工作表
3. 工作表需包含以下列：
   - 序号
   - 网站名称
   - 网站域名

### 运行程序

```bash
python main.py
```

### 运行过程

1. 程序启动后，会自动读取`list.xlsx`文件中的网站信息
2. 对每个网站进行访问和截图，截图保存在`screenshots`目录下
3. 生成Word文档，每100个网站生成一个文档，文档命名格式为`网站截图报告(起始序号-结束序号).docx`
4. 程序运行日志保存在`website_screenshot.log`文件中

### 注意事项

- 程序会自动处理网站协议（HTTP/HTTPS），无需在Excel中指定
- 对于无法访问的网站，会生成404错误图片
- 程序使用多线程处理，默认并发数为5，可在代码中调整
- 截图分辨率固定为800x600像素
- 文档中每个网站记录包含序号、网站名称、网站域名和截图



