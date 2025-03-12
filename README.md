该代码实现了一个自动化办公工具，能够将Excel表格数据按每页5行的形式复制到Word文档中，并在每页右下角添加电子印章图片，
模拟电子盖章效果。程序首先生成测试数据（包含20条员工信息的Excel文件和红色圆形印章图片），然后读取Excel数据并写入Word表格，
同时通过绝对定位将印章图片覆盖到每页的指定位置，最终生成一个规范的Word文档。整个过程自动化完成，适用于批量处理表格和电子盖章需求。
安装所有依赖包的命令如下：

```bash
pip install openpyxl python-docx Pillow faker
```

### 各依赖包的作用说明：
- **openpyxl**：用于读写Excel文件（.xlsx格式）  
- **python-docx**：用于操作Word文档（创建/修改.docx文件）  
- **Pillow**：用于图像处理（生成/操作印章图片）  
- **faker**：用于生成测试假数据（自动填充Excel表格）

### 附加说明：
1. 如果遇到权限问题，可以添加 `--user` 参数：
```bash
pip install --user openpyxl python-docx Pillow faker
```

2. 如果使用虚拟环境，建议先激活虚拟环境再安装

3. 在Linux/macOS系统若出现图像库依赖问题，可能需要先安装系统依赖：
```bash
# Ubuntu/Debian
sudo apt-get install python3-dev libjpeg-dev zlib1g-dev

# CentOS/RHEL
sudo yum install python3-devel libjpeg-turbo-devel zlib-devel
```
