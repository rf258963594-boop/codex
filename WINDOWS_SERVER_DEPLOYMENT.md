# Windows Server 部署说明

本文件用于当前 Windows Server 路径：

```text
项目根目录：
C:\Users\Administrator\Downloads\esin\codex

Python Launcher 目录：
C:\Users\Administrator\AppData\Local\Programs\Python\Launcher
```

## 先说结论

推荐的长期生产方案仍然是 **Linux + Docker**。原因是 PDF 生成依赖 LibreOffice、字体、Poppler，Docker 可以把这些环境固定住，换服务器也不容易出错。

如果当前服务器已经是 Windows Server，也可以直接运行。本次代码已改成相对路径和环境变量配置，不再依赖开发电脑路径。

## Windows Server 第一次启动

1. 进入项目目录：

```powershell
cd C:\Users\Administrator\Downloads\esin\codex
```

2. 复制环境配置：

```powershell
Copy-Item .env.windows.example .env
```

3. 编辑 `.env`，至少修改：

```text
DEFAULT_ADMIN_PASSWORD=换成强密码
```

如果 LibreOffice 自动识别失败，在 `.env` 中指定真实路径：

```text
SOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe
```

4. 安装 Python 依赖：

```powershell
& "C:\Users\Administrator\AppData\Local\Programs\Python\Launcher\py.exe" -3 -m pip install -r requirements.txt
```

5. 启动网站：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\start.ps1
```

6. 打开网站：

```text
http://服务器IP:8088
```

如果在服务器本机测试：

```text
http://127.0.0.1:8088
```

## 需要安装的软件

- Python 3.11 或 3.12
- LibreOffice
- 项目依赖：`requirements.txt`

如果只在 Windows Server 直接运行，不强制需要 Docker。

## 路径配置原则

程序默认使用项目内相对路径：

```text
app\data
app\uploads
app\generated
app\doc_templates
templates\import
outputs
```

只有需要把数据放到项目外部时，才在 `.env` 中覆盖：

```text
DATA_DIR=...
UPLOAD_DIR=...
GENERATED_DIR=...
OUTPUTS_DIR=...
DOC_TEMPLATE_DIR=...
IMPORT_TEMPLATE_DIR=...
DB_PATH=...
```

不要把个人电脑路径写进代码，例如：

```text
C:\Users\25896\...
D:\RSIN GROUP Dropbox\...
```

## Linux / Docker 部署说明

如果服务器改回 Linux，建议直接用 Docker：

```bash
cd /opt/secretary-files
cp .env.example .env
nano .env
bash scripts/start_docker.sh
```

Docker 内部路径例如 `/app/app/data`、`/tmp/libreoffice-profile`、`/usr/bin/soffice` 是容器内部路径，不需要也不应该改成 Windows 路径。

## 常见报错定位

### 找不到 Python

确认文件存在：

```powershell
Test-Path "C:\Users\Administrator\AppData\Local\Programs\Python\Launcher\py.exe"
```

如果不存在，安装 Python，或者在 `.env` 设置：

```text
PYTHON_EXE=C:\实际路径\python.exe
```

### PDF 生成失败

先确认 LibreOffice 路径：

```powershell
Test-Path "C:\Program Files\LibreOffice\program\soffice.exe"
```

如果返回 False，安装 LibreOffice 或修改 `.env`：

```text
SOFFICE_PATH=C:\实际路径\soffice.exe
```

### 页面无法访问

检查 8088 端口是否被防火墙或安全组拦截。

服务器本机能访问但外部不能访问，通常是云服务器安全组没有放行 `8088`。
