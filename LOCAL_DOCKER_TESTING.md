# 本机 Docker 测试说明

本项目可以在本机继续用 PowerShell 直接开发，也可以用 Docker 模拟 Linux 服务器环境。

## 当前建议

正式部署建议用 Linux + Docker。

本机开发建议：

1. 平时快速测试：直接运行 `start.ps1`
2. 上线前验证：用 Docker 跑一遍，确认 Linux 环境也能生成 PDF

## Windows 安装 Docker Desktop

本机需要：

- Windows 10/11 64 位
- WSL2
- Docker Desktop

可以手动下载安装：

```text
https://www.docker.com/products/docker-desktop/
```

也可以用 winget：

```powershell
winget install Docker.DockerDesktop
```

安装 Docker Desktop 通常需要管理员权限，可能需要重启电脑。重启后打开 Docker Desktop，等左下角显示 Docker 已启动，再回到项目目录测试。

## 检查 Docker 是否可用

先运行项目自带检查脚本：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\check_local_environment.ps1
```

也可以手动检查：

```powershell
docker --version
docker compose version
```

如果提示找不到 `docker`，说明 Docker Desktop 没装好，或者没有重启/没有打开 Docker Desktop。

## 本机 Docker 启动项目

```powershell
cd C:\Users\25896\Documents\Codex\2026-05-25\new-chat
Copy-Item .env.example .env
docker compose up -d --build
```

打开：

```text
http://127.0.0.1:8088
```

查看运行状态：

```powershell
docker compose ps
docker compose logs -f
```

停止：

```powershell
docker compose down
```

## 直接本机启动

不用 Docker 时：

```powershell
cd C:\Users\25896\Documents\Codex\2026-05-25\new-chat
powershell -NoProfile -ExecutionPolicy Bypass -File .\start.ps1
```

## 中文乱码规避

项目已配置：

- Docker 环境变量：`LANG=C.UTF-8`、`LC_ALL=C.UTF-8`、`PYTHONUTF8=1`
- PowerShell 启动脚本：自动设置 UTF-8 输出和 Python UTF-8
- Git/编辑器规则：`.editorconfig`、`.gitattributes`

如果手动排查时 PowerShell 仍显示乱码，先运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\set_utf8_console.ps1
```

然后再运行测试命令。
