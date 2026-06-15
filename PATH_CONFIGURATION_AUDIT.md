# 路径配置审计清单

审计目标：避免项目部署到 Windows Server 或 Linux/Docker 时仍依赖开发电脑路径。

当前 Windows Server 项目目录：

```text
C:\Users\Administrator\Downloads\esin\codex
```

当前 Windows Server Python Launcher 目录：

```text
C:\Users\Administrator\AppData\Local\Programs\Python\Launcher
```

## 已修改的运行路径

| 文件 | 原路径/原逻辑 | 修改后逻辑 |
|---|---|---|
| `app/config.py` | 固定使用项目内 `app/data`、`app/uploads`、`app/generated`、`app/doc_templates`、`templates/import`、`outputs`，不能由 `.env` 覆盖 | 增加 `.env` 读取；`DATA_DIR`、`UPLOAD_DIR`、`GENERATED_DIR`、`DOC_TEMPLATE_DIR`、`IMPORT_TEMPLATE_DIR`、`OUTPUTS_DIR`、`DB_PATH` 均可用环境变量覆盖 |
| `app/render_settings.py` | LibreOffice 默认只回退到 `D:\Program Files\program\soffice.com` | 优先读取 `SOFFICE_PATH`；其次查找系统 PATH；再自动检查常见 Windows LibreOffice 路径；最后回退到 `soffice` |
| `app/render_settings.py` | LibreOffice profile 固定基于代码里的 `outputs` | 使用 `OUTPUTS_DIR\.lo-profile-codex`，也可用 `LIBREOFFICE_PROFILE_DIR` 覆盖 |
| `start.ps1` | 固定使用开发机 Python：`C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe` | 从 `.env`/环境变量读取 `PYTHON_EXE` 或 `PYTHON_LAUNCHER_DIR`；支持服务器路径 `C:\Users\Administrator\AppData\Local\Programs\Python\Launcher\py.exe`；也会自动查找 `py` / `python` |
| `start.ps1` | 固定端口 `8088` | 支持 `.env` 中的 `APP_PORT`，也支持命令行参数覆盖 |
| `tools/website_control.ps1` | 固定使用开发机 Python：`C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe` | 与 `start.ps1` 使用同一套 Python 查找逻辑，并读取 `.env` |
| `tools/website_control.ps1` | PID 文件目录固定为 `app\data` | 若 `.env` 设置 `DATA_DIR`，后台控制脚本也使用该目录 |
| `tools/run_hidden.vbs` | 只支持 `python.exe app\server.py 8088` 固定参数形式 | 支持传入完整 Python 参数，兼容 `py.exe -3 app\server.py 8088` |
| `tools/start_website.ps1` | 独立维护启动逻辑 | 改为调用项目根目录 `start.ps1`，减少重复配置 |
| `tools/start_website.cmd` | 独立维护启动逻辑 | 改为调用项目根目录 `start.ps1` |
| `tools/start_website_visible.cmd` | 独立维护启动逻辑 | 改为调用项目根目录 `start.ps1` |
| `tools/start_website_hidden.vbs` | 依赖旧启动脚本路径逻辑 | 改为按自身位置定位项目根目录并调用 `start.ps1` |
| `tools/start_website_detached.py` | 固定开发机 Python：`C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe` | 优先用 `PYTHON_EXE`，其次用当前解释器，最后回退到 `py -3` |
| `tools/web_p1_background_generation_test.py` | 固定测试表格：`C:\Users\25896\Downloads\P1_APEX_MIND_已填写.xlsx` | 改为 `P1_TEST_XLSX` 环境变量；未设置时使用项目内测试 fixture |
| `app/template_audit.py` | 固定模板审计来源：`D:\RSIN GROUP Dropbox\...`，输出固定到 `C:\Users\25896\Documents\Codex\...\outputs` | 改为 `TEMPLATE_AUDIT_SOURCE_DIR` 和 `OUTPUTS_DIR`；未设置时使用项目内目录 |
| `app/template_matrix.py` | 输出目录固定到 `C:\Users\25896\Documents\Codex\...\outputs` | 改为 `OUTPUTS_DIR`；未设置时使用项目内 `outputs` |

## 新增文件

| 文件 | 用途 |
|---|---|
| `.env.windows.example` | Windows Server 环境变量示例。复制为 `.env` 后使用 |
| `WINDOWS_SERVER_DEPLOYMENT.md` | Windows Server 手动部署说明 |
| `PATH_CONFIGURATION_AUDIT.md` | 本路径审计清单 |

## 保留不改的路径

以下路径属于 Docker/Linux 容器内部路径或部署文档示例，不应该改成 Windows 路径：

| 路径 | 出现位置 | 原因 |
|---|---|---|
| `/app/app/data`、`/app/app/uploads`、`/app/app/generated` | `docker-compose.yml`、`scripts/*.sh`、`scripts/backup_data.ps1` | Docker 容器内部数据目录 |
| `/usr/bin/soffice` | `Dockerfile`、`docker-compose.yml` | Linux 容器内 LibreOffice 路径 |
| `/tmp/libreoffice-profile` | `Dockerfile`、`docker-compose.yml` | Linux 容器内 LibreOffice 临时 profile |
| `/opt/secretary-files` | `README_DEPLOY.md`、`DEPLOY_HANDOFF_FOR_COLLEAGUE.md` | 推荐 Linux 服务器部署目录示例 |

## 推荐部署选择

短期：当前 Windows Server 可以直接跑，按 `WINDOWS_SERVER_DEPLOYMENT.md` 操作。

长期：正式生产建议 Linux + Docker。这样 LibreOffice、字体、Poppler、Python 依赖都固定在镜像里，后续换服务器、GitHub 自动部署、备份恢复都更稳定。
