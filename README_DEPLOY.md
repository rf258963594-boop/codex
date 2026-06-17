# 新加坡秘书文件生成器部署说明

> 当前如果部署在 Windows Server，请先阅读 `WINDOWS_SERVER_DEPLOYMENT.md`。
> 本文下面主要是推荐的 Linux + Docker 生产部署方案。

本项目用于内部批量生成新加坡公司注册、维护、变更、年审签字 PDF 文件。推荐部署方式是：GitHub 保存代码和模板，阿里云服务器运行网站并保存真实业务数据。

## 推荐架构

GitHub 保存：

- 网站代码
- Word 文件模板
- Excel 空白/示例导入模板
- Dockerfile
- 部署脚本
- 部署说明

服务器保存：

- `.env` 密码配置
- SQLite 数据库
- 常用人员资料
- 自动签名图片
- 客户上传的 Excel
- 生成的 PDF/ZIP
- 备份文件

不要把客户资料、真实签名、数据库、`.env` 上传到 GitHub。

## 一次性准备服务器

推荐阿里云 ECS：

- Ubuntu 22.04 或 24.04
- 2 核 4G 起步，文件生成量大时建议 4 核 8G
- 系统盘 40G 起步
- 内部使用阶段建议安全组只允许公司固定 IP 访问

服务器需要安装：

```bash
sudo apt update
sudo apt install -y git ca-certificates curl
```

安装 Docker 可参考 Docker 官方文档。安装好后确认：

```bash
docker --version
docker compose version
```

## 第一次从 GitHub 部署

在服务器选择一个目录，例如 `/opt/secretary-files`：

```bash
sudo mkdir -p /opt/secretary-files
sudo chown -R $USER:$USER /opt/secretary-files
git clone https://github.com/rf258963594-boop/codex.git /opt/secretary-files
cd /opt/secretary-files
```

如果正式代码还在开发分支，可以临时切到开发分支：

```bash
git checkout codex/secretary-generator-stage-updates
```

正式上线后建议只用 `main` 分支。

复制环境配置：

```bash
cp .env.example .env
nano .env
```

至少修改：

```text
APP_PORT=8088
DEFAULT_ADMIN_USERNAME=admin
DEFAULT_ADMIN_PASSWORD=换成强密码
```

启动网站：

```bash
bash scripts/start_docker.sh
```

打开：

```text
http://服务器IP:8088
```

首次登录账号密码来自 `.env`。注意：`.env` 里的默认账号密码只在第一次创建数据库时生效。如果数据库已经创建，后续请在后台用户管理里改密码。

如果忘记管理员密码，或第一次部署时 `.env` 使用了临时密码，可以在服务器项目目录执行：

```bash
docker compose exec secretary-files python app/reset_admin_password.py --username admin --password admin123
```

然后使用 `admin / admin123` 登录。登录后建议立即在后台改成强密码。

## 后续手动更新

如果暂时不启用 GitHub 自动部署，服务器上手动更新：

```bash
cd /opt/secretary-files
bash scripts/deploy_server.sh
```

这个脚本会：

- 检查 `.env`
- 如果网站正在运行，先备份数据库、上传文件、生成文件
- 从 GitHub 拉取指定分支
- 重新构建并启动 Docker
- 做一次健康检查

默认部署 `main` 分支。要部署其他分支：

```bash
DEPLOY_BRANCH=codex/secretary-generator-stage-updates bash scripts/deploy_server.sh
```

## GitHub 自动部署

项目已包含：

```text
.github/workflows/deploy.yml
scripts/deploy_server.sh
```

自动部署逻辑：

```text
推送 main 分支 -> GitHub Actions 登录阿里云服务器 -> 执行 scripts/deploy_server.sh -> 网站自动更新
```

### 服务器需要先准备 SSH Key

在你本机或服务器生成一组部署专用 SSH key：

```bash
ssh-keygen -t ed25519 -C "github-deploy-secretary-files"
```

把公钥内容加入服务器：

```bash
cat ~/.ssh/id_ed25519.pub >> ~/.ssh/authorized_keys
chmod 600 ~/.ssh/authorized_keys
```

把私钥内容保存到 GitHub Secrets。

### GitHub Secrets 需要填写

进入 GitHub 仓库：

```text
Settings -> Secrets and variables -> Actions -> New repository secret
```

新增：

```text
ALIYUN_HOST       服务器 IP 或域名
ALIYUN_PORT       SSH 端口，通常是 22
ALIYUN_USER       SSH 用户，例如 root 或 ubuntu
ALIYUN_SSH_KEY    SSH 私钥完整内容
DEPLOY_PATH       服务器项目目录，例如 /opt/secretary-files
```

设置好以后，只要 `main` 分支有新提交，GitHub 会自动部署。

也可以在 GitHub 页面手动触发：

```text
Actions -> Deploy website -> Run workflow
```

## 分支建议

推荐：

```text
main                         正式生产分支，自动部署只看它
codex/xxx                    开发测试分支
```

开发流程：

```text
Codex 修改开发分支 -> 本地/测试确认 -> 合并到 main -> 自动部署
```

这样不会把开发到一半的内容直接更新到正式网站。

## 数据保存位置

Docker Compose 会创建三个持久化卷：

```text
secretary_data       数据库、常用人员、签名图片、模板版本记录
secretary_uploads    上传过的 Excel
secretary_generated  生成的 PDF/ZIP 文件
```

不要随便删除 Docker volume。删除 volume 等于删除真实业务数据。

## 备份

Linux 服务器：

```bash
bash scripts/backup_data.sh
```

Windows 本机 Docker：

```powershell
.\scripts\backup_data.ps1
```

自动部署脚本 `scripts/deploy_server.sh` 每次更新前也会自动生成一次备份。

建议把 `backups/` 定期复制到服务器外部，例如阿里云 OSS、本地硬盘或公司网盘。

## Word 模板和 Excel 模板

正式 Word 模板：

```text
app/doc_templates/
```

正式 Excel 导入模板：

```text
templates/import/
```

目前后台已经支持 Word 文件模板下载、上传草稿、启用和回滚。Excel 导入模板目前主要是下载入口；如果要后台上传并启用 Excel 模板，需要后续再加版本库功能。

## Docker 中包含的渲染环境

Docker 镜像内已安装：

- Python 3.12
- LibreOffice Writer，用于 Word 转 PDF
- Poppler，用于 PDF 检查
- Noto CJK / Noto / Liberation 字体，用于中文和英文排版

云端不会依赖本地打印机。

## Nginx / 域名 / HTTPS

内部测试可以先用：

```text
http://服务器IP:8088
```

正式使用建议加域名和 HTTPS。Nginx 示例：

```nginx
server {
    listen 80;
    server_name your-domain.com;

    client_max_body_size 30m;

    location / {
        proxy_pass http://127.0.0.1:8088;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

## 安全注意事项

- 第一次上线前必须修改 `.env` 里的管理员密码。
- 不要把 `.env`、数据库、客户文件、生成 PDF 上传 GitHub。
- 阿里云安全组不要全网开放后台，内部阶段建议限制公司 IP。
- GitHub Actions 的 SSH key 建议只用于部署，不要复用个人主密钥。
- 定期备份 `secretary_data`、`secretary_uploads`、`secretary_generated`。
- 上传文件包含客户资料，服务器磁盘和备份位置都要按敏感资料处理。

## 上线前检查清单

- `.env` 已设置强密码
- Docker 和 Docker Compose 可用
- `docker compose up -d --build` 能启动
- 登录成功
- P1 测试表能生成 PDF
- P2 v7 测试表能生成 M01-M05 对应 PDF
- 常用人员和签名图片能保存
- 生成文件能下载
- `bash scripts/backup_data.sh` 能生成备份包
- GitHub Secrets 已配置
- GitHub Actions 手动触发部署成功
- 阿里云安全组已限制访问范围
- 如有域名，HTTPS 已配置
