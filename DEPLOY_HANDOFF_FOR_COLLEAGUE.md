# 新加坡秘书文件生成器上线交接文档

这份文档给负责上线部署的同事使用。项目已经放在 GitHub，推荐从 GitHub 直接拉取代码到阿里云服务器部署。

如果当前服务器是 Windows Server，请先看 `WINDOWS_SERVER_DEPLOYMENT.md`；本文第 4 节之后主要按 Linux + Docker 服务器写。

## 1. 项目基本信息

GitHub 仓库：

```text
https://github.com/rf258963594-boop/codex
```

正式部署分支：

```text
main
```

网站类型：

```text
内部使用的文件生成网站
```

主要功能：

- 上传 Excel 资料表
- 自动生成 P1 注册文件 PDF 包
- 自动生成 P2 M01-M05 维护、变更、年审 PDF 包
- 后台管理用户、常用人员、签名、Word 模板版本

## 2. 重要安全提醒

不建议直接共享 GitHub 账号密码。

更推荐：

- 把同事加入 GitHub 仓库协作者
- 或者让同事用部署专用 SSH key / deploy key
- GitHub 密码、服务器密码、`.env` 文件不要发到群里

不要上传到 GitHub 的内容：

- 客户上传的 Excel
- 生成的 PDF / ZIP
- 网站数据库
- 真实签名图片
- `.env` 密码配置文件
- 服务器备份文件

这些真实业务数据只应保存在服务器 Docker volume 和单独备份位置。

## 3. 推荐服务器

推荐阿里云 ECS：

```text
系统：Ubuntu 22.04 或 Ubuntu 24.04
配置：2核4G 起步，文件生成多时建议 4核8G
系统盘：40G 起步
```

安全组建议：

```text
22    SSH 登录
8088  网站测试访问
80    域名 HTTP，可后期开
443   HTTPS，可后期开
```

内部使用阶段，建议安全组只允许公司固定 IP 访问 `8088`。

## 4. 服务器安装 Docker 和 Git

SSH 登录服务器：

```bash
ssh root@服务器IP
```

安装基础环境：

```bash
apt update
apt install -y git docker.io docker-compose-plugin
systemctl enable --now docker
```

检查安装结果：

```bash
docker --version
docker compose version
git --version
```

## 5. 第一次从 GitHub 部署

在服务器创建项目目录：

```bash
mkdir -p /opt/secretary-files
git clone https://github.com/rf258963594-boop/codex.git /opt/secretary-files
cd /opt/secretary-files
git checkout main
```

复制环境配置：

```bash
cp .env.example .env
nano .env
```

修改为实际密码，例如：

```text
APP_PORT=8088
DEFAULT_ADMIN_USERNAME=admin
DEFAULT_ADMIN_PASSWORD=请换成强密码
```

保存 `nano`：

```text
Ctrl + O
回车
Ctrl + X
```

启动网站：

```bash
docker compose up -d --build
```

查看状态：

```bash
docker compose ps
```

打开网站：

```text
http://服务器IP:8088
```

使用 `.env` 中设置的账号密码登录。

注意：`.env` 中的默认管理员账号密码只在第一次创建数据库时生效。如果数据库已经创建，后续需要在网站后台修改密码。

## 6. 手动更新网站

以后如果 GitHub 的 `main` 分支有更新，在服务器执行：

```bash
cd /opt/secretary-files
bash scripts/deploy_server.sh
```

这个脚本会自动做：

- 检查项目目录
- 检查 `.env`
- 如果网站正在运行，先备份运行数据
- 拉取 GitHub 最新 `main`
- 重新构建 Docker
- 重启网站
- 做健康检查

如果要临时部署其他分支：

```bash
DEPLOY_BRANCH=codex/secretary-generator-stage-updates bash scripts/deploy_server.sh
```

正式环境建议只部署 `main`。

## 7. 自动部署 GitHub Actions

项目已经包含自动部署配置：

```text
.github/workflows/deploy.yml
scripts/deploy_server.sh
```

自动部署逻辑：

```text
推送 main 分支 -> GitHub Actions 登录服务器 -> 执行 deploy_server.sh -> 网站自动更新
```

### 7.1 生成部署 SSH Key

在本机或服务器生成部署专用 key：

```bash
ssh-keygen -t ed25519 -C "github-deploy-secretary-files"
```

把公钥加入服务器：

```bash
cat ~/.ssh/id_ed25519.pub >> ~/.ssh/authorized_keys
chmod 600 ~/.ssh/authorized_keys
```

把私钥内容放入 GitHub Secrets。

### 7.2 GitHub Secrets 设置

进入 GitHub 仓库：

```text
Settings -> Secrets and variables -> Actions -> New repository secret
```

需要添加：

```text
ALIYUN_HOST       服务器 IP 或域名
ALIYUN_PORT       SSH 端口，通常是 22
ALIYUN_USER       SSH 用户，例如 root 或 ubuntu
ALIYUN_SSH_KEY    SSH 私钥完整内容
DEPLOY_PATH       服务器项目目录，例如 /opt/secretary-files
```

设置完成后，只要 `main` 分支有新提交，GitHub 会自动部署。

也可以手动触发：

```text
GitHub -> Actions -> Deploy website -> Run workflow
```

## 8. 数据和备份

Docker Compose 会创建三个持久化 volume：

```text
secretary_data       数据库、常用人员、签名图片、模板版本记录
secretary_uploads    上传过的 Excel
secretary_generated  生成的 PDF / ZIP
```

不要删除这些 volume。删除就等于删除业务数据。

手动备份：

```bash
cd /opt/secretary-files
bash scripts/backup_data.sh
```

备份默认放在：

```text
/opt/secretary-files/backups
```

建议定期复制到服务器外部，例如阿里云 OSS、公司网盘或本地硬盘。

## 9. 域名和 HTTPS

测试阶段可直接使用：

```text
http://服务器IP:8088
```

正式使用建议配置域名和 HTTPS。

Nginx 示例：

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

## 10. 常见检查命令

查看容器：

```bash
docker compose ps
```

查看日志：

```bash
docker compose logs --tail=100 secretary-files
```

重启网站：

```bash
docker compose restart
```

停止网站：

```bash
docker compose down
```

重新构建：

```bash
docker compose up -d --build
```

## 11. 上线验收清单

上线后请检查：

- 可以打开登录页
- 管理员可以登录
- 可以下载 P1 / P2 Excel 模板
- 上传 P1 测试表可以生成 PDF
- 上传 P2 v7 测试表可以生成对应 PDF
- 常用人员可以新增、编辑、删除
- 自动签名图片可以生成
- 生成文件可以下载
- 备份脚本可以生成备份包
- GitHub Actions 可以成功部署
- 阿里云安全组没有无必要地开放全网访问

## 12. 后续建议

第一阶段先保证内部可用。

后续可逐步升级：

- 后台上传 Excel 导入模板版本
- 文件生成任务队列
- PostgreSQL 数据库
- 客户后台
- 在线填写资料
- 在线电子签名
- 更细的文件权限和操作审计
