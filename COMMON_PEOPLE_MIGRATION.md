# 常用人员和签名迁移说明

用途：把本机开发环境中已经维护好的常用人员、别名、角色、证件信息、联系方式、自动签名图片迁移到服务器。

迁移包只包含 `common_people` 和 `app/data/signatures` 中被引用的签名图片，不包含用户账号、客户上传文件、生成记录、数据库其他业务数据。

## 1. 在本机导出

在本机项目目录执行：

```powershell
cd C:\Users\25896\Documents\Codex\2026-05-25\new-chat
py -3 app\export_common_people.py
```

命令会输出一个 zip 路径，例如：

```text
outputs\common_people_migration_20260617-120000.zip
```

把这个 zip 上传到服务器项目目录，例如：

```text
/opt/secretary-files/common_people_migration.zip
```

## 2. 在服务器导入

服务器进入项目目录：

```bash
cd /opt/secretary-files
git pull origin main
docker compose up -d --build
docker compose cp common_people_migration.zip secretary-files:/tmp/common_people_migration.zip
docker compose exec secretary-files python app/import_common_people.py --input /tmp/common_people_migration.zip
```

导入完成后，刷新后台“常用人员”页面检查。

## 3. 导入规则

- 按 `display_name` 匹配。
- 服务器已有同名人员：更新资料和签名。
- 服务器没有同名人员：新增。
- 不会删除服务器已有人员。
- 不会覆盖用户账号。
- 不会迁移测试生成记录。

如果确实要让服务器常用人员完全按迁移包为准，可以追加：

```bash
docker compose exec secretary-files python app/import_common_people.py --input /tmp/common_people_migration.zip --deactivate-missing
```

一般不建议首次导入就用 `--deactivate-missing`。
