# 群晖部署手册 · client-base 网页端

> 目标：在群晖上跑一个 `http://<群晖IP>:5174` 就能用的网页版。
> 全程**不需要 npm install、不需要编译**——镜像用官方 `node:20-slim`，代码只读挂进去。

---

## 0. 先说清能力边界

| 能力 | 群晖网页版 |
|---|---|
| 画布、连线、节点增删 | ✅ |
| 表格节点、参数、公式、DAG 重算 | ✅ |
| `.wfc` 导入 / 导出 | ✅ |
| 操作审计日志 | ✅ 落在 `data/logs/` |
| **OnlyOffice 文档/工作簿在线编辑** | ❌ 打不开时显示「浏览器版不支持在线编辑」提示卡 |
| 多设备访问 | ✅ 局域网任意设备、外网走反向代理 |

OnlyOffice 的文档服务（DS）在本项目里是 **Windows 二进制**，群晖跑不了。
想恢复在线编辑有两条路，都不在今天这个包里：
① 在群晖 Docker 里另起官方 `onlyoffice/documentserver` 镜像（x86 机型可以），再把 app-server 的 DS 地址指过去；
② 继续用 Windows 那台机器跑 DS。

---

## 1. 前置检查（1 分钟）

1. **套件中心 → 安装 Container Manager**（DSM 7.2；老版本叫 Docker）。
2. 确认共享文件夹 `docker` 存在（File Station 左侧能看到）。没有就新建一个。
   若你的共享文件夹叫别的名字，见第 5 节「改根目录」。
3. 确认 `5174` 端口没被占用：
   控制面板 → 登录门户 → 高级 → 确认没冲突；或用 SSH `netstat -tlnp | grep 5174`。

---

## 2. 三种装法，选一种

### 路线 A · SSH 一键（推荐，最省事）

`admin` 账号要有 SSH 权限（控制面板 → 终端机与 SNMP → 启用 SSH）。

```bash
# ① 传包（在你自己的电脑上执行；把 IP 换成群晖的）
scp client-base-synology-v0.00.15.tar.gz admin@192.168.1.x:/tmp/

# ② 登录群晖
ssh admin@192.168.1.x

# ③ 解压 + 装（sudo 会要密码，就是 admin 的密码）
sudo mkdir -p /volume1/docker/client-base
sudo tar -xzf /tmp/client-base-synology-v0.00.15.tar.gz \
     --strip-components=1 -C /volume1/docker/client-base
cd /volume1/docker/client-base
sudo sh install.sh
```

看到 `访问: http://192.168.1.x:5174/` 就好了。

### 路线 B · 纯图形界面（不用 SSH）

1. **File Station** 建目录：`docker/client-base/`，里面再建
   `app/`、`data/pages/`、`data/office/`、`data/logs/`。
2. 把 tar.gz 上传到 `docker/client-base/app/` 并**右键解压**。
   解压出来的东西如果多套了一层文件夹，进到那一层，把
   `app-server.mjs`、`dist-web/`、`scripts/`、`package.json` 全选剪切到 `app/` 根下。
   > 判据：`docker/client-base/app/app-server.mjs` 这个路径存在就对了。
3. **Container Manager → 项目 → 新增**：
   - 项目名称：`client-base`
   - 路径：`/volume1/docker/client-base`
   - 来源：选「**创建 docker-compose.yml**」，把本包里的 `docker-compose.yml` 内容整段粘进去
     （或者选「上传」直接选文件）
   - 下一步 → 完成。它会自动拉镜像并启动。

### 路线 C · 做成单体镜像（要带走/推到私有仓库时用）

```bash
cd /volume1/docker/client-base/app
sudo docker build -f Dockerfile.synology -t client-base-web:0.00.15 .
```
之后用 `docker run` 或改 compose 的 `image:` 指向它，去掉 `/app:ro` 挂载。

---

## 3. 验证（一定要看这三项）

```bash
curl -s http://127.0.0.1:5174/healthz
# 期望：{"ok":true,"version":"0.00.15","channel":"web","oo":false,...}
```

- 浏览器打开 `http://<群晖IP>:5174/` → 看到画布，**标题栏右上角有 `v0.00.15 · web`**。
- 点一个 Office 节点 → 应显示「浏览器版不支持在线编辑」提示卡，而不是一直转圈。
- 随便点两下画布，然后看 `docker/client-base/data/logs/audit-<今天>.log` 有没有新行。

```bash
tail -5 /volume1/docker/client-base/data/logs/audit-$(date +%F).log
```

---

## 4. 数据在哪、怎么备份

| 路径 | 内容 | 要不要备份 |
|---|---|---|
| `data/pages/` | 所有画布页（JSON）+ `index.json` | **必须** |
| `data/office/` | 上传的 Office 中转文件 | 建议 |
| `data/logs/` | 审计日志，按天，超 8MB 滚一代 | 可选 |

在 **Hyper Backup** 里加一个任务，只备份 `/volume1/docker/client-base/data` 即可，几 MB。
或者在网页端逐个导出 `.wfc`（画布右上角保存按钮），这是最抗灾的形态。

---

## 5. 改根目录 / 改端口

**改根目录**：进 `docker/client-base/`，把 `.env.example` 复制成 `.env`，改一行：

```
CB_ROOT=/volume2/我的项目/client-base
```

然后在 Container Manager 里重启项目。

**改端口**：编辑 `docker-compose.yml` 的 `ports`，如 `"8074:5174"`（左边随你，右边别动）。

---

## 6. 外网访问（可选）

最省事：**控制面板 → 登录门户 → 高级 → 反向代理 → 新增**

- 来源：HTTPS / 你的域名或 QuickConnect 域名 / 端口 443（或 8443）
- 目的地：HTTP / `localhost` / `5174`

配好后用 `https://xxx.synology.me` 访问，证书在「安全性 → 证书」里申请 Let's Encrypt。

> 别直接把 5174 裸映射到公网——**先看第 6.1 节**。0.00.16 起服务自带访问口令，但口令走明文 HTTP，
> 公网长期用仍建议套 DSM 反向代理 + HTTPS。

### 6.1 访问口令（0.00.16 起，默认开启）

打开网页会先见到仿 DSM 风格的登录页，输 **访问口令** 才能进画布。

| 事项 | 说明 |
|---|---|
| 口令在哪设 | `docker-compose.yml` 的 `APP_PASSWORD`，默认 `${CB_PASSWORD:-xiao2ming}`；改 `.env` 里的 `CB_PASSWORD` 后 `sudo docker compose up -d` 生效 |
| 会话 | 勾「记住这台设备」= 30 天免输；不勾 = 关浏览器即失效。标题栏右侧「退出」可主动登出 |
| 防爆破 | 同 IP 10 分钟内错 5 次 → 锁 10 分钟（429） |
| 登录记录 | 成功/失败/封禁都写进 `data/logs/audit-*.log`，搜 `[srv] 鉴权` |
| 关掉验证 | compose 里把 `APP_PASSWORD` 设为空字符串并重启容器（**仅限纯内网调试**） |
| 注意 | 口令经 HTTP 明文传输；公网长期使用请在 DSM 反代上套 HTTPS |

---

## 7. 升级到新版本

新版部署包出来之后：

```bash
# SSH 路线
sudo tar -xzf client-base-synology-v0.00.16.tar.gz \
     --strip-components=1 -C /volume1/docker/client-base
cd /volume1/docker/client-base && sudo sh install.sh
```

`install.sh` 不会动 `data/`，画布数据原地保留。
图形界面路线：File Station 里把 `app/` 里的旧文件覆盖掉 → Container Manager 里重启项目。

---

## 8. 排障

| 现象 | 原因 / 处理 |
|---|---|
| `pull` 卡住或超时 | Docker Hub 被墙。SSH 里执行：`sudo mkdir -p /etc/docker` 写 `daemon.json` 加镜像加速（`https://docker.m.daocloud.io`），再 `sudo synosystemctl restart pkgctl-ContainerManager` |
| 容器起来又退出 | `docker logs client-base`。多半是 `/app/app-server.mjs` 路径不对（app/ 多套了一层目录）或端口被占 |
| 页面 404 / 白屏 | `dist-web/index.html` 没在 `app/` 根下；或改过 `APP_STATIC_DIR` |
| 能开页面但画布空白 | `data/pages/` 没有 `index.json`。把 `app/server-data/pages/` 里的默认数据拷进 `data/pages/` |
| 保存提示失败 | `data/pages` 权限不对：`sudo chown -R 1026:100 /volume1/docker/client-base/data`（1026 是 DSM 默认 users 组） |
| 日志时间是 UTC（差 8 小时） | 镜像缺 tzdata。本项目已用 `node:20-slim` 规避；若你换了镜像，加 `TZ=Asia/Shanghai` 并确保镜像内有 tzdata |
| 审计日志没生成 | 工作目录必须是可写卷：compose 用 `working_dir: /data`（审计日志写在 `cwd/logs`）。**不要**试图把日志挂到 `/app/logs`——`/app` 是只读挂载，runc 无法在其下创建挂载点（报 `read-only file system`），这是实测踩过的坑 |
| 一直跳登录页 / cookie 无效 | 口令改过（改口令 = 全部旧登录失效，重新输一次即可）；或浏览器禁了 cookie |
| 忘了口令 | SSH 看 `/volume1/docker/client-base/.env` 的 `CB_PASSWORD`，或直接改成新值再 `docker compose up -d` |
| `docker: command not found`（任务计划/裸 sudo） | DSM 的受限 shell PATH 很短。`install.sh` 已把 `/var/packages/ContainerManager/target/usr/bin` 等补进 PATH；自己写脚本时记得同样处理，或直接用绝对路径 |
| SFTP 连上就断（`Channel closed`） | DSM 默认不给普通用户开 SFTP 子系统。自动部署脚本 `scripts/_deploy_synology.py` 因此改用 SSH stdin 管道传文件 |
| 拉镜像失败/超时 | compose 会先试官方源，失败自动换 `docker.m.daocloud.io`；也可以提前 `docker pull node:20-slim` |

---

## 9. 端口与容器速查

```bash
sudo docker ps --filter name=client-base
sudo docker logs -f --tail 100 client-base     # 实时看服务端输出
sudo docker restart client-base                # 改完 .env / compose 后
cd /volume1/docker/client-base && sudo docker compose down   # 彻底停掉
```
