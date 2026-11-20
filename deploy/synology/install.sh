#!/bin/sh
# client-base 网页端 · 群晖一键安装/更新
#
# 用法（SSH 登录群晖后）：
#   sudo sh install.sh                 # 默认装到 /volume1/docker/client-base
#   sudo sh install.sh /volume2/xxx    # 指定别的目录（共享文件夹必须已存在）
#
# 反复跑是安全的：已有画布数据不会被覆盖，容器会被重建。

set -e

ROOT=${1:-/volume1/docker/client-base}
HERE=$(cd "$(dirname "$0")" && pwd)

# DSM 的「任务计划」与裸 sudo shell 里 PATH 很短，docker 常常找不到 —— 显式补上。
# 注意别用 `[ -d x ] && ...` 的形式：set -e 下条件为假会直接退出脚本。
for d in /var/packages/ContainerManager/target/usr/bin /usr/local/bin /usr/bin /bin /usr/sbin /sbin; do
  if [ -d "$d" ]; then
    case ":$PATH:" in
      *":$d:"*) ;;
      *) PATH="$d:$PATH" ;;
    esac
  fi
done
export PATH

IMG=node:20-slim
MIRROR=docker.m.daocloud.io/library/node:20-slim
NAME=client-base
PORT=5174
# 访问口令：优先读环境变量/同目录 .env，默认 xiao2ming（用户 2026-10-17 指定）
if [ -z "$CB_PASSWORD" ] && [ -f "$HERE/.env" ]; then
  CB_PASSWORD=$(sed -n 's/^CB_PASSWORD=//p' "$HERE/.env" | tail -1)
fi
CB_PASSWORD=${CB_PASSWORD:-xiao2ming}
export CB_PASSWORD

echo "==> 部署根目录: $ROOT"
mkdir -p "$ROOT/data/pages" "$ROOT/data/office" "$ROOT/data/logs"
# DS（OnlyOffice 文档服务）的挂载目录——compose 与 docker-run 两条路都要
mkdir -p "$ROOT/ds/logs" "$ROOT/ds/data" "$ROOT/ds/lib" "$ROOT/ds/db" "$ROOT/ds/rabbitmq"

# ── 代码目录：本脚本和 app/ 同级（tar 包解压后的形态）──────────────
if [ -d "$HERE/app" ]; then
  if [ "$HERE" != "$ROOT" ]; then
    echo "==> 复制代码到 $ROOT/app"
    mkdir -p "$ROOT/app"
    cp -a "$HERE/app/." "$ROOT/app/"
  fi
else
  echo "!! 没找到 app/ 目录——请先把部署包解压到 $ROOT 再执行本脚本" >&2
  exit 1
fi

if ! command -v docker >/dev/null 2>&1; then
  echo "!! 找不到 docker —— 请先在套件中心安装并启动 Container Manager" >&2
  exit 1
fi

# ── 首次：默认画布（只有 pages 为空时才灌，避免覆盖用户数据）───────
if [ -z "$(ls -A "$ROOT/data/pages" 2>/dev/null)" ] && [ -d "$ROOT/app/server-data/pages" ]; then
  echo "==> 首次安装，导入默认画布"
  cp -a "$ROOT/app/server-data/pages/." "$ROOT/data/pages/"
fi

# ── 镜像：没有就拉；官方源不通就换国内加速 ─────────────────────────
if docker image inspect "$IMG" >/dev/null 2>&1; then
  echo "==> 镜像 $IMG 已存在，跳过拉取"
else
  echo "==> 拉取 $IMG（首次约 80MB）"
  if ! docker pull "$IMG"; then
    echo "!! 官方源拉取失败，改用镜像加速：$MIRROR"
    docker pull "$MIRROR"
    docker tag "$MIRROR" "$IMG"
  fi
fi

# ── 起容器：优先 compose（Container Manager 里能看到项目），没有就 docker run ──
DC=""
if docker compose version >/dev/null 2>&1; then
  DC="docker compose"
elif command -v docker-compose >/dev/null 2>&1 && docker-compose version >/dev/null 2>&1; then
  DC="docker-compose"
fi

if [ -n "$DC" ] && [ -f "$ROOT/docker-compose.yml" ]; then
  echo "==> 用 $DC 启动"
  cd "$ROOT"
  $DC up -d --force-recreate
else
  echo "==> 用 docker run 启动（未检测到可用的 docker compose）"
  # DS（在线编辑内核）也一并起。镜像加速白名单不含 onlyoffice，只能走官方源
  DS_IMG=onlyoffice/documentserver:latest
  docker rm -f "$NAME-ds" >/dev/null 2>&1 || true
  mkdir -p "$ROOT/ds/logs" "$ROOT/ds/data" "$ROOT/ds/lib" "$ROOT/ds/db" "$ROOT/ds/rabbitmq"
  docker run -d --name "$NAME-ds" --restart unless-stopped \
    -e TZ=Asia/Shanghai -e JWT_ENABLED=false \
    -v "$ROOT/ds/logs:/var/log/onlyoffice" \
    -v "$ROOT/ds/data:/var/www/onlyoffice/Data" \
    -v "$ROOT/ds/lib:/var/lib/onlyoffice" \
    -v "$ROOT/ds/db:/var/lib/postgresql" \
    -v "$ROOT/ds/rabbitmq:/var/lib/rabbitmq" \
    -p 8080:80 \
    "$DS_IMG"
  docker rm -f "$NAME" >/dev/null 2>&1 || true
  # -w /data：审计日志写在 cwd/logs，/app 是只读挂载没法当挂载点（见 compose 注释）
  docker run -d --name "$NAME" --restart unless-stopped \
    -w /data \
    -e TZ=Asia/Shanghai \
    -e PORT="$PORT" -e HOST=0.0.0.0 \
    -e APP_DISABLE_OO=1 \
    -e APP_PASSWORD="$CB_PASSWORD" \
    -e APP_OO_DS_URL=http://"$NAME-ds" \
    -e APP_OO_DS_PUBLIC="http://${CB_DS_HOST:-192.168.1.16}:8080" \
    -e APP_OO_ORIGIN="http://${CB_DS_HOST:-192.168.1.16}:$PORT" \
    -e APP_DATA_DIR=/data/pages -e APP_OFFICE_DIR=/data/office \
    -v "$ROOT/app:/app:ro" \
    -v "$ROOT/data:/data" \
    -p "$PORT:$PORT" \
    "$IMG" node /app/app-server.mjs
fi

sleep 3
echo
echo "==> 状态"
docker ps --filter "name=$NAME" --format '{{.Names}}  {{.Status}}  {{.Ports}}' || true
echo
IP=$(ip route get 1.1.1.1 2>/dev/null | awk '{print $7; exit}')
echo "访问: http://${IP:-<群晖IP>}:$PORT/  （需要访问口令）"
echo "探针: http://${IP:-<群晖IP>}:$PORT/healthz"
echo "日志: $ROOT/data/logs/audit-$(date +%Y-%m-%d).log"
