#!/usr/bin/env bash
set -euo pipefail

if [[ ${EUID} -ne 0 ]]; then
  echo "请使用 sudo bash 运行"
  exit 1
fi

APP_DIR=/opt/honghe-tender
TMP_DIR=$(mktemp -d)
trap 'rm -rf "$TMP_DIR"' EXIT

mkdir -p "$APP_DIR/output/publish"
chown -R honghe:honghe "$APP_DIR"

curl -fL --connect-timeout 15 --max-time 180 \
  -o "$TMP_DIR/deploy.zip" \
  https://sadfdsfdsffewf.github.io/honghe-tender-plan/deploy/honghe-tender-tencent-deploy.zip
unzip -o "$TMP_DIR/deploy.zip" -d "$TMP_DIR/pkg" >/dev/null

install -o honghe -g honghe -m 0644 "$TMP_DIR/pkg/tencent-cloud-deploy/daily_plan_workflow.py" "$APP_DIR/daily_plan_workflow.py"
install -o honghe -g honghe -m 0755 "$TMP_DIR/pkg/tencent-cloud-deploy/run_update.sh" "$APP_DIR/run_update.sh"
install -o root -g root -m 0644 "$TMP_DIR/pkg/tencent-cloud-deploy/honghe-tender.timer" /etc/systemd/system/honghe-tender.timer

if [[ -f /etc/caddy/Caddyfile && ! -f /etc/caddy/Caddyfile.before-honghe ]]; then
  cp -a /etc/caddy/Caddyfile /etc/caddy/Caddyfile.before-honghe
fi

cat >/etc/caddy/Caddyfile <<'EOF'
:80 {
    root * /opt/honghe-tender/output/publish
    encode gzip
    file_server
    try_files {path} /index.html
    header Cache-Control "no-cache, no-store, must-revalidate"
}
EOF

systemctl daemon-reload
systemctl restart honghe-tender.timer
systemctl reload caddy || systemctl restart caddy
systemctl restart honghe-tender.service

echo "直接访问地址：http://118.24.53.186/"
echo "GitHub 断网不会影响此页面更新。"