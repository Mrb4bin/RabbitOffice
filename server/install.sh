#!/usr/bin/env bash
# Rabbit Office Server installer for Debian/Ubuntu/Fedora/RHEL/Arch.
# Usage (from the unpacked project folder):  sudo ./server/install.sh
set -euo pipefail
GREEN='\033[0;32m'; NC='\033[0m'
say() { echo -e "${GREEN}==>${NC} $*"; }
[ "$(id -u)" -eq 0 ] || { echo "Запустите через sudo"; exit 1; }
SRC="$(cd "$(dirname "$0")/.." && pwd)"
APP=/opt/rabbit-office; DATA=/var/lib/rabbit-office

need_node() { command -v node >/dev/null && [ "$(node -p 'process.versions.node.split(".")[0]')" -ge 18 ]; }
if ! need_node; then
  say "Устанавливаю Node.js 20…"
  if command -v apt-get >/dev/null; then apt-get update -y && apt-get install -y curl ca-certificates openssl && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && apt-get install -y nodejs
  elif command -v dnf >/dev/null; then curl -fsSL https://rpm.nodesource.com/setup_20.x | bash - && dnf install -y nodejs openssl
  elif command -v pacman >/dev/null; then pacman -Sy --noconfirm nodejs npm openssl
  else echo "Установите Node.js 18+ вручную"; exit 1; fi
fi

say "Копирую файлы в $APP"
mkdir -p "$APP" "$DATA"
cp -r "$SRC/server" "$SRC/shared" "$SRC/web" "$APP/"
cd "$APP/server" && npm install --omit=dev --no-audit --no-fund >/dev/null

id rabbit >/dev/null 2>&1 || useradd --system --home "$DATA" --shell /usr/sbin/nologin rabbit
chown -R rabbit:rabbit "$DATA"
ln -sf "$APP/server/bin/rabbit-office.js" /usr/local/bin/rabbit-office
chmod +x "$APP/server/bin/rabbit-office.js"

if [ -f "$DATA/config.json" ]; then
  say "Найдена существующая установка — обновляю программу, данные и настройки сохраняются"
  UPDATE=1
else
  say "Запускаю мастер настройки"
  RABBIT_DATA="$DATA" node "$APP/server/bin/rabbit-office.js" setup
fi
chown -R rabbit:rabbit "$DATA"

cat > /etc/systemd/system/rabbit-office.service <<EOF
[Unit]
Description=Rabbit Office Server
After=network-online.target
Wants=network-online.target

[Service]
User=rabbit
Environment=RABBIT_DATA=$DATA
ExecStart=$(command -v node) $APP/server/src/server.js
Restart=always
RestartSec=3
AmbientCapabilities=CAP_NET_BIND_SERVICE
NoNewPrivileges=true
ProtectSystem=full
ReadWritePaths=$DATA

[Install]
WantedBy=multi-user.target
EOF
systemctl daemon-reload
systemctl enable --now rabbit-office
[ "${UPDATE:-0}" = 1 ] && systemctl restart rabbit-office

if command -v ufw >/dev/null && ufw status | grep -q active; then
  read -rp "Открыть порты 80, 443 и 8080 в ufw? [Y/n] " a; [[ "${a:-y}" =~ ^[Yy]$ ]] && ufw allow 80,443,8080/tcp
elif command -v firewall-cmd >/dev/null && firewall-cmd --state >/dev/null 2>&1; then
  firewall-cmd --permanent --add-port={80,443,8080}/tcp && firewall-cmd --reload
fi

say "Готово! Служба: systemctl status rabbit-office · журнал: journalctl -u rabbit-office -f"
say "Управление пользователями: sudo -u rabbit RABBIT_DATA=$DATA rabbit-office user add <логин>"
