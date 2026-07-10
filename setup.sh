#!/bin/bash
# ============================================================
# 一键初始化脚本
# 作用: 把 OnlyOffice SDK 静态资源复制到 Flask 的 static 目录
# ============================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SDK_SRC="/Volumes/SN770/载下/onlyoffice-web-comp-main/public/packages"
SDK_DST="$SCRIPT_DIR/static/packages"

echo "=== OnlyOffice 离线部署 - 初始化 ==="
echo ""

if [ ! -d "$SDK_SRC" ]; then
    echo "❌ 找不到 SDK 源目录: $SDK_SRC"
    echo "   请修改脚本中的 SDK_SRC 变量为 onlyoffice-web-comp 项目的 public/packages 路径"
    exit 1
fi

echo "📦 复制 SDK 静态资源（约 1GB，请耐心等待）..."
echo "   源: $SDK_SRC"
echo "   目标: $SDK_DST"
echo ""

mkdir -p "$(dirname "$SDK_DST")"
cp -r "$SDK_SRC" "$SDK_DST"

echo ""
echo "✅ 初始化完成！"
echo ""
echo "启动方式:"
echo "  cd $SCRIPT_DIR"
echo "  pip install -r requirements.txt"
echo "  python app.py"
echo ""
echo "然后访问: http://localhost:5000"
