#!/bin/bash
# PPT翻译工具 - 停止服务器
# 在macOS上可以直接双击运行

# 获取脚本所在目录
cd "$(dirname "$0")"

# 打开终端窗口（如果从Finder双击运行）
if [ -z "$TERM" ] || [ "$TERM" = "dumb" ]; then
    osascript -e 'tell application "Terminal" to do script "cd \"'"$(pwd)"'\" && bash \"'"$0"'\" && exit"'
    exit 0
fi

echo "=========================================="
echo "PPT翻译工具 - 停止服务器"
echo "=========================================="

if [ -f server.pid ]; then
    PIDS=$(cat server.pid)
    echo "正在停止服务器..."
    kill $PIDS 2>/dev/null
    rm -f server.pid
    echo "✅ 服务器已停止"
else
    echo "未找到server.pid文件，尝试查找并停止进程..."
    pkill -f "python3 app.py" 2>/dev/null
    pkill -f "python3 -m http.server 8014" 2>/dev/null
    echo "✅ 已尝试停止相关进程"
fi

echo ""
read -p "按回车键退出..."
