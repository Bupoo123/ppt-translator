#!/bin/bash
# 停止PPT翻译工具服务器

cd "$(dirname "$0")"

if [ -f server.pid ]; then
    PIDS=$(cat server.pid)
    echo "正在停止服务器..."
    kill $PIDS 2>/dev/null
    rm server.pid
    echo "✅ 服务器已停止"
else
    echo "未找到server.pid文件，尝试查找并停止进程..."
    pkill -f "python3 app.py"
    pkill -f "python3 -m http.server 8000"
    echo "✅ 已尝试停止相关进程"
fi

