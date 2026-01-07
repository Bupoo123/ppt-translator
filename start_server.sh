#!/bin/bash
# 启动PPT翻译工具服务器

cd "$(dirname "$0")"

echo "=========================================="
echo "PPT翻译工具 - 启动服务器"
echo "=========================================="

# 检查.env文件
if [ ! -f .env ]; then
    echo "⚠️  警告: 未找到.env文件"
    echo "请创建.env文件并设置DEEPSEEK_API_KEY"
    echo ""
fi

# 创建必要的目录
mkdir -p uploads outputs

# 检查端口是否被占用
if lsof -Pi :5014 -sTCP:LISTEN -t >/dev/null ; then
    echo "⚠️  端口5014已被占用，请先停止其他服务"
    exit 1
fi

if lsof -Pi :8014 -sTCP:LISTEN -t >/dev/null ; then
    echo "⚠️  端口8014已被占用，请先停止其他服务"
    exit 1
fi

echo "正在启动后端服务器 (端口5014)..."
python3 app.py > server.log 2>&1 &
BACKEND_PID=$!
echo "后端服务器已启动 (PID: $BACKEND_PID)"

sleep 2

echo "正在启动前端服务器 (端口8014)..."
cd frontend
python3 -m http.server 8014 > ../frontend.log 2>&1 &
FRONTEND_PID=$!
cd ..
echo "前端服务器已启动 (PID: $FRONTEND_PID)"

sleep 2

# 检查服务器是否正常运行
if curl -s http://localhost:5014/health > /dev/null; then
    echo "✅ 后端服务器运行正常"
else
    echo "❌ 后端服务器启动失败，请检查server.log"
fi

if curl -s http://localhost:8014 > /dev/null; then
    echo "✅ 前端服务器运行正常"
else
    echo "❌ 前端服务器启动失败，请检查frontend.log"
fi

echo ""
echo "=========================================="
echo "服务器启动完成！"
echo "=========================================="
echo ""
echo "访问地址:"
echo "  前端界面: http://localhost:8014"
echo "  后端API:  http://localhost:5014"
echo ""
echo "停止服务器:"
echo "  kill $BACKEND_PID $FRONTEND_PID"
echo "  或运行: ./stop_server.sh"
echo ""
echo "日志文件:"
echo "  后端日志: server.log"
echo "  前端日志: frontend.log"
echo ""

# 保存PID到文件
echo "$BACKEND_PID $FRONTEND_PID" > server.pid

