#!/bin/bash
# Docker容器启动脚本

# 创建必要的目录
mkdir -p uploads outputs

# 启动后端服务（后台运行）
echo "启动后端服务器 (端口5014)..."
python3 app.py > server.log 2>&1 &
BACKEND_PID=$!
echo $BACKEND_PID > /tmp/backend.pid

# 等待后端启动
sleep 3

# 启动前端服务（后台运行）
echo "启动前端服务器 (端口8014)..."
cd frontend
python3 -m http.server 8014 > ../frontend.log 2>&1 &
FRONTEND_PID=$!
echo $FRONTEND_PID > /tmp/frontend.pid
cd ..

# 等待服务启动
sleep 2

# 检查服务状态
if curl -s http://localhost:5014/health > /dev/null 2>&1; then
    echo "✅ 后端服务器运行正常"
else
    echo "⚠️  后端服务器可能未正常启动"
fi

echo "✅ 所有服务已启动"
echo "前端: http://localhost:8014"
echo "后端: http://localhost:5014"

# 保持容器运行
tail -f server.log frontend.log
