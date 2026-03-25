#!/bin/bash
# 部署脚本 - 部署到Ubuntu服务器

SERVER_USER="ubuntu"
SERVER_HOST="172.17.71.123"
SERVER_PASSWORD="Matridx01"
SERVER_DIR="~/ppt-translator"

echo "=========================================="
echo "PPT翻译工具 - 部署到服务器"
echo "=========================================="
echo "服务器: $SERVER_USER@$SERVER_HOST"
echo ""

# 检查Docker是否安装
if ! command -v docker &> /dev/null; then
    echo "❌ 本地未安装Docker，将使用SSH直接部署"
    USE_DOCKER=false
else
    echo "✅ 检测到Docker，将构建镜像并部署"
    USE_DOCKER=true
fi

# 方法1: 使用Docker部署（推荐）
if [ "$USE_DOCKER" = true ]; then
    echo ""
    echo "使用Docker部署..."
    
    # 构建Docker镜像
    echo "1. 构建Docker镜像..."
    docker build -t ppt-translator:latest .
    
    # 保存镜像
    echo "2. 保存Docker镜像..."
    docker save ppt-translator:latest | gzip > ppt-translator.tar.gz
    
    # 传输到服务器
    echo "3. 传输镜像到服务器..."
    sshpass -p "$SERVER_PASSWORD" scp ppt-translator.tar.gz docker-compose.yml docker-entrypoint.sh $SERVER_USER@$SERVER_HOST:$SERVER_DIR/
    
    # 在服务器上部署
    echo "4. 在服务器上部署..."
    sshpass -p "$SERVER_PASSWORD" ssh $SERVER_USER@$SERVER_HOST << 'ENDSSH'
        cd ~/ppt-translator
        
        # 安装Docker（如果未安装）
        if ! command -v docker &> /dev/null; then
            echo "安装Docker..."
            curl -fsSL https://get.docker.com -o get-docker.sh
            sudo sh get-docker.sh
            sudo usermod -aG docker $USER
        fi
        
        # 安装docker-compose（如果未安装）
        if ! command -v docker-compose &> /dev/null; then
            echo "安装docker-compose..."
            sudo curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
            sudo chmod +x /usr/local/bin/docker-compose
        fi
        
        # 加载镜像
        echo "加载Docker镜像..."
        docker load < ppt-translator.tar.gz
        
        # 创建.env文件（如果不存在）
        if [ ! -f .env ]; then
            echo "请创建.env文件并设置API密钥"
        fi
        
        # 启动服务
        echo "启动服务..."
        docker-compose down 2>/dev/null
        docker-compose up -d
        
        echo "✅ 部署完成！"
        echo "访问地址: http://172.17.71.123:8014"
ENDSSH

    # 清理本地镜像文件
    rm -f ppt-translator.tar.gz
    
else
    # 方法2: 直接部署（不使用Docker）
    echo ""
    echo "直接部署到服务器..."
    
    # 传输文件
    echo "1. 传输文件到服务器..."
    sshpass -p "$SERVER_PASSWORD" rsync -avz --exclude='.git' --exclude='testppt' --exclude='*.pptx' --exclude='*.ppt' --exclude='__pycache__' --exclude='.env' \
        ./ $SERVER_USER@$SERVER_HOST:$SERVER_DIR/
    
    # 在服务器上设置
    echo "2. 在服务器上设置环境..."
    sshpass -p "$SERVER_PASSWORD" ssh $SERVER_USER@$SERVER_HOST << 'ENDSSH'
        cd ~/ppt-translator
        
        # 安装Python依赖
        echo "安装Python依赖..."
        pip3 install --user -r requirements.txt
        
        # 创建必要目录
        mkdir -p uploads outputs
        
        # 创建systemd服务文件
        echo "创建systemd服务..."
        sudo tee /etc/systemd/system/ppt-translator.service > /dev/null << 'EOF'
[Unit]
Description=PPT Translator Service
After=network.target

[Service]
Type=simple
User=ubuntu
WorkingDirectory=/home/ubuntu/ppt-translator
ExecStart=/usr/bin/python3 app.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF
        
        sudo systemctl daemon-reload
        sudo systemctl enable ppt-translator
        sudo systemctl restart ppt-translator
        
        echo "✅ 部署完成！"
        echo "访问地址: http://172.17.71.123:5014"
ENDSSH
fi

echo ""
echo "=========================================="
echo "部署完成！"
echo "=========================================="
echo "访问地址: http://172.17.71.123:8014"
echo ""
