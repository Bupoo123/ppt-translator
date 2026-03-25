#!/bin/bash
# 快速部署脚本 - 一键部署到服务器

SERVER_USER="ubuntu"
SERVER_HOST="172.17.71.123"
SERVER_PASSWORD="Matridx01"
SERVER_DIR="~/ppt-translator"

echo "=========================================="
echo "PPT翻译工具 - 快速部署"
echo "=========================================="
echo "服务器: $SERVER_USER@$SERVER_HOST"
echo ""

# 检查sshpass是否安装
if ! command -v sshpass &> /dev/null; then
    echo "⚠️  未安装sshpass，将使用交互式SSH"
    echo "macOS安装: brew install hudochenkov/sshpass/sshpass"
    USE_SSHPASS=false
else
    USE_SSHPASS=true
fi

# 函数：执行SSH命令
run_ssh() {
    if [ "$USE_SSHPASS" = true ]; then
        sshpass -p "$SERVER_PASSWORD" ssh -o StrictHostKeyChecking=no $SERVER_USER@$SERVER_HOST "$1"
    else
        ssh -o StrictHostKeyChecking=no $SERVER_USER@$SERVER_HOST "$1"
    fi
}

# 函数：传输文件
transfer_file() {
    if [ "$USE_SSHPASS" = true ]; then
        sshpass -p "$SERVER_PASSWORD" scp -o StrictHostKeyChecking=no "$1" $SERVER_USER@$SERVER_HOST:"$2"
    else
        scp -o StrictHostKeyChecking=no "$1" $SERVER_USER@$SERVER_HOST:"$2"
    fi
}

echo "1. 连接到服务器并准备环境..."
run_ssh "mkdir -p $SERVER_DIR"

echo "2. 检查Docker安装..."
run_ssh "if ! command -v docker &> /dev/null; then
    echo '安装Docker...'
    curl -fsSL https://get.docker.com -o /tmp/get-docker.sh
    sudo sh /tmp/get-docker.sh
    sudo usermod -aG docker \$USER
    newgrp docker
fi"

echo "3. 检查docker-compose安装..."
run_ssh "if ! command -v docker-compose &> /dev/null; then
    echo '安装docker-compose...'
    sudo curl -L \"https://github.com/docker/compose/releases/latest/download/docker-compose-\$(uname -s)-\$(uname -m)\" -o /usr/local/bin/docker-compose
    sudo chmod +x /usr/local/bin/docker-compose
fi"

echo "4. 传输项目文件..."
# 创建临时目录
TEMP_DIR=$(mktemp -d)
cd "$(dirname "$0")"

# 复制必要文件（排除不需要的）
rsync -avz --exclude='.git' \
    --exclude='testppt' \
    --exclude='*.pptx' \
    --exclude='*.ppt' \
    --exclude='__pycache__' \
    --exclude='.env' \
    --exclude='uploads' \
    --exclude='outputs' \
    --exclude='*.log' \
    --exclude='server.pid' \
    ./ $TEMP_DIR/ppt-translator/

# 压缩并传输
cd $TEMP_DIR
tar czf ppt-translator.tar.gz ppt-translator/

if [ "$USE_SSHPASS" = true ]; then
    sshpass -p "$SERVER_PASSWORD" scp -o StrictHostKeyChecking=no ppt-translator.tar.gz $SERVER_USER@$SERVER_HOST:$SERVER_DIR/
else
    scp -o StrictHostKeyChecking=no ppt-translator.tar.gz $SERVER_USER@$SERVER_HOST:$SERVER_DIR/
fi

rm -rf $TEMP_DIR

echo "5. 在服务器上解压和部署..."
run_ssh "cd $SERVER_DIR && 
    tar xzf ppt-translator.tar.gz && 
    cd ppt-translator && 
    rm ppt-translator.tar.gz &&
    if [ ! -f .env ]; then
        echo '请创建.env文件并设置API密钥'
    fi &&
    docker-compose -f docker-compose.prod.yml down 2>/dev/null &&
    docker-compose -f docker-compose.prod.yml build &&
    docker-compose -f docker-compose.prod.yml up -d &&
    sleep 5 &&
    docker-compose -f docker-compose.prod.yml ps"

echo ""
echo "=========================================="
echo "✅ 部署完成！"
echo "=========================================="
echo ""
echo "访问地址: http://172.17.71.123:8014"
echo ""
echo "查看日志:"
echo "  ssh $SERVER_USER@$SERVER_HOST"
echo "  cd $SERVER_DIR/ppt-translator"
echo "  docker-compose -f docker-compose.prod.yml logs -f"
echo ""
echo "停止服务:"
echo "  docker-compose -f docker-compose.prod.yml down"
echo ""
