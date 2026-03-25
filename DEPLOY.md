# 部署指南

## 服务器信息
- 服务器: ubuntu@172.17.71.123
- 密码: Matridx01
- 访问地址: http://172.17.71.123:8014

## 部署方式

### 方式1: Docker部署（推荐）

#### 本地准备
1. 确保已安装Docker和docker-compose
2. 构建镜像：
   ```bash
   docker build -t ppt-translator:latest .
   ```

#### 服务器端操作

1. **SSH连接到服务器**
   ```bash
   ssh ubuntu@172.17.71.123
   ```

2. **安装Docker（如果未安装）**
   ```bash
   curl -fsSL https://get.docker.com -o get-docker.sh
   sudo sh get-docker.sh
   sudo usermod -aG docker $USER
   ```

3. **安装docker-compose**
   ```bash
   sudo curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
   sudo chmod +x /usr/local/bin/docker-compose
   ```

4. **克隆或上传代码到服务器**
   ```bash
   cd ~
   git clone https://github.com/bupoo123/ppt-translator.git
   cd ppt-translator
   ```

5. **创建.env文件**
   ```bash
   nano .env
   ```
   内容：
   ```
   DEEPSEEK_API_KEY=your_api_key_here
   API_PROVIDER=deepseek
   ```

6. **启动服务（推荐使用生产配置）**
   ```bash
   # 使用生产配置（前后端分离）
   docker-compose -f docker-compose.prod.yml up -d
   
   # 或使用单容器配置
   docker-compose up -d
   ```

7. **查看日志**
   ```bash
   docker-compose logs -f
   ```

8. **停止服务**
   ```bash
   docker-compose down
   ```

### 方式2: 直接部署（不使用Docker）

1. **SSH连接到服务器**
   ```bash
   ssh ubuntu@172.17.71.123
   ```

2. **安装依赖**
   ```bash
   sudo apt update
   sudo apt install -y python3 python3-pip
   ```

3. **克隆代码**
   ```bash
   cd ~
   git clone https://github.com/bupoo123/ppt-translator.git
   cd ppt-translator
   ```

4. **安装Python依赖**
   ```bash
   pip3 install --user -r requirements.txt
   ```

5. **创建.env文件**
   ```bash
   nano .env
   ```

6. **使用systemd管理服务**
   ```bash
   # 创建服务文件
   sudo nano /etc/systemd/system/ppt-translator.service
   ```
   
   内容：
   ```ini
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
   Environment="PATH=/usr/bin:/usr/local/bin"

   [Install]
   WantedBy=multi-user.target
   ```

7. **启动服务**
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable ppt-translator
   sudo systemctl start ppt-translator
   sudo systemctl status ppt-translator
   ```

8. **启动前端服务（需要单独处理）**
   可以使用nginx反向代理，或者使用screen/tmux运行：
   ```bash
   screen -S frontend
   cd ~/ppt-translator/frontend
   python3 -m http.server 8014
   # 按Ctrl+A然后D退出screen
   ```

## 使用自动部署脚本

如果本地有sshpass工具，可以使用自动部署脚本：

```bash
# 安装sshpass（macOS）
brew install hudochenkov/sshpass/sshpass

# 运行部署脚本
chmod +x deploy.sh
./deploy.sh
```

## 防火墙配置

确保服务器防火墙开放相应端口：

```bash
sudo ufw allow 5014/tcp
sudo ufw allow 8014/tcp
sudo ufw reload
```

## 访问地址

部署成功后，访问：
- 前端界面: http://172.17.71.123:8014
- 后端API: http://172.17.71.123:5014

## 更新部署

```bash
# 在服务器上
cd ~/ppt-translator
git pull
docker-compose restart  # 如果使用Docker
# 或
sudo systemctl restart ppt-translator  # 如果使用systemd
```
