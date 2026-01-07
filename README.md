# PPT 翻译工具

轻松将中文 PowerPoint 简报转换为英文，完整保留原始格式。

## 功能特性

- ✅ 保留原始格式（布局、字体、颜色、图表、图片、动画）
- ✅ 精准定位并翻译文本内容
- ✅ 支持术语表，保证翻译一致性
- ✅ Slide 级上下文翻译，生成自然的英文表达
- ✅ 自动跳过英文内容
- ✅ 智能处理文本溢出问题

## 技术栈

- 后端：Python + Flask + python-pptx
- 前端：HTML + JavaScript
- AI翻译：DeepSeek API / OpenAI API（可切换）

## 安装

```bash
pip install -r requirements.txt
```

## 使用

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置环境变量

创建 `.env` 文件（参考 `.env.example`）：

**使用 DeepSeek API（推荐）：**
```
DEEPSEEK_API_KEY=your_deepseek_api_key_here
API_PROVIDER=deepseek
```

**使用 OpenAI API：**
```
OPENAI_API_KEY=your_openai_api_key_here
API_PROVIDER=openai
```

> **获取 DeepSeek API Key：**
> 1. 访问 [DeepSeek 开放平台](https://platform.deepseek.com/)
> 2. 注册/登录账号
> 3. 在控制台创建 API Key
> 4. 将 API Key 复制到 `.env` 文件中

### 3. 测试功能

**测试PPT解析功能（不需要API密钥）：**
```bash
python3 test_step1.py your_file.pptx output.pptx
```

**测试翻译功能（需要API密钥）：**
```bash
python3 test_translator.py
```

### 4. 启动Web服务

**方法1：使用启动脚本（推荐）**
```bash
./start_server.sh
```

**方法2：手动启动**
```bash
# 启动后端服务（终端1）
python3 app.py

# 启动前端服务（终端2）
cd frontend
python3 -m http.server 8014
```

### 5. 访问Web界面

- 打开浏览器访问：**http://localhost:8014**
- 上传PPT文件，等待翻译完成
- 点击下载按钮获取翻译后的PPT

### 6. 停止服务

```bash
./stop_server.sh
```

## 项目结构

```
ppt-translator/
├── app.py                 # Flask后端应用
├── ppt_processor.py       # PPT处理核心模块
├── translator.py          # AI翻译模块
├── test_step1.py         # PPT解析测试脚本
├── test_translator.py    # 翻译功能测试脚本
├── requirements.txt       # Python依赖
├── frontend/
│   └── index.html        # 前端界面
└── README.md             # 项目说明
```

## 开发路线

- [x] 第一步：PPT解析和文本修改 ✅
- [x] 第二步：AI翻译集成 ✅
- [x] 第三步：Slide级上下文翻译 ✅
- [x] 前端界面 ✅
- [x] 后端API ✅
- [ ] 术语表功能（待完善）
- [ ] 批量处理功能
- [ ] 文本溢出智能处理

## 注意事项

1. **API密钥**：
   - **DeepSeek API（推荐）**：访问 [DeepSeek 开放平台](https://platform.deepseek.com/) 获取API Key，性价比高
   - **OpenAI API**：访问 [OpenAI Platform](https://platform.openai.com/) 获取API Key
   - 在 `.env` 文件中设置对应的API密钥和提供商（`API_PROVIDER=deepseek` 或 `API_PROVIDER=openai`）
2. **文件格式**：目前支持 `.pptx` 格式（`.ppt` 格式需要转换）
3. **文本提取**：自动跳过纯数字、纯英文内容（中文比例<20%）
4. **翻译质量**：
   - DeepSeek：使用 `deepseek-chat` 模型，性价比高，适合中文翻译
   - OpenAI：使用 `gpt-4` 模型，可根据需要切换到 `gpt-3.5-turbo` 降低成本
5. **API切换**：通过设置 `API_PROVIDER` 环境变量切换API提供商

