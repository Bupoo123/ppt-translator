#!/bin/bash
# 安装依赖并运行测试的脚本

cd /Users/bupoo/Github/ppt-translator

# 检查是否在虚拟环境中
if [ -z "$VIRTUAL_ENV" ]; then
    echo "建议使用虚拟环境，正在创建..."
    python3 -m venv venv
    source venv/bin/activate
fi

# 安装依赖
echo "正在安装依赖..."
pip install python-pptx

# 运行测试
echo "开始测试..."
python3 test_step1.py /Users/bupoo/Github/ppt-translator/testppt/test.pptx test_output.pptx

