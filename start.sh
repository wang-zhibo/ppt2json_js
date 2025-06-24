#!/bin/bash

echo "🚀 启动PPT转JSON工具..."

# 检查Node.js是否安装
if ! command -v node &> /dev/null; then
    echo "❌ 错误: 未找到Node.js，请先安装Node.js"
    exit 1
fi

# 检查npm是否安装
if ! command -v npm &> /dev/null; then
    echo "❌ 错误: 未找到npm，请先安装npm"
    exit 1
fi

echo "📦 安装后端依赖..."
npm install

echo "📦 安装前端依赖..."
cd client && npm install --legacy-peer-deps && cd ..

echo "🔨 构建前端应用..."
npm run build

echo "🌐 启动服务器..."
echo "✅ 应用将在 http://localhost:5001 启动"
echo "📡 API接口: http://localhost:5001/api/convert"
echo ""
echo "按 Ctrl+C 停止服务器"

npm start 