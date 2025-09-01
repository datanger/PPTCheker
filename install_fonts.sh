#!/bin/bash

echo "正在为Ubuntu系统安装中文字体..."

# 更新包列表
sudo apt update

# 安装中文字体包
echo "安装文泉驿字体..."
sudo apt install -y fonts-wqy-microhei fonts-wqy-zenhei

echo "安装Noto字体..."
sudo apt install -y fonts-noto-cjk

echo "安装Ubuntu字体..."
sudo apt install -y fonts-ubuntu

echo "安装Liberation字体..."
sudo apt install -y fonts-liberation

echo "安装DejaVu字体..."
sudo apt install -y fonts-dejavu

# 刷新字体缓存
echo "刷新字体缓存..."
sudo fc-cache -fv

echo "字体安装完成！"
echo "请重新启动GUI应用程序以使用新安装的字体。"
