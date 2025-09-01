# Ubuntu字体问题解决指南

## 问题描述
在Ubuntu系统上运行PPT审查工具GUI时，可能出现中文字体显示为矩形小框（乱码）的问题。

## 原因分析
1. **缺少中文字体**：Ubuntu默认可能没有安装完整的中文字体
2. **字体配置问题**：Tkinter无法正确识别可用的中文字体
3. **字体缓存问题**：系统字体缓存未更新

## 解决方案

### 方案1：安装中文字体（推荐）
```bash
# 运行字体安装脚本
./install_fonts.sh

# 或者手动安装
sudo apt update
sudo apt install -y fonts-wqy-microhei fonts-wqy-zenhei fonts-noto-cjk
sudo fc-cache -fv
```

### 方案2：检查当前字体
```bash
# 运行字体检测脚本
python check_fonts.py

# 查看系统中文字体
fc-list :lang=zh
```

### 方案3：手动安装字体包
```bash
# 安装文泉驿字体
sudo apt install -y fonts-wqy-microhei fonts-wqy-zenhei

# 安装Noto字体
sudo apt install -y fonts-noto-cjk

# 安装Ubuntu字体
sudo apt install -y fonts-ubuntu

# 刷新字体缓存
sudo fc-cache -fv
```

## 字体优先级
GUI程序会按以下优先级选择字体：

1. **WenQuanYi Micro Hei** (文泉驿微米黑) - 最佳选择
2. **WenQuanYi Zen Hei** (文泉驿正黑)
3. **Noto Sans CJK SC** (Google Noto中文字体)
4. **Ubuntu** (Ubuntu默认字体)
5. **DejaVu Sans** (DejaVu字体)
6. **系统默认字体** (TkDefaultFont)

## 验证修复
安装字体后，重新运行GUI：
```bash
python run_gui.py
```

如果仍有问题，请运行字体检测脚本查看详细信息。

## 常见问题

### Q: 安装字体后仍然显示乱码？
A: 请重启应用程序或重新登录系统，让字体缓存生效。

### Q: 某些字符仍然显示为方框？
A: 可能是特定字符的字体支持问题，尝试安装更多字体包。

### Q: 字体显示效果不佳？
A: 可以尝试调整字体大小或使用不同的字体主题。
