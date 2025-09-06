"""
PPT审查工具 - 简化GUI启动器（用于exe版本）

功能：
- 选择PPT文件
- 选择输出目录
- 配置LLM设置
- 运行审查
- 显示成功提示
- 实时显示控制台输出
"""
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
try:
    import yaml
except ImportError:
    import PyYAML as yaml
from datetime import datetime
import io
import contextlib

def get_resource_path(relative_path):
    """获取资源文件的绝对路径，兼容开发环境和打包环境"""
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        # 开发环境：使用当前文件所在目录
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)

# 添加项目路径
if not hasattr(sys, '_MEIPASS'):
    # 开发环境：添加项目根目录到路径
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 兼容性导入 - 支持开发环境和打包环境

from pptlint.config import load_config, ToolConfig
from pptlint.workflow import run_review_workflow
from pptlint.llm import LLMClient
from pptlint.parser import parse_pptx
from pptlint.cli import generate_output_paths
print("✅ 使用绝对导入模式")



class ConsoleCapture:
    """控制台输出捕获器"""
    def __init__(self, log_callback):
        self.log_callback = log_callback
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr
        self.stdout_buffer = io.StringIO()
        self.stderr_buffer = io.StringIO()
    
    def __enter__(self):
        # 创建自定义的输出流，实时回调
        class RealTimeStream:
            def __init__(self, original_stream, callback, prefix=""):
                self.original_stream = original_stream
                self.callback = callback
                self.prefix = prefix
                self.buffer = ""
            
            def write(self, text):
                # 安全写入原始流
                try:
                    if self.original_stream and hasattr(self.original_stream, 'write'):
                        self.original_stream.write(text)
                except Exception as e:
                    # 如果原始流写入失败，忽略错误
                    pass
                
                # 实时回调到GUI
                try:
                    if self.callback:
                        self.callback(text)
                except Exception as e:
                    # 如果回调失败，忽略错误
                    pass
            
            def flush(self):
                try:
                    if self.original_stream and hasattr(self.original_stream, 'flush'):
                        self.original_stream.flush()
                except Exception:
                    pass
            
            def close(self):
                pass
        
        # 替换标准输出和错误流
        sys.stdout = RealTimeStream(self.original_stdout, self.log_callback)
        sys.stderr = RealTimeStream(self.original_stderr, lambda x: self.log_callback(f"错误: {x}"))
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        sys.stdout = self.original_stdout
        sys.stderr = self.original_stderr
        # 安全关闭缓冲区
        try:
            if hasattr(self, 'stdout_buffer') and self.stdout_buffer:
                self.stdout_buffer.close()
        except Exception:
            pass
        try:
            if hasattr(self, 'stderr_buffer') and self.stderr_buffer:
                self.stderr_buffer.close()
        except Exception:
            pass


class SimpleApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PPT审查工具")
        
        # 获取屏幕尺寸并计算合适的窗口大小
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 计算窗口大小：屏幕宽度的80%，高度的85%，但不超过1200x900
        window_width = min(int(screen_width * 0.8), 1200)
        window_height = min(int(screen_height * 0.85), 900)
        
        # 确保最小尺寸
        window_width = max(window_width, 800)
        window_height = max(window_height, 600)
        
        # 居中显示
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.resizable(True, True)
        
        # 设置最小窗口大小
        self.minsize(800, 600)
        
        # 设置更好的字体
        self._setup_fonts()
        
        # 配置变量
        self.input_ppt = tk.StringVar()
        self.output_dir = tk.StringVar(value="output")
        self.llm_enabled = tk.BooleanVar(value=True)
        self.llm_provider = tk.StringVar(value="deepseek")
        self.llm_model = tk.StringVar(value="deepseek-chat")
        self.llm_api_key = tk.StringVar()
        self.mode = tk.StringVar(value="review")
        
        # 审查设置变量
        self.review_format = tk.BooleanVar(value=True)
        self.review_logic = tk.BooleanVar(value=True)
        self.review_acronyms = tk.BooleanVar(value=True)
        self.review_fluency = tk.BooleanVar(value=True)
        self.font_family = tk.BooleanVar(value=True)
        self.font_size = tk.BooleanVar(value=True)
        self.color_count = tk.BooleanVar(value=True)
        self.theme_harmony = tk.BooleanVar(value=True)
        self.acronym_explanation = tk.BooleanVar(value=True)
        
        # 审查规则配置变量
        self.jp_font_name = tk.StringVar(value="Meiryo UI")
        self.min_font_size_pt = tk.IntVar(value=12)
        self.color_count_threshold = tk.IntVar(value=5)
        
        # 控制台捕获器
        self.console_capture = None
        
        self._build_ui()
        self._load_default_config()

    def _setup_fonts(self):
        """设置字体样式 - Ubuntu优化版本"""
        try:
            # Ubuntu系统推荐字体
            default_font = ('WenQuanYi Micro Hei', 9)  # 文泉驿微米黑
            self.title_font = ('WenQuanYi Micro Hei', 12, 'bold')
            self.log_font = ('DejaVu Sans Mono', 8)
            
            # 配置ttk样式
            style = ttk.Style()
            style.theme_use('clam')
            
            # 设置控件字体
            style.configure('TLabel', font=default_font)
            style.configure('TButton', font=default_font)
            style.configure('TEntry', font=default_font)
            style.configure('TCombobox', font=default_font)
            style.configure('TCheckbutton', font=default_font)
            style.configure('TLabelframe.Label', font=default_font)
            
            # 尝试修改复选框的选中标记为√
            try:
                # 方法1：尝试使用不同的主题
                available_themes = style.theme_names()
                print(f"可用主题: {available_themes}")
                
                # 尝试使用alt主题，它通常有更好的复选框样式
                if 'alt' in available_themes:
                    style.theme_use('alt')
                    print("✅ 使用alt主题")
                elif 'default' in available_themes:
                    style.theme_use('default')
                    print("✅ 使用default主题")
                
                # 重新配置复选框样式
                style.configure('TCheckbutton', font=default_font)
                
                # 方法2：尝试修改复选框的映射
                style.map('TCheckbutton',
                         indicatorcolor=[('selected', 'black'),
                                       ('!selected', 'white')],
                         background=[('active', 'white'),
                                   ('!active', 'white')])
                
                print("✅ 复选框样式修改完成")
                
            except Exception as e:
                print(f"⚠️ 复选框样式修改失败: {e}")
            
            print("使用Ubuntu优化字体设置")
                
        except Exception as e:
            print(f"字体设置失败: {e}")
            # 使用系统默认字体
            self.title_font = ('TkHeadingFont', 12, 'bold')
            self.log_font = ('TkFixedFont', 8)

    def _build_ui(self):
        """构建UI界面"""
        # 创建主容器
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="PPT审查工具", font=self.title_font)
        title_label.pack(pady=(0, 20))
        
        # 第一行：文件上传窗口和LLM配置窗口并排排列
        first_row_frame = ttk.Frame(main_frame)
        first_row_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 文件上传窗口（5/10宽度）
        file_frame = ttk.LabelFrame(first_row_frame, text="📁 文件上传窗口", padding="15")
        file_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # PPT文件选择
        ppt_frame = ttk.Frame(file_frame)
        ppt_frame.pack(fill=tk.X, pady=8)
        ttk.Label(ppt_frame, text="PPT文件:", width=12).pack(side=tk.LEFT)
        ttk.Entry(ppt_frame, textvariable=self.input_ppt).pack(side=tk.LEFT, padx=(8, 8), fill=tk.X, expand=True)
        ttk.Button(ppt_frame, text="选择", command=self._select_ppt, width=10).pack(side=tk.LEFT)
        
        # 输出目录选择
        output_frame = ttk.Frame(file_frame)
        output_frame.pack(fill=tk.X, pady=8)
        ttk.Label(output_frame, text="输出目录:", width=12).pack(side=tk.LEFT)
        ttk.Entry(output_frame, textvariable=self.output_dir).pack(side=tk.LEFT, padx=(8, 8), fill=tk.X, expand=True)
        ttk.Button(output_frame, text="选择", command=self._select_output_dir, width=10).pack(side=tk.LEFT)
        
        # 运行模式
        mode_frame = ttk.Frame(file_frame)
        mode_frame.pack(fill=tk.X, pady=8)
        ttk.Label(mode_frame, text="运行模式:", width=12).pack(side=tk.LEFT)
        mode_combo = ttk.Combobox(mode_frame, textvariable=self.mode, values=["review", "edit"], 
                                 state="readonly", width=20)
        mode_combo.pack(side=tk.LEFT, padx=(8, 0))
        
        # LLM配置窗口（5/10宽度）
        llm_frame = ttk.LabelFrame(first_row_frame, text="🤖 LLM配置窗口", padding="15")
        llm_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 提供商选择
        provider_frame = ttk.Frame(llm_frame)
        provider_frame.pack(fill=tk.X, pady=8)
        ttk.Label(provider_frame, text="提供商:", width=12).pack(side=tk.LEFT)
        provider_combo = ttk.Combobox(provider_frame, textvariable=self.llm_provider, 
                                     values=["deepseek", "openai", "anthropic", "local"], 
                                     state="readonly", width=20)
        provider_combo.pack(side=tk.LEFT, padx=(8, 0))
        provider_combo.bind('<<ComboboxSelected>>', self._on_provider_change)
        
        # 模型选择
        model_frame = ttk.Frame(llm_frame)
        model_frame.pack(fill=tk.X, pady=8)
        ttk.Label(model_frame, text="模型:", width=12).pack(side=tk.LEFT)
        self.model_combo = ttk.Combobox(model_frame, textvariable=self.llm_model, 
                                       state="readonly", width=20)
        self.model_combo.pack(side=tk.LEFT, padx=(8, 0))
        
        # API密钥
        api_frame = ttk.Frame(llm_frame)
        api_frame.pack(fill=tk.X, pady=8)
        ttk.Label(api_frame, text="API密钥:", width=12).pack(side=tk.LEFT)
        api_entry = ttk.Entry(api_frame, textvariable=self.llm_api_key, show="*")
        api_entry.pack(side=tk.LEFT, padx=(8, 8), fill=tk.X, expand=True)
        ttk.Button(api_frame, text="应用", command=self._apply_api_key, width=10).pack(side=tk.LEFT)
        
        # 初始化模型列表
        self._update_model_list()
        
        # 第二行：审查配置窗口（10/10宽度，全宽）- 增加高度
        review_frame = ttk.LabelFrame(main_frame, text="⚙️ 审查配置窗口", padding="15")
        review_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 创建审查设置
        self._create_review_settings(review_frame)
        
        # 区域3：开始运行按钮 - 进一步压缩高度
        run_frame = ttk.LabelFrame(main_frame, text="▶️ 开始运行按钮", padding="3")
        run_frame.pack(fill=tk.X, pady=(0, 8))
        
        # 运行按钮居中
        self.run_button = ttk.Button(run_frame, text="🚀 开始审查", command=self._run_review, 
                                    width=25)
        self.run_button.pack(pady=2)
        
        # 状态栏居中
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(run_frame, textvariable=self.status_var, anchor=tk.CENTER)
        status_label.pack(fill=tk.X, pady=(2, 0))
        
        # 区域4：LOG日志窗口
        log_frame = ttk.LabelFrame(main_frame, text="📋 LOG日志窗口", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # 日志控制按钮
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(log_control_frame, text="清空日志", command=self._clear_log, width=12).pack(side=tk.LEFT)
        ttk.Button(log_control_frame, text="保存日志", command=self._save_log, width=12).pack(side=tk.LEFT, padx=(10, 0))
        
        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            wrap=tk.WORD, 
            font=self.log_font,
            height=20,
            width=80
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _create_review_settings(self, parent):
        """创建审查设置 - 清晰整齐的等宽布局"""
        # 创建容器Frame
        container_frame = ttk.Frame(parent)
        container_frame.pack(fill=tk.BOTH, expand=True, pady=8)
        
        # 配置grid列权重 - 确保等宽
        container_frame.grid_columnconfigure(0, weight=1)  # 左列权重1
        container_frame.grid_columnconfigure(1, weight=1)  # 右列权重1
        
        # 左列：LLM审查设置
        llm_review_frame = ttk.LabelFrame(container_frame, text="LLM审查", padding="8")
        llm_review_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        tk.Checkbutton(llm_review_frame, text="格式规范审查", variable=self.review_format, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        tk.Checkbutton(llm_review_frame, text="内容逻辑审查", variable=self.review_logic, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        tk.Checkbutton(llm_review_frame, text="缩略语审查", variable=self.review_acronyms, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        tk.Checkbutton(llm_review_frame, text="表达流畅性审查", variable=self.review_fluency, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        tk.Checkbutton(llm_review_frame, text="主题一致性检查", variable=self.theme_harmony, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        
        # 右列：审查规则设置
        rules_frame = ttk.LabelFrame(container_frame, text="规则审查", padding="8")
        rules_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        # 字体族检查 - 使用Frame包装实现整齐排列
        font_frame = ttk.Frame(rules_frame)
        font_frame.pack(fill=tk.X, pady=2)
        tk.Checkbutton(font_frame, text="字体族检查", variable=self.font_family, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(side=tk.LEFT)
        ttk.Label(font_frame, text="默认:").pack(side=tk.LEFT, padx=(10, 2))
        font_combo = ttk.Combobox(font_frame, textvariable=self.jp_font_name, 
                                 values=["Meiryo UI", "宋体", "微软雅黑", "楷体", "Time New Roman"], 
                                 state="readonly", width=12)
        font_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        # 字号检查
        size_frame = ttk.Frame(rules_frame)
        size_frame.pack(fill=tk.X, pady=2)
        tk.Checkbutton(size_frame, text="字号检查", variable=self.font_size, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(side=tk.LEFT)
        ttk.Label(size_frame, text="最小:").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Spinbox(size_frame, from_=8, to=72, textvariable=self.min_font_size_pt, width=6).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Label(size_frame, text="pt").pack(side=tk.LEFT, padx=(0, 5))
        
        # 颜色数量检查
        color_frame = ttk.Frame(rules_frame)
        color_frame.pack(fill=tk.X, pady=2)
        tk.Checkbutton(color_frame, text="颜色数量检查", variable=self.color_count, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(side=tk.LEFT)
        ttk.Label(color_frame, text="阈值:").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Spinbox(color_frame, from_=1, to=20, textvariable=self.color_count_threshold, width=6).pack(side=tk.LEFT, padx=(0, 5))
        
        # 缩略语解释检查
        tk.Checkbutton(rules_frame, text="缩略语解释检查", variable=self.acronym_explanation, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)

    def _select_ppt(self):
        """选择PPT文件"""
        filename = filedialog.askopenfilename(
            title="选择PPT文件",
            filetypes=[("PowerPoint文件", "*.pptx"), ("所有文件", "*.*")]
        )
        if filename:
            self.input_ppt.set(filename)
            # 自动设置输出目录：与输入文件同文件夹下的output文件夹，使用绝对路径
            input_dir = os.path.dirname(os.path.abspath(filename))  # 获取绝对路径
            base_name = os.path.splitext(os.path.basename(filename))[0]
            output_dir = os.path.join(input_dir, "output", f"{base_name}_{datetime.now().strftime('%Y%m%d')}")
            self.output_dir.set(output_dir)

    def _select_output_dir(self):
        """选择输出目录"""
        dirname = filedialog.askdirectory(title="选择输出目录")
        if dirname:
            # 确保使用绝对路径
            abs_dirname = os.path.abspath(dirname)
            self.output_dir.set(abs_dirname)

    def _on_provider_change(self, event=None):
        """提供商变更处理"""
        self._update_model_list()

    def _update_model_list(self):
        """更新模型列表"""
        provider = self.llm_provider.get()
        models = {
            "deepseek": ["deepseek-chat", "deepseek-coder"],
            "openai": ["gpt-4", "gpt-3.5-turbo", "gpt-4-turbo"],
            "anthropic": ["claude-3-opus", "claude-3-sonnet", "claude-3-haiku"],
            "local": ["qwen2.5-7b", "llama3.1-8b"]
        }
        
        if provider in models:
            self.model_combo['values'] = models[provider]
            if self.model_combo.get() not in models[provider]:
                self.model_combo.set(models[provider][0])

    def _apply_api_key(self):
        """应用新的API密钥"""
        new_api_key = self.llm_api_key.get().strip()
        if not new_api_key:
            messagebox.showerror("错误", "API密钥不能为空")
            return
        
        # 验证API密钥格式
        if not new_api_key.startswith(('sk-', 'Bearer ')):
            messagebox.showwarning("警告", "API密钥格式可能不正确，通常以'sk-'或'Bearer '开头")
        
        # 更新日志显示
        self._log(f"🔑 API密钥已更新: {new_api_key[:10]}...")
        self._log("✅ 新密钥将在下次运行时生效")
        
        # 显示成功消息
        messagebox.showinfo("成功", "API密钥已更新！\n新密钥将在下次运行时生效。")

    def _load_default_config(self):
        """加载默认配置"""
        # 设置默认API密钥
        self.llm_api_key.set("sk-55286a5c1f2a470081004104ec41af71")
        
        try:
            # 尝试加载配置文件，支持多种路径
            config_paths = [
                get_resource_path("configs/config.yaml"),
                "configs/config.yaml",
                "../configs/config.yaml",
                "app/configs/config.yaml"
            ]
            
            config_loaded = False
            for config_path in config_paths:
                if os.path.exists(config_path):
                    config = load_config(config_path)
                    # 加载LLM配置
                    if hasattr(config, 'llm_provider'):
                        self.llm_provider.set(config.llm_provider)
                    if hasattr(config, 'llm_model'):
                        self.llm_model.set(config.llm_model)
                    # 如果配置文件中有API密钥，则使用配置文件中的
                    if hasattr(config, 'llm_api_key') and config.llm_api_key:
                        self.llm_api_key.set(config.llm_api_key)
                    # 加载LLM启用状态
                    if hasattr(config, 'llm_enabled'):
                        self.llm_enabled.set(config.llm_enabled)
                    
                    # 加载审查设置
                    if hasattr(config, 'review_format'):
                        self.review_format.set(config.review_format)
                    if hasattr(config, 'review_logic'):
                        self.review_logic.set(config.review_logic)
                    if hasattr(config, 'review_acronyms'):
                        self.review_acronyms.set(config.review_acronyms)
                    if hasattr(config, 'review_fluency'):
                        self.review_fluency.set(config.review_fluency)
                    
                    # 加载审查规则设置
                    if hasattr(config, 'rules') and config.rules:
                        if 'font_family' in config.rules:
                            self.font_family.set(config.rules['font_family'])
                        if 'font_size' in config.rules:
                            self.font_size.set(config.rules['font_size'])
                        if 'color_count' in config.rules:
                            self.color_count.set(config.rules['color_count'])
                        if 'theme_harmony' in config.rules:
                            self.theme_harmony.set(config.rules['theme_harmony'])
                        if 'acronym_explanation' in config.rules:
                            self.acronym_explanation.set(config.rules['acronym_explanation'])
                    
                    # 加载审查规则配置值
                    if hasattr(config, 'jp_font_name'):
                        self.jp_font_name.set(config.jp_font_name)
                    if hasattr(config, 'min_font_size_pt'):
                        self.min_font_size_pt.set(config.min_font_size_pt)
                    if hasattr(config, 'color_count_threshold'):
                        self.color_count_threshold.set(config.color_count_threshold)
                    
                    self._update_model_list()
                    
                    # 记录配置加载成功
                    self._log(f"✅ 配置文件加载成功: {config_path}")
                    config_loaded = True
                    break
            
            if not config_loaded:
                self._log(f"⚠️ 配置文件不存在，尝试的路径: {config_paths}")
        except Exception as e:
            self._log(f"❌ 加载配置失败: {e}")
        
        # 启动时显示欢迎日志
        self._log("🚀 PPT审查工具已启动")
        self._log("📋 当前配置:")
        self._log(f"   - LLM提供商: {self.llm_provider.get()}")
        self._log(f"   - 模型: {self.llm_model.get()}")
        self._log(f"   - LLM启用: {'是' if self.llm_enabled.get() else '否'}")
        self._log(f"   - API密钥: {self.llm_api_key.get()[:10]}...")
        self._log("💡 请选择PPT文件开始审查")
        self._log("-" * 50)

    def _run_review(self):
        """运行审查"""
        # 验证输入
        input_ppt = os.path.abspath(self.input_ppt.get().strip())  # 确保是绝对路径
        output_dir = os.path.abspath(self.output_dir.get().strip())  # 确保是绝对路径
        
        if not input_ppt:
            messagebox.showerror("错误", "请选择PPT文件")
            return
        
        if not os.path.exists(input_ppt):
            messagebox.showerror("错误", f"PPT文件不存在: {input_ppt}")
            return
        
        if not output_dir:
            messagebox.showerror("错误", "请设置输出目录")
            return
        
        # 禁用运行按钮
        self.run_button.config(state=tk.DISABLED)
        self.status_var.set("运行中...")
        self._log("开始运行PPT审查...")
        
        # 在后台线程中运行
        def job():
            try:
                # 创建输出目录
                os.makedirs(output_dir, exist_ok=True)
                
                # 生成输出路径
                parsing_result_path, report_path, output_ppt_path = generate_output_paths(
                    input_ppt, self.mode.get(), output_dir
                )
                
                # 创建配置 - 从配置文件加载默认值，然后覆盖用户设置
                # 加载配置
                config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "configs", "config.yaml")
                cfg = load_config(config_file)
                
                # 应用用户设置的审查配置
                cfg.review_format = self.review_format.get()
                cfg.review_logic = self.review_logic.get()
                cfg.review_acronyms = self.review_acronyms.get()
                cfg.review_fluency = self.review_fluency.get()
                
                # 应用审查规则配置
                if not hasattr(cfg, 'rules'):
                    cfg.rules = {}
                cfg.rules['font_family'] = self.font_family.get()
                cfg.rules['font_size'] = self.font_size.get()
                cfg.rules['color_count'] = self.color_count.get()
                cfg.rules['theme_harmony'] = self.theme_harmony.get()
                cfg.rules['acronym_explanation'] = self.acronym_explanation.get()
                
                # 应用审查规则配置值
                cfg.jp_font_name = self.jp_font_name.get()
                cfg.min_font_size_pt = self.min_font_size_pt.get()
                cfg.color_count_threshold = self.color_count_threshold.get()
                
                # 解析PPT
                self._log("步骤1: 解析PPT文件...")
                parsing_data = parse_pptx(input_ppt, include_images=False)
                
                # 保存解析结果
                import json
                with open(parsing_result_path, "w", encoding="utf-8") as f:
                    json.dump(parsing_data, f, ensure_ascii=False, indent=2)
                self._log(f"✅ PPT解析完成")
                
                # 创建LLM客户端
                llm = LLMClient(
                    provider=getattr(cfg, 'llm_provider', 'deepseek'),
                    api_key=getattr(cfg, 'llm_api_key', None),
                    endpoint=getattr(cfg, 'llm_endpoint', None),
                    model=getattr(cfg, 'llm_model', 'deepseek-chat'),
                    temperature=getattr(cfg, 'llm_temperature', 0.2),
                    max_tokens=getattr(cfg, 'llm_max_tokens', 9999)
                )
                self._log(f"✅ LLM客户端创建成功: {getattr(cfg, 'llm_provider', 'deepseek')}/{getattr(cfg, 'llm_model', 'deepseek-chat')}")

                
                # 运行审查 - 使用控制台捕获器
                self._log("步骤2: 开始审查...")
                try:
                    with ConsoleCapture(self._log):
                        res = run_review_workflow(parsing_result_path, cfg, output_ppt_path, llm, input_ppt)
                except Exception as workflow_error:
                    self._log(f"⚠️ 控制台捕获模式失败，使用标准模式: {workflow_error}")
                    # 降级到标准模式，不使用控制台捕获
                    try:
                        res = run_review_workflow(parsing_result_path, cfg, output_ppt_path, llm, input_ppt)
                    except Exception as std_error:
                        self._log(f"❌ 标准模式也失败: {std_error}")
                        # 创建空的审查结果
                        class EmptyResult:
                            def __init__(self):
                                self.issues = []
                                self.report_md = "# PPT审查报告\n\n## ❌ 审查过程失败\n\n由于技术问题，无法完成自动审查。\n\n### 错误信息\n```\n{std_error}\n```\n\n### 建议\n1. 检查网络连接\n2. 确认API密钥有效\n3. 尝试重新运行\n"
                        res = EmptyResult()
                
                # 生成报告
                if hasattr(res, 'report_md') and res.report_md:
                    with open(report_path, "w", encoding="utf-8") as f:
                        f.write(res.report_md)
                    self._log(f"✅ 报告已生成")
                
                
                # 显示结果
                total_issues = len(getattr(res, 'issues', []))
                self._log(f"🎯 审查完成！发现 {total_issues} 个问题")
                self.status_var.set(f"完成：{total_issues} 个问题")
                
                # 显示成功对话框
                self.after(0, lambda: self._show_success_dialog(output_dir, report_path, output_ppt_path))
                
            except Exception as e:
                error_msg = f"运行失败: {e}"
                self._log(f"❌ {error_msg}")
                self.status_var.set("运行失败")
                messagebox.showerror("运行失败", str(e))
            finally:
                self.run_button.config(state=tk.NORMAL)

        # 启动后台线程，设置daemon=True避免黑框显示
        thread = threading.Thread(target=job, daemon=True)
        thread.start()

    def _show_success_dialog(self, output_dir: str, report_path: str, ppt_path: str):
        """显示成功对话框"""
        message = f"""✅ PPT审查完成！

📁 结果保存位置：
   {output_dir}

📄 生成的文件：
   • 审查报告：{os.path.basename(report_path)}
   • 标记PPT：{os.path.basename(ppt_path)}
   • 解析结果：parsing_result.json

💡 提示：
   • 可以在输出目录中查看详细的审查报告
   • 标记PPT中已标注了发现的问题
   • 建议根据报告中的建议进行PPT优化

是否打开输出目录？"""
        
        if messagebox.askyesno("审查完成", message):
            try:
                import subprocess
                import platform
                if platform.system() == "Windows":
                    subprocess.run(["explorer", output_dir])
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", output_dir])
                else:  # Linux
                    subprocess.run(["xdg-open", output_dir])
            except Exception as e:
                print(f"无法打开目录: {e}")

    def _log(self, message):
        """添加日志消息"""
        # 如果消息以换行符结尾，则移除它（因为print会自动添加）
        if message.endswith('\n'):
            message = message[:-1]
        
        # 插入消息并换行
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def _clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)

    def _save_log(self):
        """保存日志"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(self.log_text.get(1.0, tk.END))
                messagebox.showinfo("保存成功", f"日志已保存到 {filename}")
            except Exception as e:
                messagebox.showerror("保存失败", f"保存日志失败: {e}")


def main():
    """主函数"""
    app = SimpleApp()
    app.mainloop()


if __name__ == "__main__":
    main()
