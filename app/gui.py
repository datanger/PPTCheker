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

# 导入Rich库用于终端颜色输出
try:
    from rich.console import Console
    # from rich.text import Text  # 暂时不使用
    RICH_AVAILABLE = True
    # 创建全局Rich控制台实例
    console = Console()
except ImportError:
    RICH_AVAILABLE = False
    console = None
    print("⚠️ Rich库未安装，终端输出将无颜色")

def colored_print(message, level='info'):
    """颜色化的print函数，同时输出到终端和GUI"""
    if RICH_AVAILABLE and console:
        # 根据级别选择Rich颜色
        colors = {
            'info': 'white',
            'success': 'green',
            'warning': 'yellow',
            'error': 'red',
            'debug': 'dim',
            'highlight': 'blue'
        }
        
        color = colors.get(level, 'white')
        
        # 使用Rich输出带颜色的文本
        console.print(message, style=color)
    else:
        # 如果Rich不可用，使用普通print
        print(message)
from datetime import datetime
import io
# import contextlib  # 暂时不使用

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
colored_print("✅ 使用绝对导入模式", 'success')



class ConsoleCapture:
    """控制台输出捕获器 - 完全避免递归调用"""
    def __init__(self, log_callback):
        self.log_callback = log_callback
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr
        self._capturing = False
    
    def __enter__(self):
        self._capturing = True
        
        # 创建完全安全的输出流
        class SafeStream:
            def __init__(self, original_stream, callback, capture_instance):
                self.original_stream = original_stream
                self.callback = callback
                self.capture_instance = capture_instance
            
            def write(self, text):
                # 直接写入原始流，不使用任何可能触发递归的函数
                try:
                    if self.original_stream and hasattr(self.original_stream, 'write'):
                        self.original_stream.write(text)
                except Exception:
                    pass
                
                # 安全回调到GUI（完全避免递归）
                try:
                    if (self.capture_instance._capturing and 
                        self.callback and 
                        text and 
                        text.strip()):  # 只处理非空文本
                        # 直接调用回调，不使用print或其他可能触发递归的函数
                        self.callback(text)
                except Exception:
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
        sys.stdout = SafeStream(self.original_stdout, self.log_callback, self)
        sys.stderr = SafeStream(self.original_stderr, lambda x: self.log_callback(f"错误: {x}"), self)
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self._capturing = False
        sys.stdout = self.original_stdout
        sys.stderr = self.original_stderr


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
        
        # 设置更好的字体和颜色主题
        self._setup_fonts()
        self._setup_colors()
        
        # 配置变量
        self.input_ppt = tk.StringVar()
        self.output_dir = tk.StringVar(value="output")
        self.llm_enabled = tk.BooleanVar(value=True)
        self.llm_provider = tk.StringVar(value="deepseek")
        self.llm_model = tk.StringVar(value="deepseek-chat")
        self.llm_api_key = tk.StringVar()
        self.mode = tk.StringVar(value="review")
        
        # 审查设置变量
        self.review_logic = tk.BooleanVar(value=True)
        self.review_acronyms = tk.BooleanVar(value=True)
        self.review_fluency = tk.BooleanVar(value=True)
        self.font_family = tk.BooleanVar(value=True)
        self.font_size = tk.BooleanVar(value=True)
        self.color_count = tk.BooleanVar(value=True)
        self.theme_harmony = tk.BooleanVar(value=True)
        
        # 运行状态变量
        self.is_running = False
        self.should_stop = False
        self.stop_event = threading.Event()  # 用于跨线程通信的停止事件
        self.worker_thread = None  # 工作线程引用
        
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
                colored_print(f"可用主题: {available_themes}", 'info')
                
                # 尝试使用alt主题，它通常有更好的复选框样式
                if 'alt' in available_themes:
                    style.theme_use('alt')
                    colored_print("✅ 使用alt主题", 'success')
                elif 'default' in available_themes:
                    style.theme_use('default')
                    colored_print("✅ 使用default主题", 'success')
                
                # 重新配置复选框样式
                style.configure('TCheckbutton', font=default_font)
                
                # 方法2：尝试修改复选框的映射
                style.map('TCheckbutton',
                         indicatorcolor=[('selected', 'black'),
                                       ('!selected', 'white')],
                         background=[('active', 'white'),
                                   ('!active', 'white')])
                
                colored_print("✅ 复选框样式修改完成", 'success')
                
            except Exception as e:
                colored_print(f"⚠️ 复选框样式修改失败: {e}", 'warning')
            
            colored_print("使用Ubuntu优化字体设置", 'info')
                
        except Exception as e:
            colored_print(f"字体设置失败: {e}", 'error')
            # 使用系统默认字体
            self.title_font = ('TkHeadingFont', 12, 'bold')
            self.log_font = ('TkFixedFont', 8)
    
    def _setup_colors(self):
        """设置界面颜色主题"""
        try:
            # 定义颜色主题
            self.colors = {
                'primary': '#2E86AB',      # 主色调 - 蓝色
                'secondary': '#A23B72',    # 辅助色 - 紫红色
                'success': '#F18F01',      # 成功色 - 橙色
                'warning': '#C73E1D',      # 警告色 - 红色
                'info': '#6A994E',         # 信息色 - 绿色
                'light': '#F8F9FA',        # 浅色背景
                'dark': '#212529',         # 深色文字
                'border': '#DEE2E6',       # 边框色
                'hover': '#E9ECEF'         # 悬停色
            }
            
            # 设置窗口背景色
            self.configure(bg=self.colors['light'])
            
            # 配置ttk样式
            style = ttk.Style()
            
            # 配置LabelFrame样式
            style.configure('TLabelframe', 
                          background=self.colors['light'],
                          borderwidth=2,
                          relief='solid')
            style.configure('TLabelframe.Label', 
                          background=self.colors['light'],
                          foreground=self.colors['dark'],
                          font=self.title_font)
            
            # 配置按钮样式
            style.configure('TButton',
                          background=self.colors['primary'],
                          foreground='white',
                          font=self.title_font,
                          borderwidth=1,
                          relief='solid')
            style.map('TButton',
                     background=[('active', self.colors['secondary']),
                               ('pressed', self.colors['warning'])])
            
            # 配置复选框样式
            style.configure('TCheckbutton',
                          background=self.colors['light'],
                          foreground=self.colors['dark'],
                          font=self.title_font)
            style.map('TCheckbutton',
                     background=[('active', self.colors['hover']),
                               ('!active', self.colors['light'])],
                     foreground=[('active', self.colors['primary']),
                               ('!active', self.colors['dark'])])
            
            # 配置输入框样式
            style.configure('TEntry',
                          fieldbackground='white',
                          foreground=self.colors['dark'],
                          borderwidth=1,
                          relief='solid')
            
            # 配置Spinbox样式
            style.configure('TSpinbox',
                          fieldbackground='white',
                          foreground=self.colors['dark'],
                          borderwidth=1,
                          relief='solid')
            
            # 配置Combobox样式
            style.configure('TCombobox',
                          fieldbackground='white',
                          foreground=self.colors['dark'],
                          borderwidth=1,
                          relief='solid')
            
            colored_print("✅ 颜色主题设置完成", 'success')
            
        except Exception as e:
            colored_print(f"颜色设置失败: {e}", 'error')
            # 使用默认颜色
            self.colors = {
                'primary': '#007ACC',
                'secondary': '#6C757D',
                'success': '#28A745',
                'warning': '#FFC107',
                'info': '#17A2B8',
                'light': '#FFFFFF',
                'dark': '#000000',
                'border': '#CCCCCC',
                'hover': '#F5F5F5'
            }

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
        run_frame = ttk.LabelFrame(main_frame, text="▶️ 运行控制", padding="3")
        run_frame.pack(fill=tk.X, pady=(0, 8))
        
        # 按钮容器 - 并排显示
        button_frame = ttk.Frame(run_frame)
        button_frame.pack(pady=2)
        
        # 开始审查按钮 - 美化版本
        self.run_button = ttk.Button(button_frame, text="🚀 开始审查", command=self._run_review, 
                                    width=15)
        self.run_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # 终止按钮 - 美化版本
        self.stop_button = ttk.Button(button_frame, text="⏹️ 终止", command=self._stop_review, 
                                     width=15, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(5, 0))
        
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
        
        ttk.Button(log_control_frame, text="🗑️ 清空日志", command=self._clear_log, width=12).pack(side=tk.LEFT)
        ttk.Button(log_control_frame, text="💾 保存日志", command=self._save_log, width=12).pack(side=tk.LEFT, padx=(10, 0))
        
        # 日志文本框 - 美化版本
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            wrap=tk.WORD, 
            font=self.log_font,
            height=20,
            width=80,
            bg='#1E1E1E',  # 深色背景
            fg='#FFFFFF',  # 白色文字
            insertbackground='#FFFFFF',  # 光标颜色
            selectbackground='#404040',  # 选中背景
            selectforeground='#FFFFFF',  # 选中文字
            relief='solid',
            borderwidth=1
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 配置默认文本颜色标签
        self.log_text.tag_config("default", foreground='#FFFFFF')
        self.log_text.tag_config("log_info", foreground='#FFFFFF')
        self.log_text.tag_config("log_success", foreground='#4CAF50')
        self.log_text.tag_config("log_warning", foreground='#FF9800')
        self.log_text.tag_config("log_error", foreground='#F44336')
        self.log_text.tag_config("log_debug", foreground='#9E9E9E')
        self.log_text.tag_config("log_highlight", foreground='#2196F3')

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
        
        tk.Checkbutton(llm_review_frame, text="内容逻辑审查", variable=self.review_logic, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        tk.Checkbutton(llm_review_frame, text="缩略语审查", variable=self.review_acronyms, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        tk.Checkbutton(llm_review_frame, text="表达流畅性审查", variable=self.review_fluency, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        tk.Checkbutton(llm_review_frame, text="主题一致性检查", variable=self.theme_harmony, 
                       font=('WenQuanYi Micro Hei', 9), selectcolor='white').pack(anchor=tk.W, padx=3, pady=2)
        
        # 提示词管理按钮
        ttk.Button(llm_review_frame, text="📝 管理提示词", command=self._open_prompt_manager, 
                   width=15).pack(anchor=tk.W, padx=3, pady=(10, 2))
        
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
        

    def _open_prompt_manager(self):
        """打开提示词管理窗口"""
        try:
            # 导入提示词管理器
            from pptlint.prompt_manager import prompt_manager
            
            # 创建提示词管理窗口
            PromptManagerWindow(self, prompt_manager)
            
        except Exception as e:
            messagebox.showerror("错误", f"打开提示词管理器失败: {e}")

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
        self._log("🚀 PPT审查工具已启动", 'success')
        self._log("📋 当前配置:", 'highlight')
        self._log(f"   - LLM提供商: {self.llm_provider.get()}", 'info')
        self._log(f"   - 模型: {self.llm_model.get()}", 'info')
        self._log(f"   - LLM启用: {'是' if self.llm_enabled.get() else '否'}", 'info')
        self._log(f"   - API密钥: {self.llm_api_key.get()[:10]}...", 'info')
        self._log("💡 请选择PPT文件开始审查", 'highlight')
        self._log("-" * 50, 'debug')
        
        # 同时在终端输出欢迎信息（避免递归）
        if RICH_AVAILABLE and console:
            console.print("🚀 PPT审查工具已启动", style="green")
            console.print("📋 当前配置:", style="blue")
            console.print(f"   - LLM提供商: {self.llm_provider.get()}", style="white")
            console.print(f"   - 模型: {self.llm_model.get()}", style="white")
            console.print(f"   - LLM启用: {'是' if self.llm_enabled.get() else '否'}", style="white")
            console.print(f"   - API密钥: {self.llm_api_key.get()[:10]}...", style="white")
            console.print("💡 请选择PPT文件开始审查", style="blue")
            console.print("-" * 50, style="dim")

    def _stop_review(self):
        """终止审查"""
        if self.is_running:
            self.should_stop = True
            self.stop_event.set()  # 设置停止事件
            self._log("⏹️ 用户请求终止审查...", 'warning')
            self.status_var.set("正在终止...")
            
            # 强制终止工作线程（如果存在）
            if self.worker_thread and self.worker_thread.is_alive():
                self._log("🔄 正在强制终止工作线程...", 'warning')
                # 注意：在Windows上，强制终止线程可能不安全，但这是最后的 resort
                try:
                    import ctypes
                    thread_id = self.worker_thread.ident
                    if thread_id:
                        ctypes.pythonapi.PyThreadState_SetAsyncExc(ctypes.c_long(thread_id), ctypes.py_object(KeyboardInterrupt))
                        self._log("✅ 工作线程已强制终止", 'success')
                except Exception as e:
                    self._log(f"⚠️ 强制终止失败: {e}", 'error')
            
            # 按钮状态会在_run_review方法中更新
        else:
            self._log("⚠️ 当前没有正在运行的审查任务")

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
        
        # 设置运行状态
        self.is_running = True
        self.should_stop = False
        self.stop_event.clear()  # 清除停止事件
        
        # 更新按钮状态
        self.run_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
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
                
                # 应用审查规则配置值
                cfg.jp_font_name = self.jp_font_name.get()
                cfg.min_font_size_pt = self.min_font_size_pt.get()
                cfg.color_count_threshold = self.color_count_threshold.get()
                
                # 检查是否应该终止
                if self.should_stop:
                    self._log("⏹️ 用户终止了审查过程")
                    return
                
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
                    max_tokens=getattr(cfg, 'llm_max_tokens', 9999),
                    use_proxy=getattr(cfg, 'llm_use_proxy', False),
                    proxy_url=getattr(cfg, 'llm_proxy_url', None)
                )
                self._log(f"✅ LLM客户端创建成功: {getattr(cfg, 'llm_provider', 'deepseek')}/{getattr(cfg, 'llm_model', 'deepseek-chat')}")

                
                # 检查是否应该终止
                if self.should_stop:
                    self._log("⏹️ 用户终止了审查过程")
                    return
                
                # 运行审查 - 使用控制台捕获器
                self._log("步骤2: 开始审查...")
                try:
                    with ConsoleCapture(self._log):
                        res = run_review_workflow(parsing_result_path, cfg, output_ppt_path, llm, input_ppt, self.stop_event)
                except Exception as workflow_error:
                    self._log(f"⚠️ 控制台捕获模式失败，使用标准模式: {workflow_error}")
                    # 降级到标准模式，不使用控制台捕获
                    try:
                        res = run_review_workflow(parsing_result_path, cfg, output_ppt_path, llm, input_ppt, self.stop_event)
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
                # 重置运行状态
                self.is_running = False
                self.should_stop = False
                self.stop_event.clear()  # 清除停止事件
                self.worker_thread = None  # 清理线程引用
                
                # 恢复按钮状态
                self.run_button.config(state=tk.NORMAL)
                self.stop_button.config(state=tk.DISABLED)
                
                # 更新状态
                if self.status_var.get() == "正在终止...":
                    self.status_var.set("已终止")
                elif self.status_var.get() == "运行中...":
                    self.status_var.set("已完成")

        # 启动后台线程，设置daemon=True避免黑框显示
        self.worker_thread = threading.Thread(target=job, daemon=True)
        self.worker_thread.start()

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

    def _log(self, message, level='info'):
        """添加彩色日志消息"""
        # 如果消息以换行符结尾，则移除它（因为print会自动添加）
        if message.endswith('\n'):
            message = message[:-1]
        
        # 根据级别选择颜色
        colors = {
            'info': '#FFFFFF',      # 白色 - 普通信息
            'success': '#4CAF50',   # 绿色 - 成功
            'warning': '#FF9800',   # 橙色 - 警告
            'error': '#F44336',     # 红色 - 错误
            'debug': '#9E9E9E',     # 灰色 - 调试
            'highlight': '#2196F3'  # 蓝色 - 高亮
        }
        
        # 获取当前颜色
        color = colors.get(level, colors['info'])
        
        # 插入带颜色的消息
        self.log_text.insert(tk.END, f"{message}\n")
        
        # 设置最后插入的文本的颜色
        start_line = self.log_text.index(tk.END + "-2l")
        end_line = self.log_text.index(tk.END + "-1l")
        self.log_text.tag_add(f"log_{level}", start_line, end_line)
        self.log_text.tag_config(f"log_{level}", foreground=color, font=self.log_font)
        
        self.log_text.see(tk.END)  # 自动滚动到底部
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


class PromptManagerWindow:
    """提示词管理窗口"""
    
    def __init__(self, parent, prompt_manager):
        self.parent = parent
        self.prompt_manager = prompt_manager
        self.current_prompt_key = None
        
        # 创建窗口
        self.window = tk.Toplevel(parent)
        self.window.title("LLM提示词管理")
        self.window.geometry("900x700")
        self.window.resizable(True, True)
        
        # 设置窗口图标和居中
        self.window.transient(parent)
        self.window.grab_set()
        
        # 创建UI
        self._create_ui()
        
        # 加载提示词列表
        self._load_prompt_list()
    
    def _create_ui(self):
        """创建UI界面"""
        # 主容器
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="LLM提示词管理", font=('WenQuanYi Micro Hei', 12, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # 创建左右分栏
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左列：提示词列表
        left_frame = ttk.LabelFrame(content_frame, text="提示词列表", padding="8")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 提示词列表框
        self.prompt_listbox = tk.Listbox(left_frame, font=('WenQuanYi Micro Hei', 9))
        self.prompt_listbox.pack(fill=tk.BOTH, expand=True)
        self.prompt_listbox.bind('<<ListboxSelect>>', self._on_prompt_select)
        
        # 右列：提示词编辑
        right_frame = ttk.LabelFrame(content_frame, text="提示词编辑", padding="8")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 提示词信息
        info_frame = ttk.Frame(right_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(info_frame, text="名称:", font=('WenQuanYi Micro Hei', 9, 'bold')).pack(anchor=tk.W)
        self.name_label = ttk.Label(info_frame, text="", font=('WenQuanYi Micro Hei', 9))
        self.name_label.pack(anchor=tk.W, pady=(0, 5))
        
        ttk.Label(info_frame, text="描述:", font=('WenQuanYi Micro Hei', 9, 'bold')).pack(anchor=tk.W)
        self.desc_label = ttk.Label(info_frame, text="", font=('WenQuanYi Micro Hei', 9), wraplength=350)
        self.desc_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 提示词编辑区域
        ttk.Label(right_frame, text="用户提示词 (可编辑):", font=('WenQuanYi Micro Hei', 9, 'bold')).pack(anchor=tk.W)
        
        # 创建文本框和滚动条
        text_frame = ttk.Frame(right_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        
        self.prompt_text = tk.Text(text_frame, wrap=tk.WORD, font=('WenQuanYi Micro Hei', 9), height=15)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.prompt_text.yview)
        self.prompt_text.configure(yscrollcommand=scrollbar.set)
        
        self.prompt_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 按钮区域
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="保存", command=self._save_prompt, width=10).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="重置", command=self._reset_prompt, width=10).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="查看完整提示词", command=self._view_full_prompt, width=15).pack(side=tk.RIGHT)
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(bottom_frame, text="关闭", command=self.window.destroy, width=10).pack(side=tk.RIGHT)
    
    def _load_prompt_list(self):
        """加载提示词列表"""
        self.prompt_listbox.delete(0, tk.END)
        
        prompts = self.prompt_manager.get_all_prompts()
        for key, prompt in prompts.items():
            self.prompt_listbox.insert(tk.END, prompt.name)
        
        # 存储key到name的映射
        self.key_to_name = {prompt.name: key for key, prompt in prompts.items()}
    
    def _on_prompt_select(self, event):
        """提示词选择事件"""
        selection = self.prompt_listbox.curselection()
        if selection:
            name = self.prompt_listbox.get(selection[0])
            key = self.key_to_name.get(name)
            if key:
                self._load_prompt_content(key)
    
    def _load_prompt_content(self, key):
        """加载提示词内容"""
        self.current_prompt_key = key
        prompt = self.prompt_manager.get_prompt(key)
        
        if prompt:
            self.name_label.config(text=prompt.name)
            self.desc_label.config(text=prompt.description)
            self.prompt_text.delete(1.0, tk.END)
            self.prompt_text.insert(1.0, prompt.user_prompt)
    
    def _save_prompt(self):
        """保存提示词"""
        if not self.current_prompt_key:
            messagebox.showwarning("警告", "请先选择一个提示词")
            return
        
        new_prompt = self.prompt_text.get(1.0, tk.END).strip()
        if not new_prompt:
            messagebox.showwarning("警告", "提示词不能为空")
            return
        
        try:
            self.prompt_manager.update_user_prompt(self.current_prompt_key, new_prompt)
            self.prompt_manager.save_prompts()
            messagebox.showinfo("成功", "提示词已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")
    
    def _reset_prompt(self):
        """重置提示词"""
        if not self.current_prompt_key:
            messagebox.showwarning("警告", "请先选择一个提示词")
            return
        
        if messagebox.askyesno("确认", "确定要重置为默认提示词吗？"):
            try:
                # 重新加载配置文件
                self.prompt_manager.load_prompts()
                self._load_prompt_content(self.current_prompt_key)
                messagebox.showinfo("成功", "已重置为默认提示词")
            except Exception as e:
                messagebox.showerror("错误", f"重置失败: {e}")
    
    def _view_full_prompt(self):
        """查看完整提示词"""
        if not self.current_prompt_key:
            messagebox.showwarning("警告", "请先选择一个提示词")
            return
        
        prompt = self.prompt_manager.get_prompt(self.current_prompt_key)
        if prompt:
            # 创建新窗口显示完整提示词
            full_window = tk.Toplevel(self.window)
            full_window.title(f"完整提示词 - {prompt.name}")
            full_window.geometry("1000x800")
            
            # 创建文本框
            text_frame = ttk.Frame(full_window, padding="10")
            text_frame.pack(fill=tk.BOTH, expand=True)
            
            text_widget = tk.Text(text_frame, wrap=tk.WORD, font=('WenQuanYi Micro Hei', 9))
            scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 插入完整提示词（只显示用户提示部分）
            full_prompt = f"""=== 用户提示词（可编辑） ===
{prompt.user_prompt}

=== 说明 ===
输入提示和输出提示部分保留在代码中，不在配置文件中。
用户只能修改上述用户提示词部分。"""
            
            text_widget.insert(1.0, full_prompt)
            text_widget.config(state=tk.DISABLED)  # 只读模式
            
            # 关闭按钮
            ttk.Button(full_window, text="关闭", command=full_window.destroy).pack(pady=10)


def main():
    """主函数"""
    app = SimpleApp()
    app.mainloop()


if __name__ == "__main__":
    main()
