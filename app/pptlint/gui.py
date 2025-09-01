"""
完整GUI：运行参数配置与工作流执行（支持所有config配置和CLI参数）。

功能：
- 选择输入PPT文件、输出目录、配置文件
- 选择运行模式（review/edit）
- 配置所有字体、颜色、LLM、审查维度、审查规则等参数
- 支持配置文件的加载、保存和覆盖
- 运行并在状态栏显示结果
"""
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import yaml

from .config import load_config, ToolConfig
from .workflow import run_review_workflow, run_edit_workflow
from .llm import LLMClient


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PPT 审查工具 - 完整配置界面")
        self.geometry("1000x800")
        
        # 配置变量
        self.config_vars = {}
        self._init_config_vars()
        
        # 当前配置
        self.current_config = None

        self._build_ui()
        self._load_default_config()

    def _init_config_vars(self):
        """初始化所有配置变量"""
        # 基本文件路径
        self.config_vars['input_ppt'] = tk.StringVar()
        self.config_vars['output_dir'] = tk.StringVar(value="output")
        self.config_vars['config_file'] = tk.StringVar(value="configs/config.yaml")
        self.config_vars['mode'] = tk.StringVar(value="review")
        self.config_vars['edit_req'] = tk.StringVar(value="请分析PPT内容，提供改进建议")
        
        # 字体配置
        self.config_vars['jp_font_name'] = tk.StringVar(value="Meiryo UI")
        self.config_vars['min_font_size_pt'] = tk.IntVar(value=12)
        
        # 颜色配置
        self.config_vars['color_count_threshold'] = tk.IntVar(value=5)
        
        # 输出格式
        self.config_vars['output_format'] = tk.StringVar(value="md")
        
        # 自动修复开关
        self.config_vars['autofix_font'] = tk.BooleanVar(value=False)
        self.config_vars['autofix_size'] = tk.BooleanVar(value=False)
        self.config_vars['autofix_color'] = tk.BooleanVar(value=False)
        
        # 词库路径
        self.config_vars['jp_terms_path'] = tk.StringVar(value="dicts/jp_it_terms.txt")
        self.config_vars['term_mapping_path'] = tk.StringVar(value="dicts/term_mapping.csv")
        
        # LLM配置
        self.config_vars['llm_enabled'] = tk.BooleanVar(value=True)
        self.config_vars['llm_model'] = tk.StringVar(value="deepseek-chat")
        self.config_vars['llm_temperature'] = tk.DoubleVar(value=0.2)
        self.config_vars['llm_max_tokens'] = tk.IntVar(value=99999)
        
        # 审查维度开关
        self.config_vars['review_format'] = tk.BooleanVar(value=False)
        self.config_vars['review_logic'] = tk.BooleanVar(value=True)
        self.config_vars['review_acronyms'] = tk.BooleanVar(value=False)
        self.config_vars['review_fluency'] = tk.BooleanVar(value=True)
        
        # 审查规则配置
        self.config_vars['font_family'] = tk.BooleanVar(value=True)
        self.config_vars['font_size'] = tk.BooleanVar(value=True)
        self.config_vars['color_count'] = tk.BooleanVar(value=True)
        self.config_vars['theme_harmony'] = tk.BooleanVar(value=True)
        self.config_vars['acronym_explanation'] = tk.BooleanVar(value=True)
        
        # 报告配置
        self.config_vars['include_summary'] = tk.BooleanVar(value=True)
        self.config_vars['include_details'] = tk.BooleanVar(value=True)
        self.config_vars['include_suggestions'] = tk.BooleanVar(value=True)
        self.config_vars['include_statistics'] = tk.BooleanVar(value=True)
        
        # LLM控制参数
        self.config_vars['llm_override'] = tk.StringVar(value="")
        
        # 高级配置参数
        self.config_vars['font_size_override'] = tk.StringVar()
        self.config_vars['color_threshold_override'] = tk.StringVar()

    def _build_ui(self):
        """构建UI界面"""
        # 创建主框架
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建notebook用于分页显示
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # 基本设置页面
        basic_frame = self._create_basic_frame()
        notebook.add(basic_frame, text="基本设置")
        
        # 字体和颜色配置页面
        font_frame = self._create_font_frame()
        notebook.add(font_frame, text="字体和颜色")
        
        # LLM配置页面
        llm_frame = self._create_llm_frame()
        notebook.add(llm_frame, text="LLM配置")
        
        # 审查配置页面
        review_frame = self._create_review_frame()
        notebook.add(review_frame, text="审查配置")
        
        # 报告配置页面
        report_frame = self._create_report_frame()
        notebook.add(report_frame, text="报告配置")
        
        # 高级配置页面
        advanced_frame = self._create_advanced_frame()
        notebook.add(advanced_frame, text="高级配置")
        
        # 底部按钮和状态栏
        self._create_bottom_frame(main_frame)

    def _create_basic_frame(self):
        """创建基本设置页面"""
        frame = ttk.Frame()
        
        # 文件选择区域
        file_group = ttk.LabelFrame(frame, text="文件选择", padding=10)
        file_group.pack(fill=tk.X, padx=5, pady=5)
        
        self._create_file_row(file_group, "输入PPT文件:", 'input_ppt', False, [("PPTX文件", "*.pptx")])
        self._create_file_row(file_group, "输出目录:", 'output_dir', True)
        self._create_file_row(file_group, "配置文件:", 'config_file', False, [("YAML文件", "*.yaml"), ("所有文件", "*.*")])
        
        # 运行模式区域
        mode_group = ttk.LabelFrame(frame, text="运行模式", padding=10)
        mode_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(mode_group, text="模式:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        mode_combo = ttk.Combobox(mode_group, textvariable=self.config_vars['mode'], 
                                 values=["review", "edit"], state="readonly", width=20)
        mode_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(mode_group, text="编辑要求:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        edit_req_entry = ttk.Entry(mode_group, textvariable=self.config_vars['edit_req'], width=50)
        edit_req_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 配置操作区域
        config_group = ttk.LabelFrame(frame, text="配置操作", padding=10)
        config_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(config_group, text="加载配置文件", command=self._load_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(config_group, text="保存配置文件", command=self._save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(config_group, text="重置为默认", command=self._reset_config).pack(side=tk.LEFT, padx=5)
        
        return frame

    def _create_font_frame(self):
        """创建字体和颜色配置页面"""
        frame = ttk.Frame()
        
        # 字体配置
        font_group = ttk.LabelFrame(frame, text="字体配置", padding=10)
        font_group.pack(fill=tk.X, padx=5, pady=5)
        
        self._create_config_row(font_group, "日文字体:", 'jp_font_name', 0)
        self._create_config_row(font_group, "最小字号(pt):", 'min_font_size_pt', 1, int)
        
        # 颜色配置
        color_group = ttk.LabelFrame(frame, text="颜色配置", padding=10)
        color_group.pack(fill=tk.X, padx=5, pady=5)
        
        self._create_config_row(color_group, "颜色数量阈值:", 'color_count_threshold', 0, int)
        
        # 输出格式
        format_group = ttk.LabelFrame(frame, text="输出格式", padding=10)
        format_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(format_group, text="报告格式:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        format_combo = ttk.Combobox(format_group, textvariable=self.config_vars['output_format'], 
                                   values=["md", "html"], state="readonly", width=20)
        format_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 自动修复开关
        autofix_group = ttk.LabelFrame(frame, text="自动修复开关", padding=10)
        autofix_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(autofix_group, text="自动修复字体", variable=self.config_vars['autofix_font']).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(autofix_group, text="自动修复字号", variable=self.config_vars['autofix_size']).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(autofix_group, text="自动修复颜色", variable=self.config_vars['autofix_color']).pack(side=tk.LEFT, padx=5)
        
        return frame

    def _create_llm_frame(self):
        """创建LLM配置页面"""
        frame = ttk.Frame()
        
        # LLM基本配置
        llm_basic_group = ttk.LabelFrame(frame, text="LLM基本配置", padding=10)
        llm_basic_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(llm_basic_group, text="启用LLM审查", variable=self.config_vars['llm_enabled']).pack(anchor=tk.W, padx=5, pady=2)
        
        self._create_config_row(llm_basic_group, "LLM模型:", 'llm_model', 0)
        self._create_config_row(llm_basic_group, "温度参数:", 'llm_temperature', 1, float)
        self._create_config_row(llm_basic_group, "最大Token数:", 'llm_max_tokens', 2, int)
        
        # LLM控制参数
        llm_control_group = ttk.LabelFrame(frame, text="LLM控制参数", padding=10)
        llm_control_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(llm_control_group, text="LLM覆盖:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        llm_override_combo = ttk.Combobox(llm_control_group, textvariable=self.config_vars['llm_override'], 
                                         values=["", "on", "off"], state="readonly", width=20)
        llm_override_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(llm_control_group, text="(空值表示使用配置文件设置)").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        
        return frame

    def _create_review_frame(self):
        """创建审查配置页面"""
        frame = ttk.Frame()
        
        # 审查维度开关
        review_dims_group = ttk.LabelFrame(frame, text="审查维度开关", padding=10)
        review_dims_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(review_dims_group, text="格式规范审查", variable=self.config_vars['review_format']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(review_dims_group, text="内容逻辑审查", variable=self.config_vars['review_logic']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(review_dims_group, text="缩略语审查", variable=self.config_vars['review_acronyms']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(review_dims_group, text="表达流畅性审查", variable=self.config_vars['review_fluency']).pack(anchor=tk.W, padx=5, pady=2)
        
        # 审查规则配置
        rules_group = ttk.LabelFrame(frame, text="审查规则配置", padding=10)
        rules_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(rules_group, text="字体族检查", variable=self.config_vars['font_family']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(rules_group, text="字号检查", variable=self.config_vars['font_size']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(rules_group, text="颜色数量检查", variable=self.config_vars['color_count']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(rules_group, text="主题一致性检查", variable=self.config_vars['theme_harmony']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(rules_group, text="缩略语解释检查", variable=self.config_vars['acronym_explanation']).pack(anchor=tk.W, padx=5, pady=2)
        
        return frame

    def _create_report_frame(self):
        """创建报告配置页面"""
        frame = ttk.Frame()
        
        # 报告配置
        report_group = ttk.LabelFrame(frame, text="报告配置", padding=10)
        report_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(report_group, text="包含问题汇总", variable=self.config_vars['include_summary']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(report_group, text="包含详细问题", variable=self.config_vars['include_details']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(report_group, text="包含改进建议", variable=self.config_vars['include_suggestions']).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(report_group, text="包含统计信息", variable=self.config_vars['include_statistics']).pack(anchor=tk.W, padx=5, pady=2)
        
        return frame

    def _create_advanced_frame(self):
        """创建高级配置页面"""
        frame = ttk.Frame()
        
        # 高级配置参数
        advanced_group = ttk.LabelFrame(frame, text="高级配置参数", padding=10)
        advanced_group.pack(fill=tk.X, padx=5, pady=5)
        
        self._create_config_row(advanced_group, "字号覆盖:", 'font_size_override', 0)
        self._create_config_row(advanced_group, "颜色阈值覆盖:", 'color_threshold_override', 1)
        
        # 词库路径配置
        dict_group = ttk.LabelFrame(frame, text="词库路径配置", padding=10)
        dict_group.pack(fill=tk.X, padx=5, pady=5)
        
        self._create_file_row(dict_group, "日语IT术语词典:", 'jp_terms_path', False, [("文本文件", "*.txt"), ("所有文件", "*.*")])
        self._create_file_row(dict_group, "术语映射表:", 'term_mapping_path', False, [("CSV文件", "*.csv"), ("所有文件", "*.*")])
        
        return frame

    def _create_bottom_frame(self, parent):
        """创建底部按钮和状态栏"""
        # 运行按钮
        run_frame = ttk.Frame(parent)
        run_frame.pack(fill=tk.X, pady=10)
        
        self.btn_run = ttk.Button(run_frame, text="运行工作流", command=self._on_run, style="Accent.TButton")
        self.btn_run.pack(pady=5)
        
        # 状态栏
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=5)
        
        self.var_status = tk.StringVar(value="就绪")
        status_label = ttk.Label(status_frame, textvariable=self.var_status, anchor=tk.W)
        status_label.pack(fill=tk.X)
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(parent, text="运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _create_file_row(self, parent, label, var_key, is_dir=False, file_types=None):
        """创建文件选择行"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(frame, text=label, width=20).pack(side=tk.LEFT)
        entry = ttk.Entry(frame, textvariable=self.config_vars[var_key])
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        def pick():
            if is_dir:
                path = filedialog.askdirectory()
            else:
                if file_types:
                    path = filedialog.askopenfilename(filetypes=file_types)
                else:
                    path = filedialog.askopenfilename()
            if path:
                self.config_vars[var_key].set(path)
        
        ttk.Button(frame, text="选择", command=pick, width=8).pack(side=tk.LEFT)
        
        return frame

    def _create_config_row(self, parent, label, var_key, row, var_type=str):
        """创建配置行"""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        
        if var_type == bool:
            widget = ttk.Checkbutton(parent, variable=self.config_vars[var_key])
        elif var_type == int:
            widget = ttk.Spinbox(parent, from_=0, to=999999, textvariable=self.config_vars[var_key], width=20)
        elif var_type == float:
            widget = ttk.Spinbox(parent, from_=0.0, to=2.0, increment=0.1, textvariable=self.config_vars[var_key], width=20)
        else:
            widget = ttk.Entry(parent, textvariable=self.config_vars[var_key], width=20)
        
        widget.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        
        return widget

    def _load_default_config(self):
        """加载默认配置"""
        try:
            if os.path.exists(self.config_vars['config_file'].get()):
                self._load_config()
        except Exception as e:
            self._log(f"加载默认配置失败: {e}")

    def _load_config(self):
        """加载配置文件"""
        try:
            config_path = self.config_vars['config_file'].get()
            if not config_path or not os.path.exists(config_path):
                messagebox.showerror("错误", "配置文件路径无效或文件不存在")
                return
            
            self.current_config = load_config(config_path)
            self._apply_config_to_ui()
            self._log(f"成功加载配置文件: {config_path}")
            
        except Exception as e:
            messagebox.showerror("加载配置失败", str(e))
            self._log(f"加载配置失败: {e}")

    def _save_config(self):
        """保存配置文件"""
        try:
            config_path = self.config_vars['config_file'].get()
            if not config_path:
                messagebox.showerror("错误", "请先设置配置文件路径")
                return
            
            # 从UI获取配置
            config_data = self._get_config_from_ui()
            
            # 确保目录存在
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            
            # 保存到文件
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config_data, f, default_flow_style=False, allow_unicode=True, indent=2)
            
            self._log(f"成功保存配置文件: {config_path}")
            messagebox.showinfo("成功", "配置文件已保存")
            
        except Exception as e:
            messagebox.showerror("保存配置失败", str(e))
            self._log(f"保存配置失败: {e}")

    def _reset_config(self):
        """重置为默认配置"""
        try:
            self._init_config_vars()
            self._log("配置已重置为默认值")
            messagebox.showinfo("成功", "配置已重置为默认值")
        except Exception as e:
            self._log(f"重置配置失败: {e}")

    def _apply_config_to_ui(self):
        """将配置应用到UI"""
        if not self.current_config:
            return
        
        try:
            # 基本配置
            if hasattr(self.current_config, 'jp_font_name'):
                self.config_vars['jp_font_name'].set(self.current_config.jp_font_name)
            if hasattr(self.current_config, 'min_font_size_pt'):
                self.config_vars['min_font_size_pt'].set(self.current_config.min_font_size_pt)
            if hasattr(self.current_config, 'color_count_threshold'):
                self.config_vars['color_count_threshold'].set(self.current_config.color_count_threshold)
            if hasattr(self.current_config, 'output_format'):
                self.config_vars['output_format'].set(self.current_config.output_format)
            
            # 自动修复开关
            if hasattr(self.current_config, 'autofix_font'):
                self.config_vars['autofix_font'].set(self.current_config.autofix_font)
            if hasattr(self.current_config, 'autofix_size'):
                self.config_vars['autofix_size'].set(self.current_config.autofix_size)
            if hasattr(self.current_config, 'autofix_color'):
                self.config_vars['autofix_color'].set(self.current_config.autofix_color)
            
            # LLM配置
            if hasattr(self.current_config, 'llm_enabled'):
                self.config_vars['llm_enabled'].set(self.current_config.llm_enabled)
            if hasattr(self.current_config, 'llm_model'):
                self.config_vars['llm_model'].set(self.current_config.llm_model)
            if hasattr(self.current_config, 'llm_temperature'):
                self.config_vars['llm_temperature'].set(self.current_config.llm_temperature)
            if hasattr(self.current_config, 'llm_max_tokens'):
                self.config_vars['llm_max_tokens'].set(self.current_config.llm_max_tokens)
            
            # 审查维度
            if hasattr(self.current_config, 'review_format'):
                self.config_vars['review_format'].set(self.current_config.review_format)
            if hasattr(self.current_config, 'review_logic'):
                self.config_vars['review_logic'].set(self.current_config.review_logic)
            if hasattr(self.current_config, 'review_acronyms'):
                self.config_vars['review_acronyms'].set(self.current_config.review_acronyms)
            if hasattr(self.current_config, 'review_fluency'):
                self.config_vars['review_fluency'].set(self.current_config.review_fluency)
            
            # 审查规则
            if hasattr(self.current_config, 'rules'):
                rules = self.current_config.rules
                if rules:
                    if 'font_family' in rules:
                        self.config_vars['font_family'].set(rules['font_family'])
                    if 'font_size' in rules:
                        self.config_vars['font_size'].set(rules['font_size'])
                    if 'color_count' in rules:
                        self.config_vars['color_count'].set(rules['color_count'])
                    if 'theme_harmony' in rules:
                        self.config_vars['theme_harmony'].set(rules['theme_harmony'])
                    if 'acronym_explanation' in rules:
                        self.config_vars['acronym_explanation'].set(rules['acronym_explanation'])
            
            # 报告配置
            if hasattr(self.current_config, 'report'):
                report = self.current_config.report
                if report:
                    if 'include_summary' in report:
                        self.config_vars['include_summary'].set(report['include_summary'])
                    if 'include_details' in report:
                        self.config_vars['include_details'].set(report['include_details'])
                    if 'include_suggestions' in report:
                        self.config_vars['include_suggestions'].set(report['include_suggestions'])
                    if 'include_statistics' in report:
                        self.config_vars['include_statistics'].set(report['include_statistics'])
            
        except Exception as e:
            self._log(f"应用配置到UI失败: {e}")

    def _get_config_from_ui(self):
        """从UI获取配置数据"""
        config_data = {}
        
        # 基本配置
        config_data['jp_font_name'] = self.config_vars['jp_font_name'].get()
        config_data['min_font_size_pt'] = self.config_vars['min_font_size_pt'].get()
        config_data['color_count_threshold'] = self.config_vars['color_count_threshold'].get()
        config_data['output_format'] = self.config_vars['output_format'].get()
        
        # 自动修复开关
        config_data['autofix_font'] = self.config_vars['autofix_font'].get()
        config_data['autofix_size'] = self.config_vars['autofix_size'].get()
        config_data['autofix_color'] = self.config_vars['autofix_color'].get()
        
        # 词库路径
        config_data['jp_terms_path'] = self.config_vars['jp_terms_path'].get()
        config_data['term_mapping_path'] = self.config_vars['term_mapping_path'].get()
        
        # LLM配置
        config_data['llm_enabled'] = self.config_vars['llm_enabled'].get()
        config_data['llm_model'] = self.config_vars['llm_model'].get()
        config_data['llm_temperature'] = self.config_vars['llm_temperature'].get()
        config_data['llm_max_tokens'] = self.config_vars['llm_max_tokens'].get()
        
        # 审查维度
        config_data['llm_review'] = {
            'review_format': self.config_vars['review_format'].get(),
            'review_logic': self.config_vars['review_logic'].get(),
            'review_acronyms': self.config_vars['review_acronyms'].get(),
            'review_fluency': self.config_vars['review_fluency'].get()
        }
        
        # 审查规则
        config_data['rules_review'] = {
            'font_family': self.config_vars['font_family'].get(),
            'font_size': self.config_vars['font_size'].get(),
            'color_count': self.config_vars['color_count'].get(),
            'theme_harmony': self.config_vars['theme_harmony'].get(),
            'acronym_explanation': self.config_vars['acronym_explanation'].get()
        }
        
        # 报告配置
        config_data['report'] = {
            'include_summary': self.config_vars['include_summary'].get(),
            'include_details': self.config_vars['include_details'].get(),
            'include_suggestions': self.config_vars['include_suggestions'].get(),
            'include_statistics': self.config_vars['include_statistics'].get()
        }
        
        return config_data

    def _on_run(self):
        """运行工作流"""
        # 验证输入
        input_ppt = self.config_vars['input_ppt'].get().strip()
        output_dir = self.config_vars['output_dir'].get().strip()
        mode = self.config_vars['mode'].get()
        
        if not input_ppt or not output_dir:
            messagebox.showerror("参数缺失", "请输入输入PPT文件和输出目录")
            return
        
        if not os.path.exists(input_ppt):
            messagebox.showerror("文件不存在", f"输入PPT文件不存在: {input_ppt}")
            return
        
        # 禁用运行按钮
        self.btn_run.config(state=tk.DISABLED)
        self.var_status.set("运行中...")
        self._log("开始运行工作流...")

        # 在后台线程中运行
        def job():
            try:
                # 创建输出目录
                os.makedirs(output_dir, exist_ok=True)
                
                # 生成输出路径
                from .cli import generate_output_paths
                parsing_result_path, report_path, output_ppt_path = generate_output_paths(input_ppt, mode, output_dir)
                
                # 从UI获取配置
                config_data = self._get_config_from_ui()
                
                # 创建临时配置文件
                temp_config_path = os.path.join(output_dir, "temp_config.yaml")
                with open(temp_config_path, 'w', encoding='utf-8') as f:
                    yaml.dump(config_data, f, default_flow_style=False, allow_unicode=True, indent=2)
                
                # 加载配置
                cfg = load_config(temp_config_path)
                
                # 应用覆盖参数
                if self.config_vars['llm_override'].get():
                    cfg.llm_enabled = (self.config_vars['llm_override'].get() == "on")
                
                if self.config_vars['font_size_override'].get():
                    try:
                        cfg.min_font_size_pt = int(self.config_vars['font_size_override'].get())
                    except ValueError:
                        pass
                
                if self.config_vars['color_threshold_override'].get():
                    try:
                        cfg.color_count_threshold = int(self.config_vars['color_threshold_override'].get())
                    except ValueError:
                        pass
                
                # 解析PPT
                self._log("步骤1: 解析PPT文件...")
                from .parser import parse_pptx
                parsing_data = parse_pptx(input_ppt, include_images=False)
                
                # 保存解析结果
                import json
                with open(parsing_result_path, "w", encoding="utf-8") as f:
                    json.dump(parsing_data, f, ensure_ascii=False, indent=2)
                self._log(f"✅ PPT解析完成，结果保存到: {parsing_result_path}")
                
                # 创建LLM客户端
                llm = None
                if cfg.llm_enabled:
                    llm = LLMClient()
                
                # 运行工作流
                if mode == "review":
                    self._log("步骤2: 开始审查模式...")
                    res = run_review_workflow(parsing_result_path, cfg, output_ppt_path, llm, input_ppt)
                else:
                    self._log("步骤2: 开始编辑模式...")
                    edit_req = self.config_vars['edit_req'].get()
                    res = run_edit_workflow(parsing_result_path, input_ppt, cfg, output_ppt_path, llm, edit_req)
                
                # 生成报告
                if hasattr(res, 'report_md') and res.report_md:
                    with open(report_path, "w", encoding="utf-8") as f:
                        f.write(res.report_md)
                    self._log(f"✅ 报告已生成: {report_path}")
                
                # 清理临时文件
                if os.path.exists(temp_config_path):
                    os.remove(temp_config_path)
                
                # 显示结果
                self._log(f"🎯 工作流完成！")
                self._log(f"   - 规则检查：{getattr(res, 'rule_issues_count', 0)} 个问题")
                self._log(f"   - LLM审查：{getattr(res, 'llm_issues_count', 0)} 个问题")
                self._log(f"   - 总计：{len(getattr(res, 'issues', []))} 个问题")
                
                self.var_status.set(f"完成：问题 {len(getattr(res, 'issues', []))}，输出：{output_dir}")
                
            except Exception as e:
                error_msg = f"运行失败: {e}"
                self._log(f"❌ {error_msg}")
                self.var_status.set(error_msg)
                messagebox.showerror("运行失败", str(e))
            finally:
                self.btn_run.config(state=tk.NORMAL)

    def _log(self, message):
        """添加日志消息"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.update_idletasks()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()


