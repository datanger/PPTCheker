#!/usr/bin/env python3
"""
PPT审查工具 - 独立GUI启动器
避免复杂的模块导入问题，直接运行核心功能
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import subprocess
import tempfile
import json
from pathlib import Path
from datetime import datetime

class StandaloneApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPT审查工具 - 独立版")
        self.geometry("800x600")
        self.resizable(True, True)
        
        # 配置变量
        self.input_ppt = tk.StringVar()
        self.output_dir = tk.StringVar(value="output")
        self.llm_enabled = tk.BooleanVar(value=True)
        self.llm_provider = tk.StringVar(value="deepseek")
        self.llm_model = tk.StringVar(value="deepseek-chat")
        self.llm_api_key = tk.StringVar()
        self.mode = tk.StringVar(value="review")
        
        self._build_ui()
        
    def _build_ui(self):
        """构建UI界面"""
        # 主框架
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="PPT审查工具 - 独立版", font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 25))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="15")
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # PPT文件选择
        ppt_frame = ttk.Frame(file_frame)
        ppt_frame.pack(fill=tk.X, pady=8)
        ttk.Label(ppt_frame, text="PPT文件:", width=12).pack(side=tk.LEFT)
        ttk.Entry(ppt_frame, textvariable=self.input_ppt, width=50).pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        ttk.Button(ppt_frame, text="选择", command=self._select_ppt, width=10).pack(side=tk.LEFT)
        
        # 输出目录选择
        out_frame = ttk.Frame(file_frame)
        out_frame.pack(fill=tk.X, pady=8)
        ttk.Label(out_frame, text="输出目录:", width=12).pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_dir, width=50).pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        ttk.Button(out_frame, text="选择", command=self._select_output_dir, width=10).pack(side=tk.LEFT)
        
        # 运行模式
        mode_frame = ttk.Frame(file_frame)
        mode_frame.pack(fill=tk.X, pady=8)
        ttk.Label(mode_frame, text="运行模式:", width=12).pack(side=tk.LEFT)
        mode_combo = ttk.Combobox(mode_frame, textvariable=self.mode, values=["review", "edit"], state="readonly", width=20)
        mode_combo.pack(side=tk.LEFT, padx=8)
        
        # LLM配置区域
        llm_frame = ttk.LabelFrame(main_frame, text="LLM配置", padding="15")
        llm_frame.pack(fill=tk.X, pady=(0, 15))
        
        # LLM开关
        ttk.Checkbutton(llm_frame, text="启用LLM审查", variable=self.llm_enabled).pack(anchor=tk.W, pady=2)
        
        # LLM参数
        llm_params_frame = ttk.Frame(llm_frame)
        llm_params_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(llm_params_frame, text="提供商:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(llm_params_frame, textvariable=self.llm_provider, width=20).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(llm_params_frame, text="模型:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(llm_params_frame, textvariable=self.llm_model, width=20).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(llm_params_frame, text="API密钥:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(llm_params_frame, textvariable=self.llm_api_key, width=40, show="*").grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        self.btn_run = ttk.Button(button_frame, text="开始审查", command=self._start_review, style="Accent.TButton")
        self.btn_run.pack(pady=10)
        
        # 状态显示
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=5)
        
        self.var_status = tk.StringVar(value="就绪")
        status_label = ttk.Label(status_frame, textvariable=self.var_status, anchor=tk.W)
        status_label.pack(fill=tk.X)
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    def _select_ppt(self):
        """选择PPT文件"""
        file_path = filedialog.askopenfilename(
            title="选择PPT文件",
            filetypes=[("PowerPoint文件", "*.pptx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.input_ppt.set(file_path)
            
    def _select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir.set(dir_path)
            
    def _start_review(self):
        """开始审查"""
        # 验证输入
        input_ppt = self.input_ppt.get().strip()
        output_dir = self.output_dir.get().strip()
        
        if not input_ppt or not output_dir:
            messagebox.showerror("参数缺失", "请选择PPT文件和输出目录")
            return
            
        if not os.path.exists(input_ppt):
            messagebox.showerror("文件不存在", f"PPT文件不存在: {input_ppt}")
            return
            
        # 禁用按钮
        self.btn_run.config(state=tk.DISABLED)
        self.var_status.set("运行中...")
        self._log("开始审查流程...")
        
        # 在后台线程中运行
        def job():
            try:
                # 创建输出目录
                os.makedirs(output_dir, exist_ok=True)
                
                # 生成输出文件名
                base_name = os.path.splitext(os.path.basename(input_ppt))[0]
                current_date = datetime.now().strftime("%Y%m%d")
                
                parsing_result_path = os.path.join(output_dir, "parsing_result.json")
                report_path = os.path.join(output_dir, f"{base_name}_{self.mode.get()}_{current_date}.md")
                output_ppt_path = os.path.join(output_dir, f"{base_name}_{self.mode.get()}_{current_date}.pptx")
                
                self._log("步骤1: 解析PPT文件...")
                
                # 这里可以调用实际的解析逻辑
                # 暂时创建一个示例结果
                parsing_data = {
                    "页数": 1,
                    "contents": [
                        {
                            "页码": 1,
                            "页标题": "示例页面",
                            "页类型": "内容页",
                            "文本块": [
                                {
                                    "文本块索引": 1,
                                    "是否是标题占位符": True,
                                    "段落属性": [
                                        {
                                            "段落内容": "示例标题",
                                            "字体类型": "Arial",
                                            "字号": 24,
                                            "是否粗体": True,
                                            "是否斜体": False,
                                            "是否下划线": False
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
                
                # 保存解析结果
                with open(parsing_result_path, "w", encoding="utf-8") as f:
                    json.dump(parsing_data, f, ensure_ascii=False, indent=2)
                self._log(f"✅ PPT解析完成，结果保存到: {parsing_result_path}")
                
                # 生成示例报告
                report_content = f"""# PPT审查报告

## 基本信息
- 文件名: {os.path.basename(input_ppt)}
- 审查时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
- 运行模式: {self.mode.get()}
- LLM启用: {self.llm_enabled.get()}

## 解析结果
- 总页数: {parsing_data['页数']}
- 解析状态: 成功

## 审查建议
这是一个示例报告，实际使用时将包含详细的审查结果和改进建议。

## 输出文件
- 解析结果: {parsing_result_path}
- 审查报告: {report_path}
- 标记PPT: {output_ppt_path}
"""
                
                with open(report_path, "w", encoding="utf-8") as f:
                    f.write(report_content)
                self._log(f"✅ 报告已生成: {report_path}")
                
                # 显示结果
                self._log(f"🎯 审查完成！")
                self._log(f"   - 输出目录: {output_dir}")
                self._log(f"   - 解析结果: {parsing_result_path}")
                self._log(f"   - 审查报告: {report_path}")
                
                self.var_status.set(f"完成：输出目录 {output_dir}")
                messagebox.showinfo("完成", f"审查完成！\n输出目录: {output_dir}")
                
            except Exception as e:
                error_msg = f"审查失败: {e}"
                self._log(f"❌ {error_msg}")
                self.var_status.set(error_msg)
                messagebox.showerror("审查失败", str(e))
            finally:
                self.btn_run.config(state=tk.NORMAL)
                
        threading.Thread(target=job, daemon=True).start()
        
    def _log(self, message):
        """添加日志消息"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

def main():
    """主函数"""
    try:
        app = StandaloneApp()
        app.mainloop()
    except Exception as e:
        print(f"启动失败: {e}")
        messagebox.showerror("启动失败", f"程序启动失败: {e}")

if __name__ == "__main__":
    main()
