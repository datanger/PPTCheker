"""
简洁GUI：运行参数配置与工作流执行（保留CLI与GUI两种模式）。

功能：
- 选择输入PPT/目录、用户审查需求文档、配置文件；
- 选择模式（review/edit）、输出PPT路径；
- 配置大模型 provider/model/api_key（endpoint 将根据 model 推断，可覆盖）；
- 运行并在状态栏显示结果。
"""
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .config import load_config, ToolConfig
from .user_req import parse_user_requirements
from .workflow import run_review_workflow, run_edit_workflow
from .llm import LLMClient


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PPT 审查工具 - GUI")
        self.geometry("720x460")

        self.var_input = tk.StringVar()
        self.var_user_req = tk.StringVar()
        self.var_config = tk.StringVar(value="configs/config.yaml")
        self.var_output = tk.StringVar(value="out/标记版.pptx")
        self.var_mode = tk.StringVar(value="review")

        self.var_provider = tk.StringVar(value="deepseek")
        self.var_model = tk.StringVar(value="deepseek-chat")
        self.var_endpoint = tk.StringVar(value="")
        self.var_api_key = tk.StringVar(value=os.getenv("LLM_API_KEY", ""))

        self._build_ui()

    def _row_filepicker(self, parent, label, var, is_dir=False):
        fr = ttk.Frame(parent)
        ttk.Label(fr, text=label, width=16).pack(side=tk.LEFT)
        ent = ttk.Entry(fr, textvariable=var)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        def pick():
            if is_dir:
                p = filedialog.askdirectory()
            else:
                p = filedialog.askopenfilename()
            if p:
                var.set(p)
        ttk.Button(fr, text="选择", command=pick).pack(side=tk.LEFT)
        return fr

    def _row_savepicker(self, parent, label, var):
        fr = ttk.Frame(parent)
        ttk.Label(fr, text=label, width=16).pack(side=tk.LEFT)
        ent = ttk.Entry(fr, textvariable=var)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        def pick():
            p = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PPTX", ".pptx")])
            if p:
                var.set(p)
        ttk.Button(fr, text="保存到", command=pick).pack(side=tk.LEFT)
        return fr

    def _build_ui(self):
        pad = 8
        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)

        self._row_filepicker(frm, "输入(PPT/目录)", self.var_input).pack(fill=tk.X, pady=4)
        self._row_filepicker(frm, "审查需求文档", self.var_user_req).pack(fill=tk.X, pady=4)
        self._row_filepicker(frm, "配置文件", self.var_config).pack(fill=tk.X, pady=4)

        fr2 = ttk.Frame(frm)
        ttk.Label(fr2, text="运行模式", width=16).pack(side=tk.LEFT)
        ttk.Combobox(fr2, textvariable=self.var_mode, values=["review", "edit"], width=12, state="readonly").pack(side=tk.LEFT)
        fr2.pack(fill=tk.X, pady=4)

        self._row_savepicker(frm, "输出PPT", self.var_output).pack(fill=tk.X, pady=4)

        sep = ttk.Separator(frm)
        sep.pack(fill=tk.X, pady=8)

        ttk.Label(frm, text="大模型配置").pack(anchor=tk.W)
        fr3 = ttk.Frame(frm)
        ttk.Label(fr3, text="Provider", width=16).pack(side=tk.LEFT)
        ttk.Entry(fr3, textvariable=self.var_provider).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        fr3.pack(fill=tk.X, pady=2)

        fr4 = ttk.Frame(frm)
        ttk.Label(fr4, text="Model", width=16).pack(side=tk.LEFT)
        ttk.Entry(fr4, textvariable=self.var_model).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        fr4.pack(fill=tk.X, pady=2)

        fr5 = ttk.Frame(frm)
        ttk.Label(fr5, text="Endpoint(可空)", width=16).pack(side=tk.LEFT)
        ttk.Entry(fr5, textvariable=self.var_endpoint).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        fr5.pack(fill=tk.X, pady=2)

        fr6 = ttk.Frame(frm)
        ttk.Label(fr6, text="API Key", width=16).pack(side=tk.LEFT)
        ttk.Entry(fr6, textvariable=self.var_api_key, show="*").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        fr6.pack(fill=tk.X, pady=2)

        self.btn_run = ttk.Button(frm, text="运行", command=self.on_run)
        self.btn_run.pack(pady=10)

        self.var_status = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.var_status, anchor=tk.W).pack(fill=tk.X, padx=pad, pady=(0, pad))

    def on_run(self):
        input_path = self.var_input.get().strip()
        config_path = self.var_config.get().strip()
        user_req = self.var_user_req.get().strip()
        output_ppt = self.var_output.get().strip()
        mode = self.var_mode.get()
        if not input_path or not config_path or not output_ppt:
            messagebox.showerror("参数缺失", "请输入 输入、配置文件、输出PPT")
            return
        self.btn_run.config(state=tk.DISABLED)
        self.var_status.set("运行中...")

        def job():
            try:
                cfg: ToolConfig = load_config(config_path)
                if user_req:
                    cfg = parse_user_requirements(user_req, cfg)

                # 构建 LLM 客户端（endpoint 可为空，内部按 model 推断）
                endpoint = self.var_endpoint.get().strip() or None
                api_key = self.var_api_key.get().strip() or None
                model = self.var_model.get().strip() or None
                llm = LLMClient(endpoint=endpoint, api_key=api_key, model=model)

                if mode == "review":
                    res = run_review_workflow(input_path, cfg, output_ppt, llm)
                else:
                    res = run_edit_workflow(input_path, cfg, output_ppt, llm)
                self.var_status.set(f"完成：问题 {len(res.issues)}，输出：{output_ppt}")
            except Exception as e:
                self.var_status.set(f"失败：{e}")
                messagebox.showerror("运行失败", str(e))
            finally:
                self.btn_run.config(state=tk.NORMAL)

        threading.Thread(target=job, daemon=True).start()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()


