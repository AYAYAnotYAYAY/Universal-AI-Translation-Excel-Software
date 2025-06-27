import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import threading
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string as col_idx_from_str
import translator

# Global font configuration
DEFAULT_FONT = ("Microsoft YaHei UI", 10)
HEADER_FONT = ("Microsoft YaHei UI", 10, "bold")

class ExcelPreview(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg="white", bd=2, relief="sunken")
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.v_scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.v_scroll.pack(side=tk.RIGHT, fill="y")
        self.h_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.h_scroll.pack(side=tk.BOTTOM, fill="x")

        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Button-1>", self._on_left_click) # Left click for source
        self.canvas.bind("<Button-3>", self._on_right_click) # Right click for target

        self.sheet = None
        self.cell_width = 100
        self.cell_height = 25
        self.header_height = 25
        self.row_header_width = 50

        self.selected_src_col = None
        self.selected_src_row = None
        self.selected_tgt_col = None

        self.drawn_rects = {} # To store canvas item IDs for redrawing/highlighting

    def _on_canvas_configure(self, event):
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.draw_sheet() # Redraw on resize

    def load_sheet(self, sheet):
        self.sheet = sheet
        self.selected_src_col = None
        self.selected_src_row = None
        self.selected_tgt_col = None
        self.draw_sheet()

    def draw_sheet(self):
        if not self.sheet:
            return

        self.canvas.delete("all")
        self.drawn_rects = {}

        max_row = self.sheet.max_row
        max_col = self.sheet.max_column

        # Draw column headers (A, B, C...)
        for c_idx in range(1, max_col + 1):
            col_letter = get_column_letter(c_idx)
            x1 = self.row_header_width + (c_idx - 1) * self.cell_width
            y1 = 0
            x2 = x1 + self.cell_width
            y2 = self.header_height
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="gray", fill="#f0f0f0")
            self.canvas.create_text(x1 + self.cell_width / 2, y1 + self.header_height / 2,
                                    text=col_letter, font=HEADER_FONT, tags=f"col_header_{c_idx}")

        # Draw row headers (1, 2, 3...)
        for r_idx in range(1, max_row + 1):
            x1 = 0
            y1 = self.header_height + (r_idx - 1) * self.cell_height
            x2 = self.row_header_width
            y2 = y1 + self.cell_height
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="gray", fill="#f0f0f0")
            self.canvas.create_text(x1 + self.row_header_width / 2, y1 + self.cell_height / 2,
                                    text=str(r_idx), font=HEADER_FONT, tags=f"row_header_{r_idx}")

        # Draw cells
        for r_idx in range(1, max_row + 1):
            for c_idx in range(1, max_col + 1):
                cell = self.sheet.cell(row=r_idx, column=c_idx)
                x1 = self.row_header_width + (c_idx - 1) * self.cell_width
                y1 = self.header_height + (r_idx - 1) * self.cell_height
                x2 = x1 + self.cell_width
                y2 = y1 + self.cell_height

                rect_id = self.canvas.create_rectangle(x1, y1, x2, y2, outline="gray", fill="white", tags=f"cell_{r_idx}_{c_idx}")
                self.drawn_rects[(r_idx, c_idx)] = rect_id

                # Handle merged cells for display
                display_value = cell.value
                for merged_range in self.sheet.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        # If it's a merged cell, only display value in top-left cell of the merged range
                        if cell.coordinate != merged_range.coord.split(':')[0]:
                            display_value = "" # Don't display value in other cells of merged range
                        break

                self.canvas.create_text(x1 + 5, y1 + self.cell_height / 2,
                                        text=str(display_value) if display_value is not None else "",
                                        anchor="w", font=DEFAULT_FONT, tags=f"text_{r_idx}_{c_idx}")

        # Adjust scroll region
        self.canvas.config(scrollregion=(0, 0,
                                         self.row_header_width + max_col * self.cell_width,
                                         self.header_height + max_row * self.cell_height))
        self._highlight_selections()

    def _get_cell_from_coords(self, x, y):
        if x < self.row_header_width or y < self.header_height:
            return None, None # Clicked on headers

        col = int((x - self.row_header_width) / self.cell_width) + 1
        row = int((y - self.header_height) / self.cell_height) + 1
        return row, col

    def _on_left_click(self, event):
        row, col = self._get_cell_from_coords(event.x, event.y)
        if row is not None and col is not None:
            self.selected_src_col = col
            self.selected_src_row = row
            self._highlight_selections()
            # print(f"Source selected: Column {get_column_letter(col)}, Row {row}") # For debugging

    def _on_right_click(self, event):
        row, col = self._get_cell_from_coords(event.x, event.y)
        if col is not None: # Only care about column for target
            self.selected_tgt_col = col
            self._highlight_selections()
            # print(f"Target selected: Column {get_column_letter(col)}") # For debugging

    def _highlight_selections(self):
        # Clear previous highlights
        self.canvas.delete("highlight")

        # Reset all cell fills to white first
        for (r, c), rect_id in self.drawn_rects.items():
            self.canvas.itemconfig(rect_id, fill="white")
        # Reset all column header fills to default
        for c_idx in range(1, self.sheet.max_column + 1):
            col_header_id = self.canvas.find_withtag(f"col_header_{c_idx}")
            if col_header_id:
                self.canvas.itemconfig(col_header_id, fill="#f0f0f0")

        if self.selected_src_col and self.selected_src_row:
            rect_id = self.drawn_rects.get((self.selected_src_row, self.selected_src_col))
            if rect_id:
                self.canvas.itemconfig(rect_id, fill="#e0ffe0", tags=("highlight",))
            # Highlight the source column header
            col_header_id = self.canvas.find_withtag(f"col_header_{self.selected_src_col}")
            if col_header_id:
                self.canvas.itemconfig(col_header_id, fill="#c0ffc0", tags=("highlight",))

        if self.selected_tgt_col:
            # Highlight the entire target column
            for r_idx in range(1, self.sheet.max_row + 1):
                rect_id = self.drawn_rects.get((r_idx, self.selected_tgt_col))
                if rect_id:
                    self.canvas.itemconfig(rect_id, fill="#e0e0ff", tags=("highlight",))
            # Highlight the target column header
            col_header_id = self.canvas.find_withtag(f"col_header_{self.selected_tgt_col}")
            if col_header_id:
                self.canvas.itemconfig(col_header_id, fill="#c0c0ff", tags=("highlight",))

    def get_selected_source_coords(self):
        if self.selected_src_col and self.selected_src_row:
            return get_column_letter(self.selected_src_col), self.selected_src_row
        return None, None

    def get_selected_target_col(self):
        if self.selected_tgt_col:
            return get_column_letter(self.selected_tgt_col)
        return None

    def set_selected_source_coords(self, col_letter, row):
        if col_letter and row:
            self.selected_src_col = col_idx_from_str(col_letter)
            self.selected_src_row = row
            self._highlight_selections()

    def set_selected_target_col(self, col_letter):
        if col_letter:
            self.selected_tgt_col = col_idx_from_str(col_letter)
            self._highlight_selections()

class TranslatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("通用AI翻译工具")
        self.geometry("1200x800") # Make window larger for preview
        self.tk.call("tk", "scaling", 1.2) # Scale UI for better visibility

        # Set default font for all widgets
        self.option_add("*Font", DEFAULT_FONT)

        self.translator = None
        self.translation_thread = None
        self.config_file = "config.json"

        # 主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 文件选择和预览 --- 
        file_preview_frame = ttk.LabelFrame(main_frame, text="1. Excel文件预览与列选择", padding="10")
        file_preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.file_path_var = tk.StringVar()
        ttk.Label(file_preview_frame, text="Excel文件:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(file_preview_frame, textvariable=self.file_path_var, width=80).grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(file_preview_frame, text="浏览...", command=self.browse_file).grid(row=0, column=2, padx=5)

        self.excel_preview = ExcelPreview(file_preview_frame)
        self.excel_preview.grid(row=1, column=0, columnspan=3, sticky=tk.NSEW, pady=5)
        
        file_preview_frame.columnconfigure(1, weight=1)
        file_preview_frame.rowconfigure(1, weight=1)

        # --- 语言设置 ---

        # --- 语言设置 ---
        lang_frame = ttk.LabelFrame(main_frame, text="2. 语言设置", padding="10")
        lang_frame.pack(fill=tk.X, pady=5)

        self.src_lang_var = tk.StringVar(value="俄语")
        self.tgt_lang_var = tk.StringVar(value="简体中文")
        ttk.Label(lang_frame, text="源语言:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(lang_frame, textvariable=self.src_lang_var).grid(row=0, column=1, padx=5)
        ttk.Label(lang_frame, text="目标语言:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(lang_frame, textvariable=self.tgt_lang_var).grid(row=0, column=3, padx=5)

        # --- API设置 ---
        api_frame = ttk.LabelFrame(main_frame, text="3. AI模型设置", padding="10")
        api_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.api_provider_var = tk.StringVar(value="DeepSeek")
        self.api_key_var = tk.StringVar()
        
        ttk.Label(api_frame, text="API提供商:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(api_frame, textvariable=self.api_provider_var, values=["DeepSeek"], state="readonly").grid(row=0, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(api_frame, text="API密钥:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(api_frame, textvariable=self.api_key_var, show="*", width=50).grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5)

        ttk.Label(api_frame, text="提示词模板:").grid(row=2, column=0, sticky=tk.NW, padx=5, pady=5)
        self.prompt_text = tk.Text(api_frame, height=10, wrap=tk.WORD)
        self.prompt_text.grid(row=2, column=1, columnspan=2, sticky=tk.NSEW, padx=5, pady=5)
        
        prompt_template = (
            "请严格遵循以下要求进行翻译：\n"
            "1. 将下面的文本内容从【{source_language}】逐行翻译成【{target_language}】。\n"
            "2. 保持与原文完全相同的行数，如果某行为空，则在翻译结果中也保留空行。\n"
            "3. 不要在每行译文前添加任何序号、数字、点、破折号或任何其他标记。\n"
            "4. 直接输出译文，不要包含任何解释或额外说明。\n\n"
            "待翻译内容如下：\n"
            "---批量翻译开始---\n"
            "{text_to_translate}\n"
            "---批量翻译结束---"
        )
        self.prompt_text.insert("1.0", prompt_template)
        
        api_frame.columnconfigure(1, weight=1)
        api_frame.rowconfigure(2, weight=1)

        # --- 控制和状态 ---
        control_frame = ttk.Frame(main_frame, padding="10")
        control_frame.pack(fill=tk.X)
        
        self.start_button = ttk.Button(control_frame, text="开始翻译", command=self.start_translation)
        self.start_button.pack(side=tk.RIGHT, padx=5)
        
        self.save_config_button = ttk.Button(control_frame, text="保存配置", command=lambda: self.save_config(self.config_file))
        self.save_config_button.pack(side=tk.LEFT, padx=5)
        
        self.load_config_button = ttk.Button(control_frame, text="加载配置", command=lambda: self.load_config(self.config_file))
        self.load_config_button.pack(side=tk.LEFT, padx=5)

        # --- 状态栏 ---
        status_frame = ttk.LabelFrame(main_frame, text="状态和日志", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.status_text = tk.Text(status_frame, height=5, wrap=tk.WORD, state=tk.DISABLED)
        self.status_text.pack(fill=tk.BOTH, expand=True)

        self.load_config(self.config_file) # 启动时自动加载配置

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*" ))
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.log(f"已选择文件: {file_path}")
            try:
                workbook = openpyxl.load_workbook(file_path)
                self.excel_preview.load_sheet(workbook.active)
                workbook.close()
                self.log("Excel文件已加载到预览。请在预览中左键选择源单元格，右键选择目标列。")
            except Exception as e:
                self.log(f"错误: 无法加载Excel文件进行预览: {e}")
                messagebox.showerror("错误", f"无法加载Excel文件进行预览:\n{e}")

    def start_translation(self):
        file_path = self.file_path_var.get()
        src_col_letter, src_row_num = self.excel_preview.get_selected_source_coords()
        tgt_col_letter = self.excel_preview.get_selected_target_col()

        if not file_path:
            messagebox.showerror("错误", "请先选择一个Excel文件。")
            return
        if not src_col_letter or not src_row_num:
            messagebox.showerror("错误", "请在预览中左键选择源语言的起始单元格。")
            return
        if not tgt_col_letter:
            messagebox.showerror("错误", "请在预览中右键选择目标语言的输出列。")
            return
        if not self.api_key_var.get():
            messagebox.showerror("错误", "请输入您的API密钥。")
            return
        
        # Store selected columns in vars for config saving
        self.src_col_var.set(src_col_letter)
        self.tgt_col_var.set(tgt_col_letter)
        self.src_row_var.set(src_row_num) # New var to store source row

        self.start_button.config(state=tk.DISABLED)
        self.log("翻译任务启动...")
        
        self.translation_thread = threading.Thread(target=self._translation_worker, daemon=True)
        self.translation_thread.start()

    def _translation_worker(self):
        workbook = None
        try:
            # 从GUI获取所有配置
            file_path = self.file_path_var.get()
            src_col_str = self.src_col_var.get()
            tgt_col_str = self.tgt_col_var.get()
            src_row_num = int(self.src_row_var.get())
            api_key = self.api_key_var.get()
            api_provider = self.api_provider_var.get()
            prompt_template = self.prompt_text.get("1.0", tk.END)
            source_language = self.src_lang_var.get()
            target_language = self.tgt_lang_var.get()

            # 初始化翻译器
            self.translator = translator.Translator(api_key, api_provider)
            
            # 加载工作簿并获取列索引
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active
            src_col_idx = column_index_from_string(src_col_str)
            tgt_col_idx = column_index_from_string(tgt_col_str)
            max_row = sheet.max_row

            self.log(f"文件加载成功，共 {max_row} 行。")
            self.log(f"翻译从 {src_col_str}{src_row_num} 开始到 {tgt_col_str} 列。")

            # 批量处理
            batch_size = 100
            for batch_num, start_row in enumerate(range(src_row_num, max_row + 1, batch_size), 1):
                end_row = min(start_row + batch_size - 1, max_row)
                self.log(f"正在处理批次 {batch_num} (行 {start_row}-{end_row})...")

                # 读取源文本
                sources = []
                for row_idx in range(start_row, end_row + 1):
                    cell_value = None
                    # 检查当前单元格是否在合并单元格区域内
                    for merged_range in sheet.merged_cells.ranges:
                        if sheet.cell(row=row_idx, column=src_col_idx).coordinate in merged_range:
                            # 如果在合并区域内，则取合并区域左上角的值
                            cell_value = sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
                            break
                    if cell_value is None:
                        # 如果不在合并区域内，则取当前单元格的值
                        cell_value = sheet.cell(row=row_idx, column=src_col_idx).value
                    sources.append(str(cell_value) if cell_value else "")

                # 调用API翻译
                translated = self.translator.translate_batch(sources, prompt_template, source_language, target_language)

                # 确保翻译结果数量与源文本一致
                if len(translated) < len(sources):
                    translated.extend([''] * (len(sources) - len(translated)))
                elif len(translated) > len(sources):
                    translated = translated[:len(sources)]

                # 写入结果
                for idx, row in enumerate(range(start_row, end_row + 1)):
                    sheet.cell(row=row, column=tgt_col_idx, value=translated[idx])
                
                # 实时保存
                workbook.save(file_path)
                self.log(f"批次 {batch_num} 完成并已保存。")

            self.log("所有翻译任务已完成！")
            self.after(0, lambda: messagebox.showinfo("完成", "文件翻译已成功完成！"))

        except Exception as e:
            self.log(f"发生严重错误: {e}")
            self.after(0, lambda: messagebox.showerror("错误", f"翻译过程中发生错误:\n{e}"))
        finally:
            if workbook:
                workbook.close()
            # 在主线程中重新启用按钮
            self.after(0, self.enable_start_button)

    def enable_start_button(self):
        self.start_button.config(state=tk.NORMAL)

    def save_config(self, file_path):
        config_data = {
            "api_provider": self.api_provider_var.get(),
            "api_key": self.api_key_var.get(),
            "source_language": self.src_lang_var.get(),
            "target_language": self.tgt_lang_var.get(),
            "prompt_template": self.prompt_text.get("1.0", tk.END),
            "last_file_path": self.file_path_var.get(),
            "src_col": self.src_col_var.get(),
            "tgt_col": self.tgt_col_var.get(),
            "src_row": self.src_row_var.get()
        }
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
            self.log(f"配置已保存到: {file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")

    def load_config(self, file_path):
        if not os.path.exists(file_path):
            self.log("未找到默认配置文件，使用默认设置。")
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            self.api_provider_var.set(config_data.get("api_provider", "DeepSeek"))
            self.api_key_var.set(config_data.get("api_key", ""))
            self.src_lang_var.set(config_data.get("source_language", "俄语"))
            self.tgt_lang_var.set(config_data.get("target_language", "简体中文"))
            self.prompt_text.delete("1.0", tk.END)
            self.prompt_text.insert("1.0", config_data.get("prompt_template", ""))
            
            last_file = config_data.get("last_file_path", "")
            if last_file and os.path.exists(last_file):
                self.file_path_var.set(last_file)
                try:
                    workbook = openpyxl.load_workbook(last_file)
                    self.excel_preview.load_sheet(workbook.active)
                    workbook.close()
                    self.log("Excel文件已加载到预览。请在预览中左键选择源单元格，右键选择目标列。")
                except Exception as e:
                    self.log(f"错误: 无法加载Excel文件进行预览: {e}")

            # 加载保存的源语言列和目标语言列
            src_col_letter = config_data.get("src_col", "")
            tgt_col_letter = config_data.get("tgt_col", "")
            src_row_num = config_data.get("src_row", "")

            self.src_col_var.set(src_col_letter)
            self.tgt_col_var.set(tgt_col_letter)
            self.src_row_var.set(src_row_num)

            # Set selections in ExcelPreview if file was loaded
            if self.excel_preview.sheet:
                self.excel_preview.set_selected_source_coords(src_col_letter, src_row_num)
                self.excel_preview.set_selected_target_col(tgt_col_letter)

            self.log(f"配置已从 {file_path} 加载")
        except Exception as e:
            messagebox.showerror("错误", f"加载配置失败: {e}")
            self.log(f"加载配置失败: {e}")

    def log(self, message):
        def _log():
            self.status_text.config(state=tk.NORMAL)
            self.status_text.insert(tk.END, f"{message}\n")
            self.status_text.see(tk.END)
            self.status_text.config(state=tk.DISABLED)
        if self.winfo_exists():
             self.after(0, _log)

if __name__ == "__main__":
    app = TranslatorApp()
    app.mainloop()