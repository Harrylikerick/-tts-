import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import time
import requests.exceptions
from docx import Document
import fitz  # PyMuPDF
import re
from gtts import gTTS
from typing import *
import logging
from PIL import Image, ImageTk
import json

def batch_text_to_speech(file_path, folder_path, progress_callback=None):
    """批量转换Word和PDF文件中的纯梵文段落为音频"""
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    texts = []
    titles = []
    current_title = "未命名"
    sanskrit_ratio_threshold = 0.8

    # 记录处理进度
    total_progress = 0
    current_progress = 0

    try:
        if file_path.lower().endswith('.pdf'):
            pdf_document = fitz.open(file_path)
            total_progress += len(pdf_document)
            logging.info(f"开始处理PDF文件: {file_path}")
            
            for page_index, page in enumerate(pdf_document):
                page_text_blocks = page.get_text("blocks")
                logging.info(f"处理第 {page_index + 1} 页")
                
                for block in page_text_blocks:
                    text = block[4].strip()
                    if not text:
                        continue

                    process_text_block(text, texts, titles, current_title, sanskrit_ratio_threshold)

                current_progress += 1
                if progress_callback:
                    progress_callback(current_progress, total_progress)

            pdf_document.close()

        elif file_path.lower().endswith('.docx'):
            doc = Document(file_path)
            total_progress += len(doc.paragraphs)
            logging.info(f"开始处理Word文件: {file_path}")

            for para_index, para in enumerate(doc.paragraphs):
                text = para.text.strip()
                if text:
                    process_text_block(text, texts, titles, current_title, sanskrit_ratio_threshold)

                current_progress += 1
                if progress_callback:
                    progress_callback(current_progress, total_progress)

        else:
            error_msg = "不支持的文件类型，请提供 .docx 或 .pdf 文件"
            logging.error(error_msg)
            raise ValueError(error_msg)

        # 音频转换阶段
        os.environ['HTTP_PROXY'] = 'http://127.0.0.1:7890'
        os.environ['HTTPS_PROXY'] = 'http://127.0.0.1:7890'
        
        audio_paths = []
        total_progress = len(texts)
        current_progress = 0

        record_path = os.path.join(folder_path, "audio_record.txt")
        index_counter = 1

        for text, title in zip(texts, titles):
            try:
                audio_name = generate_audio_name(text, title, index_counter)
                audio_path = os.path.join(folder_path, f"{audio_name}.mp3")
                
                # 记录转换信息
                with open(record_path, "a", encoding='utf-8') as f:
                    f.write(f"音频文件: {audio_name}.mp3\n标题: {title}\n梵文内容: {text}\n{'='*50}\n")

                # 尝试转换音频
                convert_to_audio(text, audio_path)
                audio_paths.append(audio_path)
                logging.info(f"成功转换: {audio_name}.mp3")

                current_progress += 1
                if progress_callback:
                    progress_callback(current_progress, total_progress)

            except Exception as e:
                logging.error(f"转换失败 '{title}': {str(e)}")
                continue

        return audio_paths

    except Exception as e:
        logging.error(f"处理文件时发生错误: {str(e)}")
        raise

    finally:
        os.environ.pop('HTTP_PROXY', None)
        os.environ.pop('HTTPS_PROXY', None)

def process_text_block(text, texts, titles, current_title, sanskrit_ratio_threshold):
    """处理文本块，提取梵文内容"""
    if "卍" in text:
        parts = re.split(r'卍', text)
        if len(parts) >= 3:
            current_title = parts[1].strip()
            sanskrit_text = extract_sanskrit_text(parts[2].strip())
            if sanskrit_text:
                texts.append(sanskrit_text)
                titles.append(current_title)
        elif len(parts) == 2 and text.startswith("卍"):
            current_title = parts[1].strip() or "未命名"
            sanskrit_text = extract_sanskrit_text(parts[1].strip())
            if sanskrit_text:
                texts.append(sanskrit_text)
                titles.append(current_title)
    else:
        sanskrit_text = extract_sanskrit_text(text)
        if sanskrit_text and is_sanskrit_paragraph(text, sanskrit_text, sanskrit_ratio_threshold):
            texts.append(sanskrit_text)
            titles.append(current_title)

def extract_sanskrit_text(text):
    """提取梵文文本"""
    sanskrit_pattern = get_sanskrit_pattern()
    matches = sanskrit_pattern.findall(text)
    return ''.join(matches).strip()

def is_sanskrit_paragraph(text, sanskrit_text, threshold):
    """判断是否为梵文段落"""
    return len(text) > 0 and len(sanskrit_text) / len(text) >= threshold

def generate_audio_name(text, title, index_counter):
    """生成音频文件名"""
    prefix_number = None
    if "卍" in text:
        parts = re.split(r'卍', text)
        if parts and parts[0].strip():
            number_match = re.match(r'^(\d+)', parts[0].strip())
            if number_match:
                prefix_number = number_match.group(1)

    audio_name = f"{prefix_number}_{title}" if prefix_number else f"{index_counter}_{title}"
    return re.sub(r'[<>:"/\\|?*]', '', audio_name)

def convert_to_audio(text, audio_path, max_retries=3, retry_delay=2):
    """转换文本为音频"""
    for attempt in range(max_retries):
        try:
            tts = gTTS(text=text, lang='en')
            tts.save(audio_path)
            time.sleep(1)  # 避免请求过于频繁
            return
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            time.sleep(retry_delay)


class SanskritAudioConverterGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("梵文音频批量转换工具 v1.1")
        self.geometry("800x600")
        
        # 设置应用程序图标
        try:
            self.iconbitmap("app_icon.ico")
        except:
            pass  # 如果图标文件不存在则使用默认图标

        # 加载配置文件
        self.config_file = "audio_converter_config.json"
        self.load_config()

        # 创建主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建样式
        style = ttk.Style()
        style.configure("Title.TLabel", font=("Helvetica", 16, "bold"))
        style.configure("Status.TLabel", font=("Helvetica", 10))

        # 标题
        title_label = ttk.Label(main_frame, text="梵文音频批量转换工具", style="Title.TLabel")
        title_label.pack(pady=10)

        # 输入文件框架
        input_frame = ttk.LabelFrame(main_frame, text="输入设置", padding="5")
        input_frame.pack(fill=tk.X, padx=5, pady=5)

        # 输入文件路径
        input_file_label = ttk.Label(input_frame, text="选择 Word 或 PDF 文件:")
        input_file_label.pack(anchor=tk.W)
        
        input_file_frame = ttk.Frame(input_frame)
        input_file_frame.pack(fill=tk.X, pady=2)
        
        self.input_file_path_var = tk.StringVar()
        self.input_file_entry = ttk.Entry(input_file_frame, textvariable=self.input_file_path_var)
        self.input_file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.input_file_button = ttk.Button(input_file_frame, text="浏览", command=self.select_input_file)
        self.input_file_button.pack(side=tk.RIGHT)

        # 输出文件夹框架
        output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="5")
        output_frame.pack(fill=tk.X, padx=5, pady=5)

        # 输出文件夹路径
        output_folder_label = ttk.Label(output_frame, text="选择音频输出文件夹:")
        output_folder_label.pack(anchor=tk.W)
        
        output_folder_frame = ttk.Frame(output_frame)
        output_folder_frame.pack(fill=tk.X, pady=2)
        
        self.output_folder_path_var = tk.StringVar()
        self.output_folder_entry = ttk.Entry(output_folder_frame, textvariable=self.output_folder_path_var)
        self.output_folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.output_folder_button = ttk.Button(output_folder_frame, text="浏览", command=self.select_output_folder)
        self.output_folder_button.pack(side=tk.RIGHT)

        # 进度框架
        progress_frame = ttk.LabelFrame(main_frame, text="转换进度", padding="5")
        progress_frame.pack(fill=tk.X, padx=5, pady=5)

        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # 状态标签
        self.status_label = ttk.Label(progress_frame, text="就绪", style="Status.TLabel", wraplength=700)
        self.status_label.pack(fill=tk.X, pady=5)

        # 转换按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        # 转换按钮
        self.convert_button = ttk.Button(button_frame, text="音频转换", command=self.start_conversion)
        self.convert_button.pack(side=tk.LEFT, padx=5)

        # 添加PDF转Word按钮
        self.pdf_to_word_button = ttk.Button(button_frame, text="PDF转Word", command=self.start_pdf_to_word)
        self.pdf_to_word_button.pack(side=tk.LEFT, padx=5)

        # 初始化日志
        self.setup_logging()

    def setup_logging(self):
        """设置日志记录"""
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)

        log_file = os.path.join(log_dir, f"conversion_{time.strftime('%Y%m%d_%H%M%S')}.log")
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )

    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'last_file' in config:
                        self.input_file_path_var.set(config['last_file'])
                    if 'last_folder' in config:
                        self.output_folder_path_var.set(config['last_folder'])
        except Exception as e:
            logging.error(f"加载配置文件失败: {e}")

    def save_config(self):
        """保存配置文件"""
        try:
            config = {
                'last_file': self.input_file_path_var.get(),
                'last_folder': self.output_folder_path_var.get(),
                'file_history': [],
                'folder_history': []
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"保存配置文件失败: {e}")

    def select_input_file(self):
        filetypes = (
            ('Word files', '*.docx'),
            ('PDF files', '*.pdf'),
            ('All files', '*.*')
        )
        initial_dir = os.path.dirname(self.input_file_path_var.get()) if self.input_file_path_var.get() else os.getcwd()
        filepath = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=filetypes,
            initialdir=initial_dir
        )
        if filepath:
            self.input_file_path_var.set(filepath)
            logging.info(f"已选择输入文件: {filepath}")
            self.save_config()

    def select_output_folder(self):
        initial_dir = self.output_folder_path_var.get() if self.output_folder_path_var.get() else os.getcwd()
        folder_selected = filedialog.askdirectory(
            title="选择输出文件夹",
            initialdir=initial_dir
        )
        if folder_selected:
            self.output_folder_path_var.set(folder_selected)
            logging.info(f"已选择输出文件夹: {folder_selected}")
            self.save_config()

    def update_progress(self, current, total):
        """更新进度条"""
        progress = (current / total) * 100
        self.progress_var.set(progress)
        self.update_idletasks()

    def start_conversion(self):
        input_file_path = self.input_file_path_var.get()
        output_folder_path = self.output_folder_path_var.get()

        if not input_file_path or not output_folder_path:
            messagebox.showerror("错误", "请先选择输入文件和输出文件夹。")
            return

        if not os.path.exists(output_folder_path):
            try:
                os.makedirs(output_folder_path)
                logging.info(f"创建输出文件夹: {output_folder_path}")
            except Exception as e:
                error_msg = f"无法创建输出文件夹: {e}"
                logging.error(error_msg)
                messagebox.showerror("错误", error_msg)
                return

        self.status_label.config(text="转换开始，请稍候...")
        self.convert_button.config(state='disabled')
        self.progress_var.set(0)
        self.update_idletasks()

        try:
            logging.info("开始转换过程")
            audio_paths = batch_text_to_speech(input_file_path, output_folder_path)
            if audio_paths:
                success_message = f"成功生成 {len(audio_paths)} 个音频文件到文件夹: {output_folder_path}"
                self.status_label.config(text=success_message)
                logging.info(success_message)
                messagebox.showinfo("完成", success_message)
            else:
                error_message = "音频文件生成过程中可能出现错误，请查看日志文件获取详细信息。"
                self.status_label.config(text=error_message)
                logging.warning(error_message)
                messagebox.showwarning("警告", error_message)

        except Exception as e:
            error_message = f"转换过程中发生错误: {e}"
            self.status_label.config(text=error_message)
            logging.error(error_message)
            messagebox.showerror("错误", error_message)

        finally:
            self.convert_button.config(state='normal')
            self.progress_var.set(0)
            logging.info("转换过程结束")


    def pdf_to_word(self, input_file, output_file, progress_callback=None):
        """将PDF转换为Word文档"""
        try:
            pdf_document = fitz.open(input_file)
            doc = Document()
            total_pages = len(pdf_document)
            current_page = 0

            for page in pdf_document:
                # 获取页面上的文本块
                blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
                for block in blocks:
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                text = span["text"].strip()
                                if not text:
                                    continue
                                
                                try:
                                    font_name = span.get("font", "")
                                    font_size = span.get("size", 0)
                                    
                                    # 跳过MicrosoftYaHei-Bold和ZH-SANSKRIT字体的文本
                                    if "microsoftyahei-bold" in font_name.lower() or "zh-sanskrit" in font_name.lower():
                                        logging.info(f"跳过{font_name}字体文本: {text}")
                                        continue
                                    
                                    # 跳过Arial-BoldMT字体且字号为10.3的数字文本
                                    if font_name.lower() == "arial-boldmt" and abs(font_size - 10.3) < 0.1 and text.strip().isdigit():
                                        logging.info(f"跳过Arial-BoldMT字体数字文本: {text}")
                                        continue
                                    
                                    # 添加文本到Word文档
                                    paragraph = doc.add_paragraph()
                                    run = paragraph.add_run(text)
                                    run.font.name = 'Times New Roman'  # 使用默认字体
                                    logging.debug(f"添加文本到Word: {text}")
                                    
                                except Exception as e:
                                    logging.warning(f"处理文本时出现错误: {str(e)}，文本: {text}")
                                    # 如果处理出错，仍然添加文本
                                    paragraph = doc.add_paragraph()
                                    run = paragraph.add_run(text)
                                    run.font.name = 'ZH-SANSKRIT'
                
                current_page += 1
                if progress_callback:
                    progress_callback(current_page, total_pages)

            # 保存Word文档
            doc.save(output_file)
            logging.info(f"成功保存Word文档: {output_file}")
            return True

        except Exception as e:
            logging.error(f"PDF转Word过程中发生错误: {str(e)}")
            raise

        finally:
            if 'pdf_document' in locals():
                pdf_document.close()

    def start_pdf_to_word(self):
        input_file_path = self.input_file_path_var.get()
        if not input_file_path:
            messagebox.showerror("错误", "请先选择输入PDF文件。")
            return

        if not input_file_path.lower().endswith('.pdf'):
            messagebox.showerror("错误", "请选择PDF文件进行转换。")
            return

        # 选择Word文件保存位置
        output_file = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")],
            title="选择保存位置"
        )

        if not output_file:
            return

        self.status_label.config(text="正在转换PDF到Word，请稍候...")
        self.convert_button.config(state='disabled')
        self.pdf_to_word_button.config(state='disabled')
        self.progress_var.set(0)
        self.update_idletasks()

        try:
            self.pdf_to_word(input_file_path, output_file, self.update_progress)
            success_message = f"PDF已成功转换为Word文档: {output_file}"
            self.status_label.config(text=success_message)
            logging.info(success_message)
            messagebox.showinfo("完成", success_message)

        except Exception as e:
            error_message = f"转换过程中发生错误: {e}"
            self.status_label.config(text=error_message)
            logging.error(error_message)
            messagebox.showerror("错误", error_message)

        finally:
            self.convert_button.config(state='normal')
            self.pdf_to_word_button.config(state='normal')
            self.progress_var.set(0)

if __name__ == "__main__":
    app = SanskritAudioConverterGUI()
    app.mainloop()