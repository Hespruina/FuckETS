import os
import json
import re
import sys
import platform
from pathlib import Path

# 检查系统
is_win = platform.system() == 'Windows'

# 库的导入情况
PYQT_AVAILABLE = False
# tk的文件对话框
tk_fd = None
if is_win:
    try:
        import tkinter as tk
        from tkinter import filedialog as tk_fd
        # 隐藏tk主窗口
        root = tk.Tk()
        root.withdraw()
        # 试下PyQt5
        try:
            from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                                        QLineEdit, QPushButton, QTextEdit)
            from PyQt5.QtCore import Qt, QPoint
            from PyQt5.QtGui import QFont
            PYQT_AVAILABLE = True
        except ImportError:
            pass
    except ImportError:
        pass

class ETS数据提取器:
    def __init__(self, root_dir):
        self.root_dir = Path(root_dir).resolve()
        if not self.root_dir.is_dir():
            raise ValueError(f"无效目录: {root_dir}")
        self.all_data = []
        self.setup_colors()
        self._parse_all_data()

    def setup_colors(self):
        # 设置颜色代码
        self.RED = '\033[1;31m'
        self.GREEN = '\033[1;32m'
        self.YELLOW = '\033[1;33m'
        self.BLUE = '\033[1;34m'
        self.PURPLE = '\033[1;35m'
        self.CYAN = '\033[1;36m'
        self.WHITE = '\033[1;37m'
        self.ORANGE = '\033[1;38;5;214m'
        self.MAGENTA = '\033[1;35m'
        self.NC = '\033[0m'
        if platform.system() == 'Windows':
            try:
                import ctypes
                kernel32 = ctypes.windll.kernel32
                kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
            except Exception:
                for attr in ['RED', 'GREEN', 'YELLOW', 'BLUE', 'PURPLE', 'CYAN', 'WHITE', 'ORANGE', 'MAGENTA', 'NC']:
                    setattr(self, attr, '')

    def _parse_all_data(self):
        # 解析所有文件
        template_dirs = []
        for dir_path in self.root_dir.rglob("*"):
            if dir_path.is_dir() and (dir_path / "ctrl.json").exists() and \
               (dir_path / "info.json").exists() and (dir_path / "res.json").exists():
                template_dirs.append(dir_path)
                self._parse_pc_template(dir_path)
                
        # 同时支持移动版
        for content_path in self.root_dir.rglob("content.json"):
            self._parse_content_file(content_path)
            
        # 输出结果
        if template_dirs:
            print(f"{self.GREEN}✅ 成功解析了 {len(template_dirs)} 个题库文件夹{self.NC}")
    
    def _parse_pc_template(self, dir_path: Path):
        # 解析电脑版题库
        try:
            # 读取 info.json
            with open(dir_path / "info.json", 'r', encoding='utf-8') as f:
                info_data = json.load(f)
            
            # 读取 res.json
            with open(dir_path / "res.json", 'r', encoding='utf-8') as f:
                res_data = json.load(f)
            
            # 创建信息映射
            info_map = {item['code_id']: item['code_value'] for item in info_data}
            
            # 处理考试类型
            for exam_type in res_data.get('exam_type_list', []):
                exam_type_name = exam_type.get('exam_type_name', '')
                exam_type_collect = exam_type.get('exam_type_collect', '')
                
                for exam in exam_type.get('exam_list', []):
                    exam_id = exam.get('exam_id', '')
                    
                    # 处理不同类型
                    if exam_type_collect == 'collector.read':
                        # 模仿朗读
                        content_file = dir_path / "material" / "content.mp3"
                        if content_file.exists():
                            self.all_data.append({
                                'type': 'read',
                                'id': exam_id,
                                'content': f"{exam_type_name}",
                                'analyze': '',
                                'audio': str(content_file),
                                'directory': str(dir_path),
                            })
                    elif exam_type_collect in ['collector.role', 'collector.dialogue']:
                        # 听选信息和回答问题
                        for i in range(1, 5):  # 最多4个问题
                            audio_file = dir_path / "material" / f"ques{i}askaudio.mp3"
                            if audio_file.exists():
                                self.all_data.append({
                                    'type': 'dialogue',
                                    'id': f"{exam_id}_{i}",
                                    'question': f"{exam_type_name} 问题 {i}",
                                    'listening_text': '',
                                    'standard_answers': [],
                                    'keywords': '',
                                    'audio': str(audio_file),
                                    'directory': str(dir_path),
                                })
                    elif exam_type_collect == 'collector.picture':
                        # 信息转述
                        content_file = dir_path / "material" / "content.mp3"
                        if content_file.exists():
                            self.all_data.append({
                                'type': 'picture',
                                'id': exam_id,
                                'content': f"{exam_type_name}",
                                'topic': '',
                                'keypoint': '',
                                'analyze': '',
                                'image': '',
                                'audio': str(content_file),
                                'directory': str(dir_path),
                            })
        except Exception as e:
            print(f"{self.RED}❌ 解析题库失败（{dir_path}）: {e}{self.NC}")

    def _parse_content_file(self, file_path: Path):
        # 解析content.json文件
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except json.JSONDecodeError as e:
            print(f"{self.RED}❌ JSON 格式错误（{file_path}）: {e}{self.NC}")
            return
        except Exception as e:
            print(f"{self.RED}❌ 读取文件失败（{file_path}）: {e}{self.NC}")
            return

        if not isinstance(data, dict):
            print(f"{self.RED}❌ content.json 根节点非对象: {file_path}{self.NC}")
            return

        stype = data.get('structure_type')
        info = data.get('info')
        if stype is None or info is None:
            print(f"{self.YELLOW}⚠️  content.json 缺少 structure_type 或 info: {file_path}{self.NC}")
            return

        dir_path = file_path.parent
        handlers = {
            'collector.dialogue': self._parse_dialogue_data,
            'collector.read': self._parse_read_data,
            'collector.choose': self._parse_choose_data,
            'collector.fill': self._parse_fill_data,
            'collector.picture': self._parse_picture_data,
            'collector.role': self._parse_dialogue_data,  # 添加对collector.role类型的支持
        }
        handler = handlers.get(stype)
        if handler:
            handler(dir_path, info)
        else:
            print(f"{self.YELLOW}⚠️  未知 structure_type: {stype} in {file_path}{self.NC}")

    def _safe_get_audio(self, dir_path: Path, audio_name):
        # 安全获取音频文件路径
        if not isinstance(audio_name, str) or not audio_name:
            return ''
        material_path = dir_path / "material" / audio_name
        return str(material_path) if material_path.is_file() else ''

    def _parse_dialogue_data(self, dir_path: Path, info):
        # 处理对话类型数据
        if not isinstance(info, dict):
            return
        questions = info.get('question')  # 获取问题列表
        if not isinstance(questions, list):
            return
        for q in questions:
            if not isinstance(q, dict):
                continue
            # 改进：确保从std字段提取标准答案，并处理可能的空情况
            std_answers = []
            std_list = q.get('std')
            if isinstance(std_list, list):
                std_answers = [std.get('value', '').strip() for std in std_list if isinstance(std, dict) and std.get('value', '').strip()]
            
            # 获取问题文本，优先使用ask字段，清理HTML标签和占位符
            ask_text = q.get('ask', '')
            question_text = self._clean_html(ask_text)
            
            # 从问题文本中提取纯问题部分（移除选项）
            # 查找括号内的选项部分并移除
            clean_question = question_text
            if '(' in question_text and ')' in question_text:
                # 尝试提取问题部分
                parts = question_text.split('(')
                if parts:
                    clean_question = parts[0].strip()
            
            # 获取音频文件路径
            audio = self._safe_get_audio(dir_path, q.get('askaudio'))
            
            # 添加到数据列表中
            self.all_data.append({
                'type': 'dialogue',
                'id': q.get('xh', ''),
                'question': clean_question,  # 清理后的问题
                'listening_text': question_text,  # 完整的听力原文（包含选项）
                'standard_answers': std_answers,
                'keywords': q.get('keywords', ''),
                'audio': audio,
                'directory': str(dir_path),
            })

    def _parse_read_data(self, dir_path: Path, info):
        # 处理阅读类型
        if not isinstance(info, dict):
            return
        audio = self._safe_get_audio(dir_path, info.get('audio'))
        # 直接构造数据
        self.all_data.append({
            'type': 'read',
            'id': info.get('stid', ''),
            'content': self._clean_html(info.get('value', '')),
            'analyze': self._clean_html(info.get('analyze', '')),
            'audio': audio,
            'directory': str(dir_path),
        })

    def _parse_choose_data(self, dir_path: Path, info):
        if not isinstance(info, dict):
            return
        dialogue = self._clean_html(info.get('st_nr', ''))
        audio = self._safe_get_audio(dir_path, info.get('audio'))
        xtlist = info.get('xtlist')  # 题目列表
        if not isinstance(xtlist, list):
            return
        for q in xtlist:
            if not isinstance(q, dict):
                continue
            xxlist = q.get('xxlist')  # 选项列表
            if not isinstance(xxlist, list):
                options = []
            else:
                options = [f"{opt.get('xx_mc', '')}. {opt.get('xx_nr', '').strip()}"
                           for opt in xxlist if isinstance(opt, dict)]
            # 这里要注意移除ets_th占位符
            question_text = q.get('xt_nr', '').replace('ets_th1', '').replace('ets_th2', '').strip()
            self.all_data.append({
                'type': 'choose',
                'id': q.get('xt_xh', ''),
                'dialogue': dialogue,
                'question': self._clean_html(question_text),
                'options': options,
                'answer': q.get('answer', ''),
                'analyze': self._clean_html(q.get('xt_analy', '')),
                'audio': audio,
                'directory': str(dir_path),
            })

    def _parse_fill_data(self, dir_path: Path, info):
        if not isinstance(info, dict):
            return
        audio = self._safe_get_audio(dir_path, info.get('audio'))
        std_list = info.get('std')  # 标准答案列表
        if not isinstance(std_list, list):
            answers = []
        else:
            answers = [{'number': item.get('th', ''), 'value': item.get('value', '')}
                       for item in std_list if isinstance(item, dict)]
        self.all_data.append({
            'type': 'fill',
            'id': info.get('stid', ''),
            'content': self._clean_html(info.get('value', '')),
            'answers': answers,
            'keypoint': self._clean_html(info.get('keypoint', '')),
            'audio': audio,
            'directory': str(dir_path),
        })

    def _parse_picture_data(self, dir_path: Path, info):
        # 图片题处理，需要处理图片路径
        if not isinstance(info, dict):
            return
        audio = self._safe_get_audio(dir_path, info.get('audio'))
        image_name = info.get('image')
        image_path = dir_path / "material" / image_name if image_name else None
        image = str(image_path) if image_path and image_path.is_file() else ''
        self.all_data.append({
            'type': 'picture',
            'id': info.get('stid', ''),
            'content': self._clean_html(info.get('value', '')),
            'topic': self._clean_html(info.get('topic', '')),
            'keypoint': self._clean_html(info.get('keypoint', '')),
            'analyze': self._clean_html(info.get('analyze', '')),
            'image': image,
            'audio': audio,
            'directory': str(dir_path),
        })

    def _clean_html(self, text):
        # 清理HTML标签和特殊字符
        if not isinstance(text, str):
            return ""
        # 移除标签
        text = re.sub(r'<[^>]+>', '', text)
        # 替换实体字符
        replacements = {
            '&nbsp;': ' ',
            '&amp;': '&',
            '&quot;': '"',
            '<': '<',
            '>': '>',
        }
        for k, v in replacements.items():
            text = text.replace(k, v)
        # 处理换行
        text = re.sub(r'</?br\s*/?>', '\n', text, flags=re.IGNORECASE)
        # 清理占位符
        text = re.sub(r'ets_th\d+', '', text)
        # 合并空格
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    # --------------------------下面是各种打印方法

    def _print_choose(self, q):
        if q.get('dialogue'):
            print(f"{self.BLUE}对话原文:{self.NC}\n{self.WHITE}{q['dialogue']}{self.NC}\n")
        print(f"{self.CYAN}题目 {q.get('id', '')}:{self.NC} {q.get('question', '')}")
        print(f"{self.GREEN}正确答案: {q.get('answer', '')}{self.NC}")
        print(f"{self.YELLOW}选项:{self.NC}")
        for opt in q.get('options', []):
            print(f"  {opt}")
        print()

    def _print_dialogue(self, q):
        print(f"{self.CYAN}问题 {q.get('id', '')}:{self.NC} {q.get('question', '')}")
        if q.get('listening_text'):
            print(f"{self.BLUE}听力原文:{self.NC} {q.get('listening_text', '')}")
        if q.get('standard_answers'):
            print(f"{self.GREEN}标准答案:{self.NC} {'; '.join(q['standard_answers'])}")
        if q.get('keywords'):
            print(f"{self.YELLOW}关键词:{self.NC} {q.get('keywords', '')}")
        print()

    def _print_read(self, q):
        print(f"{self.ORANGE}阅读内容:{self.NC}\n{self.WHITE}{q.get('content', '')}{self.NC}\n")

    def _print_fill(self, q):
        print(f"{self.MAGENTA}填空题:{self.NC}")
        print(f"{self.BLUE}原文:{self.NC}\n{self.WHITE}{q.get('content', '')}{self.NC}\n")
        print(f"{self.YELLOW}填空答案:{self.NC}")
        for i, ans in enumerate(q.get('answers', []), 1):
            print(f"{self.CYAN}空 {i} (题号{ans.get('number', '')}):{self.NC} {ans.get('value', '')}")
        print()

    def _print_picture(self, q):
        topic = q.get('topic', '')
        content = q.get('content', '')
        keypoint = q.get('keypoint', '')
        print(f"{self.ORANGE}主题: {topic}{self.NC}")
        print(f"{self.BLUE}内容:{self.NC}\n{self.WHITE}{content}{self.NC}\n")
        if keypoint:
            print(f"{self.YELLOW}关键点:{self.NC}\n{self.WHITE}{keypoint}{self.NC}\n")
        print()

    # 搜索功能

    def search_questions(self, keyword):
        found = False
        k = keyword.lower()
        for item in self.all_data:
            match = False
            if item['type'] == 'choose':
                if (k in item.get('question', '').lower() or
                    k in item.get('dialogue', '').lower() or
                    k in item.get('analyze', '').lower()):
                    match = True
            elif item['type'] == 'dialogue':
                if (k in item.get('question', '').lower() or
                    k in item.get('listening_text', '').lower() or
                    any(k in ans.lower() for ans in item.get('standard_answers', [])) or
                    k in item.get('keywords', '').lower()):
                    match = True
            elif item['type'] in ['read', 'fill', 'picture']:
                # 把所有文本字段拼起来搜索
                text_parts = []
                for v in item.values():
                    if isinstance(v, str):
                        text_parts.append(v)
                    elif isinstance(v, list):
                        for sub in v:
                            if isinstance(sub, dict):
                                text_parts.extend(str(val) for val in sub.values() if isinstance(val, str))
                            elif isinstance(sub, str):
                                text_parts.append(sub)
                full_text = ' '.join(text_parts)
                if k in full_text.lower():
                    match = True
            if match:
                found = True
                if item['type'] == 'choose':
                    self._print_choose(item)
                elif item['type'] == 'dialogue':
                    self._print_dialogue(item)
                elif item['type'] == 'read':
                    self._print_read(item)
                elif item['type'] == 'fill':
                    self._print_fill(item)
                elif item['type'] == 'picture':
                    self._print_picture(item)
        if not found:
            print(f"{self.RED}❌ 未找到包含 \"{keyword}\" 的题目。{self.NC}")

    def interactive_mode(self):
        total = len(self.all_data)
        if total == 0:
            print(f"{self.RED}❌ 未在目录 {self.root_dir} 中找到任何题目数据！{self.NC}")
            return
        print(f"{self.GREEN}✅ 成功加载 {total} 条题目！{self.NC}")
        print(f"{self.CYAN}🔍 输入关键词搜索题目，输入 {self.RED}/exit{self.CYAN} 退出。{self.NC}\n")
        while True:
            try:
                user_input = input("请输入: ").strip()
                if user_input.lower() in ['/exit', 'quit', 'q']:
                    print(f"{self.GREEN}再见！{self.NC}")
                    break
                elif user_input:
                    self.search_questions(user_input)
                else:
                    print(f"{self.YELLOW}⚠️  请输入关键词或命令。{self.NC}")
            except KeyboardInterrupt:
                # 防止嵌套的KeyboardInterrupt异常
                try:
                    print(f"\n{self.YELLOW}⚠️  已取消，输入 /exit 退出。{self.NC}")
                except KeyboardInterrupt:
                    # 如果再次收到中断信号，就直接退出程序
                    print(f"\n{self.GREEN}程序已退出。{self.NC}")
                    break
            except Exception as e:
                print(f"{self.RED}❌ 错误: {e}{self.NC}")

    # -----------GUI相关的搜索方法
    
    def search_questions_for_gui(self, keyword):
        # 返回给GUI的搜索结果
        results = []
        k = keyword.lower()
        for item in self.all_data:
            match = False
            if item['type'] == 'choose':
                if (k in item.get('question', '').lower() or
                    k in item.get('dialogue', '').lower() or
                    k in item.get('analyze', '').lower()):
                    match = True
            elif item['type'] == 'dialogue':
                if (k in item.get('question', '').lower() or
                    k in item.get('listening_text', '').lower() or
                    any(k in ans.lower() for ans in item.get('standard_answers', [])) or
                    k in item.get('keywords', '').lower()):
                    match = True
            elif item['type'] in ['read', 'fill', 'picture']:
                text_parts = []
                for v in item.values():
                    if isinstance(v, str):
                        text_parts.append(v)
                    elif isinstance(v, list):
                        for sub in v:
                            if isinstance(sub, dict):
                                text_parts.extend(str(val) for val in sub.values() if isinstance(val, str))
                            elif isinstance(sub, str):
                                text_parts.append(sub)
                full_text = ' '.join(text_parts)
                if k in full_text.lower():
                    match = True
            if match:
                results.append(item)
        return results

    def format_item_for_gui(self, item):
        # 转成HTML格式给GUI显示
        if item['type'] == 'choose':
            html = "<div style='margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px;'>"
            if item.get('dialogue'):
                html += f"<div style='color: #0d6efd; font-weight: bold; margin-bottom: 8px;'>对话原文:</div>"
                html += f"<div style='color: #495057; margin-bottom: 12px;'>{item['dialogue']}</div>"
            html += f"<div style='color: #0dcaf0; font-weight: bold; margin-bottom: 8px;'>题目 {item.get('id', '')}:</div>"
            html += f"<div style='color: #212529; margin-bottom: 12px;'>{item.get('question', '')}</div>"
            html += f"<div style='color: #198754; font-weight: bold; margin-bottom: 8px;'>正确答案: {item.get('answer', '')}</div>"
            html += f"<div style='color: #ffc107; font-weight: bold; margin-bottom: 8px;'>选项:</div>"
            for opt in item.get('options', []):
                html += f"<div style='color: #6c757d; margin-left: 20px; margin-bottom: 4px;'>{opt}</div>"
            html += "</div>"
            return html
            
        elif item['type'] == 'dialogue':
            html = "<div style='margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px;'>"
            html += f"<div style='color: #0dcaf0; font-weight: bold; margin-bottom: 8px;'>问题 {item.get('id', '')}:</div>"
            html += f"<div style='color: #212529; margin-bottom: 12px;'>{item.get('question', '')}</div>"
            if item.get('listening_text'):
                html += f"<div style='color: #0d6efd; font-weight: bold; margin-bottom: 8px;'>听力原文:</div>"
                html += f"<div style='color: #495057; margin-bottom: 12px;'>{item.get('listening_text', '')}</div>"
            if item.get('standard_answers'):
                html += f"<div style='color: #198754; font-weight: bold; margin-bottom: 8px;'>标准答案:</div>"
                html += f"<div style='color: #212529; margin-bottom: 12px;'>{'; '.join(item['standard_answers'])}</div>"
            if item.get('keywords'):
                html += f"<div style='color: #ffc107; font-weight: bold; margin-bottom: 8px;'>关键词:</div>"
                html += f"<div style='color: #212529; margin-bottom: 12px;'>{item.get('keywords', '')}</div>"
            html += "</div>"
            return html
            
        elif item['type'] == 'read':
            html = "<div style='margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px;'>"
            html += f"<div style='color: #fd7e14; font-weight: bold; margin-bottom: 8px;'>阅读内容:</div>"
            html += f"<div style='color: #212529; margin-bottom: 12px;'>{item.get('content', '')}</div>"
            html += "</div>"
            return html
            
        elif item['type'] == 'fill':
            html = "<div style='margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px;'>"
            html += f"<div style='color: #d63384; font-weight: bold; margin-bottom: 8px;'>填空题:</div>"
            html += f"<div style='color: #0d6efd; font-weight: bold; margin-bottom: 8px;'>原文:</div>"
            html += f"<div style='color: #212529; margin-bottom: 12px;'>{item.get('content', '')}</div>"
            html += f"<div style='color: #ffc107; font-weight: bold; margin-bottom: 8px;'>填空答案:</div>"
            for i, ans in enumerate(item.get('answers', []), 1):
                html += f"<div style='color: #0dcaf0; margin-left: 20px; margin-bottom: 4px;'>空 {i} (题号{ans.get('number', '')}): {ans.get('value', '')}</div>"
            html += "</div>"
            return html
            
        elif item['type'] == 'picture':
            html = "<div style='margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px;'>"
            topic = item.get('topic', '')
            content = item.get('content', '')
            keypoint = item.get('keypoint', '')
            if topic:
                html += f"<div style='color: #fd7e14; font-weight: bold; margin-bottom: 8px;'>主题: {topic}</div>"
            html += f"<div style='color: #0d6efd; font-weight: bold; margin-bottom: 8px;'>内容:</div>"
            html += f"<div style='color: #212529; margin-bottom: 12px;'>{content}</div>"
            if keypoint:
                html += f"<div style='color: #ffc107; font-weight: bold; margin-bottom: 8px;'>关键点:</div>"
                html += f"<div style='color: #212529; margin-bottom: 12px;'>{keypoint}</div>"
            html += "</div>"
            return html
        
        return ""

# 自定义悬浮窗 - 只在Windows系统且PyQt5可用时定义
if is_win and PYQT_AVAILABLE:
    class 自定义悬浮窗(QMainWindow):
        def __init__(self, extractor):
            super().__init__()
            self.extractor = extractor
            self.dragging = False
            self.drag_position = QPoint()
            self.initUI()
            
        def initUI(self):
            # 设置窗口属性
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.resize(600, 500)
            
            # 创建主容器
            central_widget = QWidget()
            central_widget.setObjectName("centralWidget")
            self.setCentralWidget(central_widget)
            
            # 主布局
            main_layout = QVBoxLayout(central_widget)
            main_layout.setContentsMargins(20, 20, 20, 20)
            main_layout.setSpacing(15)
            
            # 顶部控制区域
            top_layout = QHBoxLayout()
            top_layout.setContentsMargins(0, 0, 0, 0)
            top_layout.setSpacing(10)
            
            # 搜索输入框
            self.search_input = QLineEdit()
            self.search_input.setPlaceholderText("输入关键词搜索题目...")
            self.search_input.setFixedHeight(40)
            self.search_input.setStyleSheet("""
                QLineEdit {
                    border: 2px solid #e9ecef;
                    border-radius: 20px;
                    padding: 0 15px;
                    font-size: 14px;
                    background-color: white;
                }
                QLineEdit:focus {
                    border-color: #0d6efd;
                    outline: none;
                }
            """)
            
            # 搜索按钮
            search_button = QPushButton("搜索")
            search_button.setFixedHeight(40)
            search_button.setFixedWidth(80)
            search_button.setStyleSheet("""
                QPushButton {
                    background-color: #0d6efd;
                    color: white;
                    border: none;
                    border-radius: 20px;
                    font-weight: bold;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #0b5ed7;
                }
                QPushButton:pressed {
                    background-color: #0a58ca;
                }
            """)
            search_button.clicked.connect(self.perform_search)
            
            # 退出程序按钮
            exit_button = QPushButton("退出程序")
            exit_button.setFixedHeight(40)
            exit_button.setFixedWidth(100)
            exit_button.setStyleSheet("""
                QPushButton {
                    background-color: #dc3545;
                    color: white;
                    border: none;
                    border-radius: 20px;
                    font-weight: bold;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #c82333;
                }
                QPushButton:pressed {
                    background-color: #bd2130;
                }
            """)
            exit_button.clicked.connect(self.exit_program)
            
            top_layout.addWidget(self.search_input, 1)
            top_layout.addWidget(search_button)
            top_layout.addWidget(exit_button)
            
            # 结果显示区域
            self.result_display = QTextEdit()
            self.result_display.setReadOnly(True)
            self.result_display.setStyleSheet("""
                QTextEdit {
                    border: 2px solid #e9ecef;
                    border-radius: 15px;
                    padding: 15px;
                    font-size: 13px;
                    line-height: 1.6;
                    background-color: white;
                }
            """)
            self.result_display.setHtml("<div style='color: #6c757d; text-align: center; padding: 20px;'>请输入关键词开始搜索...</div>")
            
            # 添加到主布局
            main_layout.addLayout(top_layout)
            main_layout.addWidget(self.result_display, 1)
            
            # 设置样式表
            self.setStyleSheet("""
                QMainWindow {
                    background-color: transparent;
                }
                #centralWidget {
                    background-color: white;
                    border-radius: 20px;
                    border: 1px solid #dee2e6;
                }
            """)
            
            # 连接回车键事件
            self.search_input.returnPressed.connect(self.perform_search)
            
        def perform_search(self):
            keyword = self.search_input.text().strip()
            if not keyword:
                self.result_display.setHtml("<div style='color: #dc3545; text-align: center; padding: 20px;'>请输入搜索关键词</div>")
                return
                
            results = self.extractor.search_questions_for_gui(keyword)
            if not results:
                self.result_display.setHtml(f"<div style='color: #dc3545; text-align: center; padding: 20px;'>未找到包含 \"{keyword}\" 的题目</div>")
                return
                
            html_content = ""
            for item in results:
                html_content += self.extractor.format_item_for_gui(item)
            self.result_display.setHtml(html_content)
            
        def exit_program(self):
        # 退出程序
            sys.exit(0)
            
        def mousePressEvent(self, event):
            if event.button() == Qt.LeftButton:
                self.dragging = True
                self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
                event.accept()
                
        def mouseMoveEvent(self, event):
            if self.dragging:
                self.move(event.globalPos() - self.drag_position)
                event.accept()
                
        def mouseReleaseEvent(self, event):
            self.dragging = False
            
        def paintEvent(self, event):
            # 绘制圆角背景（移除了自定义绘制，使用样式表实现）
            pass

def 获取默认路径():
    # 获取Windows默认ETS路径
    try:
        # 直接返回已知的正确路径格式
        default_path = os.path.join("C:", "Users", os.getlogin(), "AppData", "Roaming", "ETS")
        # 如果路径不存在，尝试使用另一种方式获取用户名
        if not os.path.exists(default_path):
            try:
                import getpass
                username = getpass.getuser()
                default_path = os.path.join("C:", "Users", username, "AppData", "Roaming", "ETS")
            except:
                # 如果都失败，返回一个更通用的路径
                default_path = os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")))
        return default_path
    except Exception:
        return os.path.expanduser("~")

# 主程序
if __name__ == "__main__":
    # 检测是否强制使用控制台模式，同时支持 -console 和 --console
    use_console = any(flag in sys.argv for flag in ['-console', '--console'])
    if use_console:
        # 移除所有控制台相关的参数
        sys.argv = [arg for arg in sys.argv if arg not in ['-console', '--console']]
        print("🔄 使用控制台模式运行")
    
    # 判断是否在 Windows 系统上运行
    is_windows = platform.system() == 'Windows'
    
    BUILTIN_PATH = "/storage/emulated/0/Android/data/com.ets100.secondary/files/Download/ETS_secondary/resource/"
    root_dir = ""
    
    print("ETS 听说考试搜题工具")
    
    # 如果是 Windows 系统，弹出文件选择窗口
    if is_windows and tk_fd is not None:
        try:
            print("正在打开文件选择窗口...")
            default_path = 获取默认路径()
            root_dir = tk_fd.askdirectory(title="选择 ETS 题库目录", initialdir=default_path)
            if not root_dir:
                # 用户取消选择
                print("用户取消了选择，将使用命令行输入方式。")
                print(f"默认路径: {BUILTIN_PATH}")
                root_dir = input("目录路径: ").strip().strip('"')
                # 如果用户直接回车，使用内置路径
                if not root_dir:
                    root_dir = BUILTIN_PATH
        except Exception as e:
            print(f"打开文件选择窗口失败: {e}")
            print(f"默认路径: {BUILTIN_PATH}")
            root_dir = input("目录路径: ").strip().strip('"')
            # 如果用户直接回车，使用内置路径
            if not root_dir:
                root_dir = BUILTIN_PATH
    else:
        # 非 Windows 系统或tk_filedialog不可用，使用原有的命令行输入方式
        print("请输入试卷所在目录路径，直接回车使用默认路径（推荐）：")
        print(f"默认路径: {BUILTIN_PATH}")
        root_dir = input("目录路径: ").strip().strip('"')
        # 如果用户直接回车，使用内置路径
        if not root_dir:
            root_dir = BUILTIN_PATH

    try:
        extractor = ETS数据提取器(root_dir)
        
        # 检查是否为 Windows 且 PyQt5 可用
        if is_windows and PYQT_AVAILABLE and '自定义悬浮窗' in globals():
            # Windows系统且PyQt5可用，询问用户使用GUI还是控制台
            if not use_console:
                while True:
                    try:
                        choice = input("请选择运行模式 (y: GUI, n: 控制台): ").strip().lower()
                        if choice == 'y':
                            app = QApplication(sys.argv)
                            # 设置应用样式
                            app.setStyle('Fusion')
                            
                            window = 自定义悬浮窗(extractor)
                            window.show()
                            sys.exit(app.exec_())
                        elif choice == 'n':
                            break  # 使用控制台模式
                        else:
                            print("请输入 y 或 n。")
                    except KeyboardInterrupt:
                        print("\n操作已取消。")
                        break
            # 使用控制台模式
            extractor.interactive_mode()
        else:
            # 非Windows系统或PyQt5不可用，直接使用控制台模式
            if not is_windows:
                print("非 Windows 系统，使用控制台模式")
            elif not PYQT_AVAILABLE:
                print("PyQt5 未安装，使用控制台模式（可通过 pip install PyQt5 安装）")
            elif '自定义悬浮窗' not in globals():
                print("GUI 窗口类未定义，使用控制台模式")
            extractor.interactive_mode()
            
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        sys.exit(1)