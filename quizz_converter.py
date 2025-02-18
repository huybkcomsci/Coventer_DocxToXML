import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import xml.sax.saxutils as saxutils
import re
from docx import Document
from docx.shared import RGBColor
import os

class QuizzConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Quiz Converter to Moodle XML")
        self.root.geometry("1200x600")
        
        self.create_widgets()
        self.input_file_path = ""
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input Section
        input_frame = ttk.LabelFrame(main_frame, text="Đầu vào")
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            input_frame, 
            text="Chọn file câu hỏi",
            command=self.browse_input_file
        ).pack(side=tk.LEFT, padx=5)
        
        self.lbl_input_file = ttk.Label(input_frame, text="Chưa chọn file")
        self.lbl_input_file.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Control Section
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            control_frame,
            text="Xem trước XML",
            command=self.preview_xml
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            control_frame,
            text="Chuyển đổi",
            command=self.convert_file,
            style="Primary.TButton"
        ).pack(side=tk.RIGHT, padx=5)
        
        # Log Console
        self.log = scrolledtext.ScrolledText(main_frame, height=10, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True)
        
        # Style Configuration
        style = ttk.Style()
        style.configure("Primary.TButton", foreground="Green", background="#99FF99")
    
    def browse_input_file(self):
        file_types = (
            ('Word files', '*.docx'),
            ('Text files', '*.txt'),
            ('All files', '*.*')
        )
        self.input_file_path = filedialog.askopenfilename(
            title="Chọn file câu hỏi",
            filetypes=file_types
        )
        if self.input_file_path:
            self.lbl_input_file.config(text=self.input_file_path)
            self.log_message(f"Đã chọn file: {self.input_file_path}")
    
    def log_message(self, message):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)
    
    def preview_xml(self):
        if not self.input_file_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn file đầu vào trước!")
            return
        
        try:
            questions = self.parse_file()
            xml_content = self.generate_xml_content(questions)
            
            preview_window = tk.Toplevel(self.root)
            preview_window.title("Xem trước XML")
            
            text_area = scrolledtext.ScrolledText(
                preview_window, 
                wrap=tk.WORD,
                width=80,
                height=25
            )
            text_area.pack(fill=tk.BOTH, expand=True)
            text_area.insert(tk.INSERT, xml_content)
            text_area.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tạo bản xem trước:\n{str(e)}")
    
    def parse_file(self):
        if self.input_file_path.endswith('.docx'):
            return self.parse_docx(self.input_file_path)
        else:
            with open(self.input_file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            return self.parse_text(content)
    
    def parse_docx(self, file_path):
        doc = Document(file_path)
        questions = []
        current_question = None
        answer_line = None
        
        all_text = '\n'.join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
        answer_match = re.search(r'(\d+[A-Z]+)+$', all_text)
        if answer_match:
            answer_line = answer_match.group()
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            
            if re.fullmatch(r'Câu hỏi \d+:.+', text, re.IGNORECASE):
                if current_question:
                    questions.append(current_question)
                
                match = re.match(r'Câu hỏi (\d+):\s*(.*)', text, re.IGNORECASE)
                current_question = {
                    'id': int(match.group(1)),
                    'text': match.group(2),
                    'options': [],
                    'correct': [],
                    'explicit_answer': False
                }
                continue
            
            if current_question:
                option_match = re.fullmatch(r'([A-Z])\.\s*.+', text, re.IGNORECASE)
                if option_match:
                    letter = option_match.group(1).upper()
                    option_text = re.sub(r'^[A-Z]\.\s*', '', text)
                    
                    is_correct = any(self.check_formatting(run) for run in para.runs)
                    if is_correct:
                        current_question['correct'].append(letter)
                    
                    current_question['options'].append((letter, option_text))
                    continue
                
                answer_match = re.match(r'Đáp án đúng:\s*([A-Z,\s]+)', text, re.IGNORECASE)
                if answer_match:
                    answers = re.findall(r'[A-Z]', answer_match.group(1).upper())
                    current_question['correct'] = answers
                    current_question['explicit_answer'] = True
        
        if current_question:
            questions.append(current_question)
        
        if answer_line:
            answer_dict = {}
            matches = re.finditer(r'(\d+)([A-Z]+)', answer_line)
            for match in matches:
                qid = int(match.group(1))
                answers = list(match.group(2).upper())
                answer_dict[qid] = answers
            
            for question in questions:
                if not question['explicit_answer'] and not question['correct']:
                    question['correct'] = answer_dict.get(question['id'], [])
        
        return questions
    
    def parse_text(self, content):
        answer_line = None
        lines = content.strip().split('\n')
        for line in reversed(lines):
            stripped_line = line.strip().upper()
            if re.match(r'^(\d+[A-Z]+)+$', stripped_line):
                answer_line = stripped_line
                break
        
        questions = []
        current_question = None
        
        for line in lines:
            line = line.strip()
            if answer_line and line.upper() == answer_line:
                continue
            
            question_match = re.match(r'^Câu hỏi (\d+):\s*(.+)$', line, re.IGNORECASE)
            if question_match:
                if current_question:
                    questions.append(current_question)
                current_question = {
                    'id': int(question_match.group(1)),
                    'text': question_match.group(2).strip(),
                    'options': [],
                    'correct': []
                }
                continue
            
            if current_question:
                option_match = re.match(r'^([A-Z])\.\s*(.+)$', line, re.IGNORECASE)
                if option_match:
                    letter = option_match.group(1).upper()
                    text = option_match.group(2).strip()
                    current_question['options'].append((letter, text))
                    continue
                
                answer_match = re.match(r'Đáp án đúng:\s*([A-Z,\s]+)', line, re.IGNORECASE)
                if answer_match:
                    answers = re.findall(r'[A-Z]', answer_match.group(1).upper())
                    current_question['correct'] = answers
        
        if current_question:
            questions.append(current_question)
        
        return questions
    
    def check_formatting(self, run):
        if run.font.color.rgb == RGBColor(255, 0, 0):
            return True
        if run.font.bold or run.font.italic or run.font.underline:
            return True
        return False
    
    def generate_xml_content(self, questions):
        xml_content = '<?xml version="1.0" encoding="UTF-8"?>\n<quiz>\n'
        
        for question in questions:
            xml_content += '    <question type="multichoice">\n'
            xml_content += '        <name>\n'
            xml_content += f'            <text>{saxutils.escape(question["text"])}</text>\n'
            xml_content += '        </name>\n'
            xml_content += '        <questiontext format="html">\n'
            xml_content += f'            <text>{saxutils.escape(question["text"])}</text>\n'
            xml_content += '        </questiontext>\n'
            
            sorted_options = sorted(question['options'], key=lambda x: x[0])
            num_correct = len(question['correct'])
            
            for letter, text in sorted_options:
                if num_correct > 0:
                    fraction = 100.0 / num_correct if letter in question['correct'] else 0.0
                else:
                    fraction = 0.0
                
                fraction_str = f"{fraction:.2f}".rstrip('0').rstrip('.') if fraction % 1 else f"{int(fraction)}"
                xml_content += f'        <answer fraction="{fraction_str}">\n'
                xml_content += f'            <text>{saxutils.escape(text)}</text>\n'
                xml_content += '        </answer>\n'
            
            xml_content += '    </question>\n'
        
        xml_content += '</quiz>'
        return xml_content
    
    def convert_file(self):
        if not self.input_file_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn file đầu vào!")
            return
        
        try:
            questions = self.parse_file()
            
            # Tạo thư mục OUTPUT
            output_dir = os.path.join(os.path.dirname(self.input_file_path), "OUTPUT")
            os.makedirs(output_dir, exist_ok=True)
            
            # Tạo tên file đầu ra
            base_name = os.path.splitext(os.path.basename(self.input_file_path))[0]
            output_path = os.path.join(output_dir, f"{base_name}_quiz.xml")
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(self.generate_xml_content(questions))
            
            messagebox.showinfo("Thành công", f"Đã tạo file XML thành công!\n{output_path}")
            self.log_message(f"Đã xuất file thành công: {output_path}")
            os.startfile(output_dir)
        
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi xử lý file:\n{str(e)}")
            self.log_message(f"Lỗi: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = QuizzConverter(root)
    root.mainloop()