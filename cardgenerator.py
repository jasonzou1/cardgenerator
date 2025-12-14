import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import pandas as pd
import os
import json
import re
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
from openai import OpenAI

# ================= Configuration =================
class Config:
    SPLIT_ANCHOR = "RECIPIENT FULL ADDRESS" 
    COL_IDX_ADDRESS = 3  # Column D
    COL_IDX_MESSAGE = 8  # Column I
    
    # Garbage Filter (Skip these rows)
    IGNORE_KEYWORDS = [
        "recipient full address",
        "form instructions",
        "basket name",
        "delivery date",
        "special instructions",
        "must be a valid address"
    ]
    
    # User Blacklist (Do not generate cards for these)
    USER_BLACKLIST = [
        "750 millway", 
        "my baskets", 
        "unit #4"
    ]

# ================= Logic Class =================
class CardGenerator:
    def __init__(self, log_callback):
        self.log = log_callback
        self.api_key = None
        self.base_url = None
        self.model_name = "gpt-3.5-turbo"

    def update_settings(self, api_key, base_url, model_name):
        self.api_key = api_key
        self.base_url = base_url if base_url.strip() else "https://api.openai.com/v1"
        self.model_name = model_name if model_name.strip() else "gpt-3.5-turbo"

    def _is_garbage(self, text):
        t = text.lower()
        if not t.strip(): return True
        for kw in Config.IGNORE_KEYWORDS:
            if kw in t: return True
        return False

    def _looks_like_phone(self, text):
        t = text.lower().strip()
        if "tel" in t or "phone" in t: return True
        clean_nums = re.sub(r'[\d\-\(\)\.\+\s]', '', t)
        if len(clean_nums) == 0 and len(t) > 5: return True
        if len(t) > 0 and t[0].isdigit(): return True
        return False

    def _clean_labels(self, text):
        """
        Removes 'Phone:', 'TEL:', etc. from the text, keeping the numbers.
        """
        lines = text.split('\n')
        cleaned_lines = []
        for line in lines:
            # Regex to remove "Tel:", "Phone:", "Ph." at the start of the line (case insensitive)
            # Replaces "Phone: 123-456" with "123-456"
            new_line = re.sub(r'(?i)^\s*(tel|phone|ph|mobile|cell)[:\.]?\s*', '', line)
            if new_line.strip():
                cleaned_lines.append(new_line)
        return "\n".join(cleaned_lines)

    def ai_format_block(self, raw_text):
        if not self.api_key: return raw_text
        try:
            client = OpenAI(api_key=self.api_key, base_url=self.base_url)
            prompt = f"""
            Format this shipping address block.
            Input:
            ---
            {raw_text}
            ---
            Rules:
            1. STRICTLY REMOVE all labels like "TEL:", "Phone:", "Attention:". Just keep the value.
            2. Do not write "[insert phone number]". If no number exists, leave it blank.
            3. Put the phone number on the last line.
            4. Output plain text only.
            """
            response = client.chat.completions.create(
                model=self.model_name,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            self.log(f"AI Error: {e}")
            return raw_text

    def read_excel_strict(self, file_path):
        self.log(f"Reading file: {os.path.basename(file_path)}...")
        try:
            df = pd.read_excel(file_path, header=None, dtype=str)
            
            start_row = 0
            for r in range(len(df)):
                val = str(df.iloc[r, Config.COL_IDX_ADDRESS]).strip()
                if Config.SPLIT_ANCHOR.lower() in val.lower():
                    start_row = r
                    self.log(f"✅ Found anchor at Row {r+1}")
                    break

            parsed_cards = []
            current_addr_lines = []
            current_msg = ""
            empty_gap = 0
            
            # === Reading Loop ===
            for i in range(start_row + 1, len(df)):
                raw_addr = str(df.iloc[i, Config.COL_IDX_ADDRESS]).strip()
                if raw_addr.lower() == 'nan': raw_addr = ""
                
                raw_msg = str(df.iloc[i, Config.COL_IDX_MESSAGE]).strip()
                if raw_msg.lower() == 'nan': raw_msg = ""

                if self._is_garbage(raw_addr): continue

                is_new_person = False
                if raw_msg: 
                    is_new_person = True
                elif raw_addr:
                    if not self._looks_like_phone(raw_addr) and empty_gap >= 2:
                        is_new_person = True

                if is_new_person and current_addr_lines:
                    self._add_card(parsed_cards, current_addr_lines, current_msg)
                    current_addr_lines = []
                    current_msg = ""
                    empty_gap = 0

                if is_new_person:
                    current_addr_lines.append(raw_addr)
                    current_msg = raw_msg
                    empty_gap = 0
                else:
                    if raw_addr:
                        current_addr_lines.append(raw_addr)
                        if raw_msg and not current_msg: current_msg = raw_msg
                        empty_gap = 0
                    else:
                        empty_gap += 1

            if current_addr_lines:
                self._add_card(parsed_cards, current_addr_lines, current_msg)

            # === Final Cleaning (Remove short/empty addresses) ===
            final_cards = []
            for card in parsed_cards:
                addr_clean = card['address'].strip().replace('\n', '')
                if len(addr_clean) > 5:
                    final_cards.append(card)

            self.log(f"📊 Extraction Complete. Valid Cards: {len(final_cards)}")
            return final_cards

        except Exception as e:
            self.log(f"❌ Error Reading File: {e}")
            return []

    def _add_card(self, card_list, lines, msg):
        full_text = "\n".join(lines)
        
        # Blacklist check
        for bw in Config.USER_BLACKLIST:
            if bw in full_text.lower(): return
        
        # 1. Programmatic cleaning (Remove "Phone:" prefix)
        full_text = self._clean_labels(full_text)

        # 2. AI formatting (Optional)
        if self.api_key:
            full_text = self.ai_format_block(full_text)
            # Clean again in case AI added labels back (rare but possible)
            full_text = self._clean_labels(full_text)

        card_list.append({'address': full_text, 'message': msg})

    def generate_word(self, data_list, output_path):
        """Word Generation (English Logs + Arial + Size Limits)"""
        if not data_list: return

        self.log("Generating Word Document...")
        doc = docx.Document()
        section = doc.sections[0]
        section.page_height = Cm(29.7); section.page_width = Cm(21.0)
        
        margin = Cm(1.0)
        section.top_margin = margin; section.bottom_margin = margin
        section.left_margin = margin; section.right_margin = margin

        # Row Height: 6.75cm (Prevent Overflow)
        ROW_HEIGHT = Cm(6.75)
        COL_WIDTH = Cm(9.5)

        chunk_size = 4
        total_pages = (len(data_list) + chunk_size - 1) // chunk_size

        for page_idx in range(total_pages):
            start = page_idx * chunk_size
            end = start + chunk_size
            page_data = data_list[start:end]

            table = doc.add_table(rows=4, cols=2)
            table.autofit = False 
            
            for row in table.rows:
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                row.height = ROW_HEIGHT
                row.cells[0].width = COL_WIDTH
                row.cells[1].width = COL_WIDTH

            for i, item in enumerate(page_data):
                row = table.rows[i]
                
                # === Left: Address (Left Align, Arial, 10.5pt) ===
                cell_l = row.cells[0]
                cell_l.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p_l = cell_l.paragraphs[0]
                p_l.text = item['address']
                p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if p_l.runs:
                    run_l = p_l.runs[0]
                    run_l.font.name = 'Arial' 
                    run_l.font.size = Pt(10.5) 
                
                # === Right: Message (Center, Arial, 10-16pt) ===
                cell_r = row.cells[1]
                cell_r.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p_r = cell_r.paragraphs[0]
                p_r.text = item['message']
                p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if p_r.runs:
                    run_r = p_r.runs[0]
                    run_r.font.name = 'Arial'
                    
                    # Smart Sizing
                    length = len(item['message'])
                    if length < 30:
                        run_r.font.size = Pt(16) # Max
                    elif length < 80:
                        run_r.font.size = Pt(13) # Mid
                    else:
                        run_r.font.size = Pt(10) # Min

            if page_idx < total_pages - 1: 
                doc.add_page_break()
        
        try:
            doc.save(output_path)
            self.log(f"✅ Saved Successfully: {output_path}")
            messagebox.showinfo("Success", "Processing Complete!\n- Font: Arial\n- Labels Removed\n- Blank Pages Removed")
        except Exception as e:
            self.log(f"❌ Error Saving File: {e}")

# ================= UI Class =================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Card Generator (V9 English)")
        self.root.geometry("650x600")
        
        self.api_key = tk.StringVar()
        self.base_url = tk.StringVar()
        self.model_name = tk.StringVar(value="gpt-3.5-turbo") 
        self.file_path = tk.StringVar()
        
        self.logic = CardGenerator(self.log_msg)
        self._load_config()
        self._setup_ui()

    def _load_config(self):
        if os.path.exists("config_ai.json"):
            try:
                with open("config_ai.json", "r") as f:
                    c = json.load(f)
                    self.api_key.set(c.get("api_key", ""))
                    self.base_url.set(c.get("base_url", ""))
                    self.model_name.set(c.get("model_name", "gpt-3.5-turbo"))
            except: pass

    def _save_config(self):
        with open("config_ai.json", "w") as f:
            json.dump({
                "api_key": self.api_key.get(),
                "base_url": self.base_url.get(),
                "model_name": self.model_name.get()
            }, f)

    def _setup_ui(self):
        # AI Config
        f_ai = tk.LabelFrame(self.root, text="AI Settings (Optional)", padx=10, pady=10)
        f_ai.pack(fill="x", padx=10, pady=5)
        
        tk.Label(f_ai, text="API Key:").grid(row=0, column=0, sticky="e")
        tk.Entry(f_ai, textvariable=self.api_key, width=40, show="*").grid(row=0, column=1, padx=5, pady=2)
        tk.Label(f_ai, text="Base URL:").grid(row=1, column=0, sticky="e")
        tk.Entry(f_ai, textvariable=self.base_url, width=40).grid(row=1, column=1, padx=5, pady=2)
        tk.Label(f_ai, text="Model:").grid(row=2, column=0, sticky="e")
        tk.Entry(f_ai, textvariable=self.model_name, width=40).grid(row=2, column=1, padx=5, pady=2)

        # File
        f_file = tk.LabelFrame(self.root, text="File Selection", padx=10, pady=10)
        f_file.pack(fill="x", padx=10, pady=5)
        tk.Entry(f_file, textvariable=self.file_path, width=50).pack(side="left")
        tk.Button(f_file, text="Browse Excel", command=self.sel_file).pack(side="left", padx=5)

        # Run Button
        tk.Button(self.root, text="🚀 Start Processing", command=self.run_thread, 
                  bg="#28a745", fg="white", font=("Arial", 12, "bold"), height=2).pack(fill="x", padx=20, pady=10)

        # Log Area
        self.log_area = scrolledtext.ScrolledText(self.root, height=15)
        self.log_area.pack(fill="both", expand=True, padx=10, pady=5)

    def sel_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p: self.file_path.set(p)

    def log_msg(self, msg):
        self.root.after(0, lambda: self.log_area.insert(tk.END, f">> {msg}\n"))
        self.root.after(0, lambda: self.log_area.see(tk.END))

    def run_thread(self):
        self._save_config()
        threading.Thread(target=self.run, daemon=True).start()

    def run(self):
        f = self.file_path.get()
        if not f: return
        self.logic.update_settings(self.api_key.get(), self.base_url.get(), self.model_name.get())
        data = self.logic.read_excel_strict(f)
        if data:
            base = os.path.splitext(f)[0]
            out = f"{base}_English_V9.docx"
            self.logic.generate_word(data, out)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
