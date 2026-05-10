import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import xlwings as xw
from groq import Groq
import os
import re

class CA_Excel_Copilot:
    def __init__(self, root):
        self.root = root
        self.root.title("CA Copilot - OSS 120b")
        self.root.geometry("360x520")
        self.root.attributes('-topmost', True) 
        self.root.configure(padx=10, pady=10)

        # Hide the main window while asking for the key
        self.root.withdraw()
        
        # 1. Prompt for the API key on startup
        api_key = simpledialog.askstring(
            "API Key Required",
            "Enter your Groq API Key:",
            parent=self.root,
            show='*' # Masks the input for security
        )

        # Check if user cancelled or entered nothing
        if not api_key or not api_key.strip():
            messagebox.showerror("Auth Error", "API Key is required to run the Copilot.")
            self.root.destroy()
            return
            
        self.groq_client = Groq(api_key=api_key.strip())
        
        # Restore the main window
        self.root.deiconify()

        # 2. Build the Floating UI
        style = ttk.Style()
        style.configure('TButton', font=('Segoe UI', 9))
        
        ttk.Label(root, text="⚡ CA Excel Copilot", font=("Segoe UI", 12, "bold")).pack(pady=(0, 10))
        
        self.chat_display = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=20, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4")
        self.chat_display.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.log_system("System ready. Open an Excel file and type a command.")

        self.cmd_input = ttk.Entry(root, font=("Segoe UI", 10))
        self.cmd_input.pack(fill=tk.X, pady=(0, 5))
        self.cmd_input.bind("<Return>", lambda event: self.execute_command())
        
        self.run_btn = ttk.Button(root, text="Execute Logic", command=self.execute_command)
        self.run_btn.pack(fill=tk.X)

    def log_system(self, text):
        self.chat_display.insert(tk.END, f"[SYS] {text}\n\n")
        self.chat_display.see(tk.END)
        
    def log_user(self, text):
        self.chat_display.insert(tk.END, f"> {text}\n", "user")
        self.chat_display.tag_config("user", foreground="#569cd6")
        self.chat_display.see(tk.END)

    def log_action(self, text):
        self.chat_display.insert(tk.END, f"✓ {text}\n\n", "action")
        self.chat_display.tag_config("action", foreground="#4ec9b0")
        self.chat_display.see(tk.END)

    def get_excel_context(self):
        try:
            wb = xw.books.active
            sheet = wb.sheets.active
            data = sheet.used_range.options(ndim=2).value
            if not data:
                return None, None
            
            context_str = "Row 1 (Headers): " + " | ".join([str(x) for x in data[0]]) + "\n"
            context_str += "Data Sample:\n"
            for row in data[1:min(20, len(data))]:
                context_str += " | ".join([str(x) for x in row]) + "\n"
                
            return sheet, context_str
        except Exception as e:
            self.log_system(f"Could not connect to Excel. Make sure a workbook is open. Error: {e}")
            return None, None

    def execute_command(self):
        cmd = self.cmd_input.get().strip()
        if not cmd: return
        
        self.cmd_input.delete(0, tk.END)
        self.log_user(cmd)
        self.root.update()

        sheet, context = self.get_excel_context()
        if not context: return

        self.log_system("Thinking & writing code...")
        self.root.update()

        sys_prompt = f"""
        You are an advanced Python agent controlling Microsoft Excel via the 'xlwings' library.
        The user will give you a command. 
        You must output ONLY valid, executable Python code to fulfill the command.
        
        ENVIRONMENT:
        - The xlwings library is imported as 'xw'.
        - The active sheet is already defined in the execution environment as the variable 'sheet'.
        - Do NOT import xlwings or define 'sheet' in your code. Just write the operational logic.
        - Example: `sheet.range('A1').value = 'Test'`
        
        CURRENT EXCEL CONTEXT:
        {context}
        
        CRITICAL RULES:
        - Output ONLY python code.
        - Do not use markdown blocks (```python). Just raw text.
        - Do not explain yourself.
        """

        try:
            response = self.groq_client.chat.completions.create(
                model="openai/gpt-oss-120b",
                messages=[
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": cmd}
                ],
                temperature=0.1
            )
            
            ai_code = response.choices[0].message.content.strip()
            
            # Clean up markdown if the AI disobeys the prompt rules
            ai_code = re.sub(r"^```python\s*", "", ai_code, flags=re.MULTILINE)
            ai_code = re.sub(r"^```\s*", "", ai_code, flags=re.MULTILINE)

            # Execute the AI-generated code live on the machine
            exec(ai_code, {'xw': xw, 'sheet': sheet})
            
            self.log_action("Task completed in Excel.")
            
        except Exception as e:
            self.log_system(f"Execution failed:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CA_Excel_Copilot(root)
    root.mainloop()
