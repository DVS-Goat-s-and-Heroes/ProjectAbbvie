import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import win32com.client as win32
import pythoncom # Importante para evitar erros de thread

class DocFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DocFlow Pro - Desktop")
        self.root.geometry("700x600")

        self.files_data = []
        
        style = ttk.Style()
        style.configure("Bold.TLabel", font=("Segoe UI", 9, "bold"))

        # --- Inputs ---
        input_frame = ttk.LabelFrame(root, text="Informações do Processo", padding=10)
        input_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(input_frame, text="Referência:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_ref = ttk.Entry(input_frame)
        self.entry_ref.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Número do PO:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.entry_po = ttk.Entry(input_frame)
        self.entry_po.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Ref. Cliente:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_cli = ttk.Entry(input_frame)
        self.entry_cli.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)

        # --- Lista ---
        list_frame = ttk.LabelFrame(root, text="Arquivos Selecionados", padding=10)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.canvas = tk.Canvas(list_frame)
        self.scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # --- Botões ---
        btn_frame = ttk.Frame(root, padding=10)
        btn_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(btn_frame, text="Adicionar PDFs", command=self.add_files).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Limpar Lista", command=self.clear_list).pack(side="left", padx=5)
        self.btn_process = ttk.Button(btn_frame, text="Gerar Documento Word", command=self.generate_word)
        self.btn_process.pack(side="right", padx=5)

    def add_files(self):
        files = filedialog.askopenfilenames(title="Selecione os PDFs", filetypes=[("Arquivos PDF", "*.pdf")])
        if files:
            for path in files:
                filename = os.path.basename(path)
                cat_var = tk.StringVar(value="Outros")
                
                row = ttk.Frame(self.scrollable_frame)
                row.pack(fill="x", pady=2, padx=5)
                
                ttk.Label(row, text=filename, width=40).pack(side="left", padx=5)
                ttk.Combobox(row, textvariable=cat_var, values=["Fatura", "Capa de Faturamento", "DI", "Outros"], 
                             state="readonly", width=20).pack(side="left", padx=5)
                
                self.files_data.append({"path": path, "name": filename, "category_var": cat_var, "widget": row})

    def clear_list(self):
        for item in self.files_data:
            item["widget"].destroy()
        self.files_data = []

    def sanitize_filename(self, text):
        for char in '<>:"/\|?*':
            text = text.replace(char, '_')
        return text.strip()

    def generate_word(self):
        ref = self.entry_ref.get().strip()
        po = self.entry_po.get().strip()
        cli = self.entry_cli.get().strip()

        if not all([ref, po, cli]):
            messagebox.showerror("Erro", "Preencha todos os campos.")
            return
        if not self.files_data:
            messagebox.showerror("Erro", "Adicione arquivos.")
            return

        temp_files = []

        try:
            self.btn_process.config(state="disabled", text="Processando...")
            self.root.update()
            
            # 1. Inicializa Threads do Windows (CRUCIAL para evitar erros aleatórios)
            pythoncom.CoInitialize()

            # 2. Usa dynamic.Dispatch para evitar cache corrompido do pywin32
            word_app = win32.dynamic.Dispatch("Word.Application")
            word_app.Visible = False
            
            # 3. Cria o documento
            doc = word_app.Documents.Add()

            # 4. Usa .Range() em vez de .Content (Mais seguro)
            rng = doc.Range() 
            rng.Collapse(0) # Vai para o fim

            rng.InsertAfter(f"Referência: {ref}\nPO: {po}\nCliente: {cli}\n\n")
            rng.InsertParagraphAfter()

            for item in self.files_data:
                original_path = os.path.abspath(item["path"])
                category = item["category_var"].get()
                
                # Renomeação
                safe_name = f"{self.sanitize_filename(category)}_{self.sanitize_filename(ref)}_{self.sanitize_filename(cli)}.pdf"
                temp_path = os.path.join(os.path.dirname(original_path), safe_name)
                shutil.copy2(original_path, temp_path)
                temp_files.append(temp_path)

                # Move para o fim
                rng = doc.Range()
                rng.Collapse(0)

                try:
                    # Tenta anexar. Se o Word não gostar do Range explicito no parametro, usamos o Selection (plano B)
                    obj = rng.InlineShapes.AddOLEObject(
                        FileName=temp_path,
                        LinkToFile=False,
                        DisplayAsIcon=True,
                        IconLabel=safe_name,
                        Range=rng
                    )
                    rng.InsertParagraphAfter()
                    rng.InsertParagraphAfter()
                except Exception as e_ole:
                    rng.InsertAfter(f"[Erro no anexo: {safe_name}]")
                    print(f"Erro OLE: {e_ole}")

            # Salvar
            save_name = f"Processo_{self.sanitize_filename(ref)}_{self.sanitize_filename(po)}.docx"
            save_path = os.path.abspath(os.path.join(os.path.dirname(self.files_data[0]["path"]), save_name))
            
            doc.SaveAs(save_path)
            doc.Close(False)
            try: word_app.Quit()
            except: pass

            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{save_path}")
            
            # Limpa temporários
            for f in temp_files:
                try: os.remove(f)
                except: pass

            if messagebox.askyesno("Abrir", "Deseja abrir o arquivo?"):
                os.startfile(save_path)

        except Exception as e:
            # Mostra o erro técnico completo (repr) para facilitar debug
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro técnico:\n{repr(e)}")
            try:
                if 'doc' in locals() and doc: doc.Close(False)
                if 'word_app' in locals() and word_app: word_app.Quit()
            except: pass
        
        finally:
            self.btn_process.config(state="normal", text="Gerar Documento Word")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocFlowApp(root)
    root.mainloop()
