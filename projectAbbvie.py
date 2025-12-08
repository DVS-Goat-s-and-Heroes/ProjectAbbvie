import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import win32com.client as win32
import tempfile
import pythoncom # Importante adicionar para evitar erros de thread em algumas máquinas

# Tentar importar tkinterdnd2 para habilitar Drag & Drop.
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TKDND_AVAILABLE = True
except Exception:
    TKDND_AVAILABLE = False
    DND_FILES = None
    TkinterDnD = None


class DocFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DocFlow Pro - Desktop")
        self.root.geometry("820x600")

        # Variáveis de armazenamento
        self.files_data = []

        # --- Estilos ---
        style = ttk.Style()
        style.configure("Bold.TLabel", font=("Segoe UI", 9, "bold"))

        # ---------------------------
        # Área de inputs (Referência/PO/Cliente)
        # ---------------------------
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

        # ---------------------------
        # Lista de arquivos
        # ---------------------------
        list_frame = ttk.LabelFrame(
            root,
            text="Arquivos Selecionados (arraste arquivos aqui se disponível)",
            padding=10,
        )
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.canvas = tk.Canvas(list_frame)
        self.scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Drag & Drop
        if TKDND_AVAILABLE:
            try:
                self.canvas.drop_target_register(DND_FILES)
                self.canvas.dnd_bind("<<Drop>>", self.drop_files)
            except:
                pass

        # ---------------------------
        # Botões principais
        # ---------------------------
        btn_frame = ttk.Frame(root, padding=10)
        btn_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(btn_frame, text="Adicionar PDFs", command=self.add_files).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Limpar Lista", command=self.clear_list).pack(side="left", padx=5)

        self.btn_process = ttk.Button(btn_frame, text="Gerar Documento Word", command=self.generate_word)
        self.btn_process.pack(side="right", padx=5)

        if not TKDND_AVAILABLE:
            ttk.Label(btn_frame, text="(Drag&Drop requer tkinterdnd2 instalado)", foreground="gray").pack(
                side="left", padx=10
            )

    # ======================================================================
    # Funções de adicionar arquivos
    # ======================================================================

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Selecione os arquivos PDF",
            filetypes=[("Arquivos PDF", "*.pdf")],
        )
        if files:
            for path in files:
                self._add_file_programmatically(path)

    def _add_file_programmatically(self, path):
        if any(f["path"] == path for f in self.files_data):
            return

        filename = os.path.basename(path)
        cat_var = tk.StringVar(value="Outros")

        row_frame = ttk.Frame(self.scrollable_frame)
        row_frame.pack(fill="x", pady=2, padx=5)

        ttk.Label(row_frame, text=filename, width=60, anchor="w").pack(
            side="left", padx=5, fill="x", expand=True
        )

        categories = ["Fatura", "Capa de Faturamento", "DI", "Outros"]
        combo = ttk.Combobox(row_frame, textvariable=cat_var, values=categories, state="readonly", width=22)
        combo.pack(side="left", padx=5)

        ttk.Button(row_frame, text="X", width=3,
                   command=lambda rf=row_frame, p=path: self.remove_file(rf, p)).pack(side="right", padx=5)

        self.files_data.append({
            "path": path,
            "name": filename,
            "category_var": cat_var,
            "widget": row_frame,
            "combo_widget": combo
        })

    def remove_file(self, row_frame, path):
        try:
            row_frame.destroy()
        except:
            pass
        self.files_data = [f for f in self.files_data if f["path"] != path]

    def clear_list(self):
        for item in list(self.files_data):
            try:
                item["widget"].destroy()
            except:
                pass
        self.files_data = []

    # ======================================================================
    # Drag & Drop
    # ======================================================================

    def drop_files(self, event):
        try:
            paths = self.root.splitlist(event.data)
        except:
            paths = event.data.replace("{", "").replace("}", "").split()

        for p in paths:
            if p.lower().endswith(".pdf"):
                clean = p.strip().strip("{}").strip('"')
                if os.path.exists(clean):
                    self._add_file_programmatically(clean)

    # ======================================================================
    # Geração do Word
    # ======================================================================

    def generate_word(self):
        ref = self.entry_ref.get().strip()
        po = self.entry_po.get().strip()
        cli = self.entry_cli.get().strip()

        if not all([ref, po, cli]):
            messagebox.showerror("Erro", "Preencha Referência, PO e Cliente.")
            return

        if not self.files_data:
            messagebox.showerror("Erro", "Adicione pelo menos um PDF.")
            return

        try:
            self.btn_process.config(state="disabled", text="Processando...")
            self.root.update()

            pythoncom.CoInitialize() # Inicializa threads para segurança
            
            # Dispatch simples, mas se der erro "Add.Content", troque para win32.dynamic.Dispatch
            word_app = win32.Dispatch("Word.Application")
            word_app.Visible = False
            doc = word_app.Documents.Add()

            # Cabeçalho
            rng = doc.Content
            rng.InsertAfter(f"Referência: {ref}\n")
            rng.InsertAfter(f"PO: {po}\n")
            rng.InsertAfter(f"Cliente: {cli}\n\n")

            # Tempdir para cópias
            temp_dir = tempfile.mkdtemp(prefix="docflow_pdf_")

            safe = lambda s: "".join(
                c if c.isalnum() or c in ("-", "_") else "_" for c in str(s)
            )

            for item in self.files_data:
                original_pdf = os.path.abspath(item["path"])
                category = item["category_var"].get()

                category_clean = safe(category).strip("_")

                # cria nome sem duplicar _pdf ou .pdf
                new_name = f"{category_clean}_{ref}_{cli}_{po}"
                new_name = new_name.replace("_pdf", "")
                new_name = new_name.replace(".pdf", "")
                new_name += ".pdf"

                new_path = os.path.join(temp_dir, new_name)

                # Evita sobrescrever
                if os.path.exists(new_path):
                    base, ext = os.path.splitext(new_name)
                    counter = 1
                    while True:
                        attempt = os.path.join(temp_dir, f"{base}_{counter}{ext}")
                        if not os.path.exists(attempt):
                            new_path = attempt
                            break
                        counter += 1

                try:
                    shutil.copy2(original_pdf, new_path)
                except:
                    new_path = original_pdf

                rng = doc.Content
                rng.Collapse(0)

                try:
                    # --- CORREÇÃO DO ÍCONE E PREVIEW ---
                    # 1. DisplayAsIcon=True: Obriga a ser ícone, não preview.
                    # 2. IconLabel: Define o nome que vai aparecer embaixo.
                    # 3. ClassType: Removido ou mantido? Se mantiver "AcroExch.Document", 
                    #    ele força o ícone do Adobe. Se remover, usa o do sistema.
                    #    Vou manter o que estava no seu, mas adicionei DisplayAsIcon.
                    
                    obj = rng.InlineShapes.AddOLEObject(
                        ClassType="AcroExch.Document", # Se der erro em PCs sem Adobe, remova essa linha
                        FileName=new_path,
                        LinkToFile=False,
                        DisplayAsIcon=True, # <--- ESSENCIAL PARA EVITAR O "PREVIEW"
                        IconLabel=new_name, # <--- ESSENCIAL PARA O NOME APARECER
                        Range=rng,
                    )

                    rng.InsertParagraphAfter()
                    

                except Exception as e:
                    # Fallback caso falhe com ClassType fixo
                    try:
                        rng.InlineShapes.AddOLEObject(
                            FileName=new_path,
                            LinkToFile=False,
                            DisplayAsIcon=True,
                            IconLabel=new_name,
                            Range=rng
                        )
                    except:
                        rng.InsertAfter(f"[ERRO ao anexar {new_path}: {e}]")
                        rng.InsertParagraphAfter()

            default_dir = os.path.dirname(self.files_data[0]["path"])
            save_filename = f"Processo_{ref}_{po}.docx"
            save_path = os.path.join(default_dir, save_filename)

            # Evitar sobrescrever .docx existente
            if os.path.exists(save_path):
                base, ext = os.path.splitext(save_filename)
                counter = 1
                while True:
                    attempt = os.path.join(default_dir, f"{base}_{counter}{ext}")
                    if not os.path.exists(attempt):
                        save_path = attempt
                        break
                    counter += 1

            doc.SaveAs(save_path)
            doc.Close(False)
            word_app.Quit()
            
            # Limpa temporários
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

            messagebox.showinfo("Sucesso", f"Documento gerado em:\n{save_path}")

            if messagebox.askyesno("Abrir", "Deseja abrir o arquivo agora?"):
                os.startfile(save_path)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro crítico:\n{e}")
            try:
                if 'doc' in locals() and doc: doc.Close(False)
                if 'word_app' in locals() and word_app: word_app.Quit()
            except: pass

        finally:
            self.btn_process.config(state="normal", text="Gerar Documento Word")


if __name__ == "__main__":
    if TKDND_AVAILABLE and TkinterDnD is not None:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    app = DocFlowApp(root)
    root.mainloop()
