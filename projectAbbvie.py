import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import win32com.client as win32

class DocFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DocFlow Pro - Desktop")
        self.root.geometry("700x600")

        # Variáveis de armazenamento
        self.files_data = [] # Lista de dicionários: {'path': str, 'name': str, 'category_var': tk.StringVar}
        
        # --- Estilos ---
        style = ttk.Style()
        style.configure("Bold.TLabel", font=("Segoe UI", 9, "bold"))

        # --- Área de Inputs Globais ---
        input_frame = ttk.LabelFrame(root, text="Informações do Processo", padding=10)
        input_frame.pack(fill="x", padx=10, pady=5)

        # Grid para inputs
        ttk.Label(input_frame, text="Referência:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_ref = ttk.Entry(input_frame)
        self.entry_ref.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Número do PO:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.entry_po = ttk.Entry(input_frame)
        self.entry_po.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Ref. Cliente:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_cli = ttk.Entry(input_frame)
        self.entry_cli.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Configura expansão das colunas
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)

        # --- Área de Lista de Arquivos ---
        list_frame = ttk.LabelFrame(root, text="Arquivos Selecionados", padding=10)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Canvas e Scrollbar para a lista de arquivos
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

        # --- Área de Botões ---
        btn_frame = ttk.Frame(root, padding=10)
        btn_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(btn_frame, text="Adicionar PDFs", command=self.add_files).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Limpar Lista", command=self.clear_list).pack(side="left", padx=5)
        
        self.btn_process = ttk.Button(btn_frame, text="Gerar Documento Word", command=self.generate_word)
        self.btn_process.pack(side="right", padx=5)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Selecione os arquivos PDF",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        
        if files:
            for path in files:
                filename = os.path.basename(path)
                
                # Cria variavel para guardar a categoria selecionada
                cat_var = tk.StringVar(value="Outros")
                
                # Cria container visual para a linha
                row_frame = ttk.Frame(self.scrollable_frame)
                row_frame.pack(fill="x", pady=2, padx=5)
                
                # Label do nome do arquivo
                lbl_name = ttk.Label(row_frame, text=filename, width=40)
                lbl_name.pack(side="left", padx=5)
                
                # Dropdown de Categoria
                categories = ["Fatura", "Capa de Faturamento", "DI", "Outros"]
                combo = ttk.Combobox(row_frame, textvariable=cat_var, values=categories, state="readonly", width=20)
                combo.pack(side="left", padx=5)
                
                # Salva os dados
                self.files_data.append({
                    "path": path,
                    "name": filename,
                    "category_var": cat_var,
                    "widget": row_frame # Guardamos a ref do widget se precisarmos deletar depois
                })

    def clear_list(self):
        for item in self.files_data:
            item["widget"].destroy()
        self.files_data = []

    def generate_word(self):
        # 1. Validação
        ref = self.entry_ref.get().strip()
        po = self.entry_po.get().strip()
        cli = self.entry_cli.get().strip()

        if not all([ref, po, cli]):
            messagebox.showerror("Erro", "Preencha todos os campos (Referência, PO e Cliente).")
            return
        
        if not self.files_data:
            messagebox.showerror("Erro", "Adicione pelo menos um arquivo PDF.")
            return

        # 2. Processamento
        try:
            self.btn_process.config(state="disabled", text="Processando...")
            self.root.update() # Atualiza a interface

            word_app = win32.Dispatch("Word.Application")
            word_app.Visible = False
            doc = word_app.Documents.Add()

            # --- Cabeçalho (Estilo Clean) ---
            # Adiciona as informações
            rng = doc.Content
            rng.Collapse(0) # Fim do doc
            rng.InsertAfter(f"Referência: {ref}\n")
            rng.InsertAfter(f"PO: {po}\n")
            rng.InsertAfter(f"Cliente: {cli}\n")
            rng.InsertParagraphAfter()
            
            # Adiciona uma linha separadora ou espaço
            rng.InsertParagraphAfter()

            # Itera sobre os arquivos na lista
            for item in self.files_data:
                pdf_path = os.path.abspath(item["path"])
                category = item["category_var"].get()
                
                # Label do ícone (Nome do arquivo PDF gerado)
                icon_label = f"{category[:3].upper()}_{ref}.pdf"

                # Move cursor para o fim
                rng = doc.Content
                rng.Collapse(0) # wdCollapseEnd
                
                # Adiciona o título da categoria (opcional, pode remover se quiser EXATAMENTE só o ícone)
                # rng.InsertAfter(f"{category}:") 
                # rng.InsertParagraphAfter()
                
                rng.Collapse(0)

                try:
                    # --- CORREÇÃO PRINCIPAL AQUI ---
                    # 1. Mudado ClassName -> ClassType
                    # 2. Removido IconFileName para usar o ícone padrão do sistema (igual imagem 2)
                    # 3. Adicionado Range no final para inserir exatamente na posição
                    obj = rng.InlineShapes.AddOLEObject(
                        ClassType="AcroExch.Document.DC", # Nome correto do parâmetro
                        FileName=pdf_path,
                        LinkToFile=False,
                        Range=rng
                    )
                    
                    # Tenta formatar o ícone (Opcional: Centralizar)
                    # obj.Range.ParagraphFormat.Alignment = 1 # 0=Left, 1=Center, 2=Right
                    
                    # Adiciona espaço após o ícone
                    rng.InsertParagraphAfter()
                    rng.InsertParagraphAfter()

                except Exception as e_ole:
                    # Se der erro, tenta avisar no documento
                    rng.InsertAfter(f"[ERRO AO ANEXAR {category}: {str(e_ole)}]")
                    rng.InsertParagraphAfter()
                    print(f"Erro OLE: {e_ole}")

            # Salvar
            save_filename = f"Processo_{ref}_{po}.docx"
            save_path = os.path.join(os.path.dirname(self.files_data[0]["path"]), save_filename)
            save_path = os.path.abspath(save_path) # Garante caminho absoluto
            
            doc.SaveAs(save_path)
            doc.Close(False)
            word_app.Quit()

            messagebox.showinfo("Sucesso", f"Arquivo gerado com sucesso em:\n{save_path}")
            
            # Pergunta se quer abrir
            if messagebox.askyesno("Abrir", "Deseja abrir o arquivo gerado agora?"):
                os.startfile(save_path)

        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro na automação:\n{str(e)}")
            # Tenta fechar o word se ficou aberto
            try:
                if 'doc' in locals() and doc: doc.Close(False)
                if 'word_app' in locals() and word_app: word_app.Quit()
            except:
                pass
        
        finally:
            self.btn_process.config(state="normal", text="Gerar Documento Word")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocFlowApp(root)
    root.mainloop()
