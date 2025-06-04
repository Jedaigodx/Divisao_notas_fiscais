import customtkinter as ctk
import pandas as pd
from tkinter import filedialog
import os
from datetime import datetime

# Configuração de aparência
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Janela principal
app = ctk.CTk()
app.title("Conversor de Mapas - PMGU/GNA_COPESP")
app.geometry("800x600")
app.resizable(True,True)

arquivo_mapa = ""
arquivo_inex = None
pasta_destino = os.path.join(os.path.expanduser("~"), "Downloads")

# Funções
def selecionar_arquivo():
    global arquivo_mapa
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
    if file_path:
        arquivo_mapa = file_path
        status.configure(text=f"📄 Arquivo selecionado: {os.path.basename(file_path)}", text_color="#1E3A8A", font=("Arial", 12, "bold"))

def selecionar_pasta():
    global pasta_destino
    folder = filedialog.askdirectory()
    if folder:
        pasta_destino = folder
        status.configure(text=f"📁 Pasta selecionada: {pasta_destino}", text_color="#1E3A8A", font=("Arial", 12, "bold"))

def selecionar_inex():
    global arquivo_inex
    file_path = filedialog.askopenfilename(title="Selecione o arquivo INEX", filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
    if file_path:
        arquivo_inex = file_path

def popup_incluir_inex():
    popup = ctk.CTkToplevel(app)
    popup.title("Incluir INEX?")
    popup.geometry("350x150")
    popup.grab_set()

    ctk.CTkLabel(popup, text="Deseja incluir o arquivo INEX?", font=("Arial", 14)).pack(pady=20)

    def incluir():
        popup.destroy()
        selecionar_inex()

    def nao_incluir():
        popup.destroy()

    botoes_frame = ctk.CTkFrame(popup, fg_color="transparent")
    botoes_frame.pack(pady=10)
    ctk.CTkButton(botoes_frame, text="✅ Sim", command=incluir, width=100).pack(side="left", padx=10)
    ctk.CTkButton(botoes_frame, text="❌ Não", command=nao_incluir, width=100).pack(side="right", padx=10)

def formatar_identificador(val):
    try:
        val_str = str(int(val)).zfill(14)
        stripped = str(int(val))
        if len(stripped) <= 11:
            cpf = val_str[-11:]
            return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
        else:
            cnpj = val_str
            return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    except:
        return val

def converter():
    try:
        progress.pack(fill="x", padx=30, pady=(10, 10))  
        progress.start()  
        app.update_idletasks() 

        mapa_df = pd.read_excel(arquivo_mapa, dtype={"CNPJ": str, "CPF": str, "Fatura": str})
        mapa_df["CNPJ"] = mapa_df["CNPJ"].replace([None, "nan", "0", "0.0", "", " "], pd.NA)
        mapa_df["Identificador"] = mapa_df["CNPJ"].fillna(mapa_df["CPF"])

        resultado = mapa_df.groupby(['Identificador', 'Plano Interno']).agg({
            'Nome': 'first',
            'Fatura': lambda x: ', '.join(
                map(lambda v: str(v).split(".")[0] if str(v).replace(".", "").isdigit() else str(v), x.unique())
            ),
            'Valor': 'sum'
        }).reset_index()

        resultado = resultado[['Nome', 'Identificador', 'Plano Interno', 'Fatura', 'Valor']]
        resultado.rename(columns={'Identificador': 'CNPJ/CPF'}, inplace=True)

        resultado['CNPJ/CPF'] = resultado['CNPJ/CPF'].apply(formatar_identificador)
        resultado['Valor'] = resultado['Valor'].apply(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        if arquivo_inex:
            inex_df = pd.read_excel(arquivo_inex, dtype={"CNPJ": str})
            inex_df["CNPJ"] = inex_df["CNPJ"].astype(str).str.zfill(14)
            resultado["CNPJ_Base"] = resultado["CNPJ/CPF"].str.replace(r'\D', '', regex=True).str.zfill(14)
            resultado = resultado.merge(inex_df[['CNPJ', 'INEX']], how='left', left_on="CNPJ_Base", right_on="CNPJ")
            resultado.drop(columns=["CNPJ_Base", "CNPJ"], inplace=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"relatorio_por_cnpj_{timestamp}.xlsx"
        caminho_completo = os.path.join(pasta_destino, nome_arquivo)

        resultado.to_excel(caminho_completo, index=False)
        status.configure(text=f"✅ Arquivo salvo: {nome_arquivo}", text_color="green", font=("Arial", 14, "bold"))
    except Exception as e:
        status.configure(text=f"❌ Erro: {str(e)}", text_color="red", font=("Arial", 14, "bold"))
    finally:
        progress.stop()            
        progress.pack_forget()    

# Frame principal
frame = ctk.CTkFrame(app, fg_color="black")
frame.pack(padx=20, pady=20, fill="both", expand=True)

# Barra de progresso (inicialmente oculta)
progress = ctk.CTkProgressBar(frame, mode="indeterminate")
progress.pack(fill="x", padx=30, pady=(10, 10))
progress.set(0)
progress.pack_forget()

# Título
ctk.CTkLabel(frame, text="Conversor de Execução Orçamentária", font=("Arial", 22, "bold"), text_color="white").pack(pady=(20, 10))
ctk.CTkFrame(frame, height=2, fg_color="gray").pack(fill="x", padx=30, pady=(10, 50))

# Botão principal
ctk.CTkButton(frame, text="📂 Anexar Mapa", command=selecionar_arquivo, height=45, width=200, font=("Arial", 14)).pack(pady=(0, 40))

buttons_frame = ctk.CTkFrame(frame, fg_color="transparent")
buttons_frame.pack()

ctk.CTkButton(buttons_frame, text="📁 Selecionar Pasta", command=selecionar_pasta, height=40, width=180, font=("Arial", 14)).pack(side="left", padx=20)
ctk.CTkButton(buttons_frame, text="📤 Converter", command=converter, height=40, width=180, font=("Arial", 14)).pack(side="right", padx=20)

# Status
status = ctk.CTkLabel(frame, text="📂 Selecione um mapa para gerar o relatório.", text_color="#5166A0", font=("Arial", 14, "bold"))
status.pack(pady=(30, 20))

ctk.CTkFrame(frame, height=2, fg_color="gray").pack(fill="x", padx=30, pady=(0, 20))

# Texto
texto_info = (
    "Este sistema importa mapas orçamentários fornecidos pelo SIPEO/DPGO, processando e organizando os dados por CNPJ/CPF e planos internos.\n"
    "Permite a inclusão opcional de dados complementares do arquivo INEX para enriquecer o relatório final.\n"
    "Realiza a formatação automática de identificadores fiscais (CPF e CNPJ), agrega informações e calcula valores totais.\n"
    "O resultado é um relatório detalhado em Excel, salvo na pasta escolhida pelo usuário.\n"
    "Destinado exclusivamente à Seção Administrativa da Base Administrativa do COPESP, facilita o controle e a solicitação de notas fiscais."
)

label_texto = ctk.CTkLabel(frame, text=texto_info, font=("Arial", 11), justify="left", text_color="white", anchor="w")
label_texto.pack(fill="both", expand=True, padx=30, pady=(0, 10))

ctk.CTkLabel(app, text="Desenvolvido por  Cb Pacífico", font=("Arial", 10, "italic"), text_color="gray").pack(anchor="w", padx=20, pady=(0, 10))


app.after(100, popup_incluir_inex)

app.mainloop()
