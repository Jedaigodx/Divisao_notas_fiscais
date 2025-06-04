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
app.geometry("700x600")
app.resizable(False, False)

arquivo_mapa = ""
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

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"relatorio_por_cnpj_{timestamp}.xlsx"
        caminho_completo = os.path.join(pasta_destino, nome_arquivo)

        resultado.to_excel(caminho_completo, index=False)
        status.configure(text=f"✅ Arquivo salvo: {nome_arquivo}", text_color="green", font=("Arial", 14, "bold"))
    except Exception as e:
        status.configure(text=f"❌ Erro: {str(e)}", text_color="red", font=("Arial", 14, "bold"))

# Frame principal
frame = ctk.CTkFrame(app, fg_color="black")
frame.pack(padx=20, pady=20, fill="both", expand=True)

# Título
ctk.CTkLabel(frame, text="Conversor de Execução Orçamentária", font=("Arial", 25, "bold"), text_color="white").pack(pady=(20, 10))
ctk.CTkFrame(frame, height=2, fg_color="gray").pack(fill="x", padx=30, pady=(10, 50))


# Botão principal
ctk.CTkButton(frame, text="📂 Anexar Mapa", command=selecionar_arquivo, height=45, width=200, font=("Arial", 14)).pack(pady=(0, 40))

# Linha de botões
buttons_frame = ctk.CTkFrame(frame, fg_color="transparent")
buttons_frame.pack()

ctk.CTkButton(buttons_frame, text="📁 Selecionar Pasta", command=selecionar_pasta, height=40, width=180, font=("Arial", 14)).pack(side="left", padx=20)
ctk.CTkButton(buttons_frame, text="📤 Converter", command=converter, height=40, width=180, font=("Arial", 14)).pack(side="right", padx=20)

# Status
status = ctk.CTkLabel(frame, text="📂 Selecione um mapa para gerar o relatório.", text_color="#5166A0", font=("Arial", 14, "bold"))
status.pack(pady=(30, 20))

# Texto explicativo
ctk.CTkFrame(frame, height=2, fg_color="gray").pack(fill="x", padx=30, pady=(0, 50))
texto_info = (
    "Este sistema transforma mapas disponibilizados pelo SIPEO/DPGO em relatórios organizados por OCS/PSA e os seus plano Interno, facilitando a solicitação de notas fiscais.\n"
    "O seu uso é exclusivo da Seção Administrativa da Base Administrativa do COPESP."
)
ctk.CTkLabel(frame, text=texto_info, font=("Arial", 11), justify="center", text_color="white").pack(pady=(0, 10))

# desenvolvido por
ctk.CTkLabel(app, text="Desenvolvido por  Cb Pacífico", font=("Arial", 10, "italic"), text_color="gray").pack(anchor="w", padx=20, pady=(0, 10))

app.mainloop()
