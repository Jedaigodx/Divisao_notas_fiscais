import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ConversorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Conversor de Mapas - PMGU/GNA_COPESP")
        self.geometry("600x500")
        self.resizable(False, False)

        self.arquivo = ""
        self.pasta_destino = os.path.join(os.path.expanduser("~"), "Downloads")

        self.create_widgets()

    def create_widgets(self):
        # Título
        ctk.CTkLabel(self, text="Conversor de Execução Orçamentária", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(10, 5))
        
        # Logo (substitua por um logo válido se quiser)
        try:
            from PIL import Image
            logo_image = ctk.CTkImage(Image.open("logo_pmgu.png"), size=(80, 80))
            ctk.CTkLabel(self, image=logo_image, text="").pack(pady=5)
        except:
            pass

        # Status
        self.status_label = ctk.CTkLabel(self, text="📂 Selecione um mapa para gerar o relatório.", text_color="blue")
        self.status_label.pack(pady=10)

        # Botões
        ctk.CTkButton(self, text="Anexar Mapa", command=self.selecionar_arquivo).pack(pady=5)
        ctk.CTkButton(self, text="Selecionar Pasta", command=self.selecionar_pasta).pack(pady=5)
        ctk.CTkButton(self, text="Converter", command=self.converter_mapa).pack(pady=10)

        # Rodapé
        rodape = (
            "Este sistema transforma mapas disponibilizados pelo SIPEO/DPGO em relatórios organizados por OCS/PSA e seus planos internos, "
            "facilitando a solicitação de notas fiscais.\nUso exclusivo da Seção Administrativa da Base Administrativa do COPESP."
        )
        ctk.CTkLabel(self, text=rodape, wraplength=550, font=ctk.CTkFont(size=12)).pack(pady=15)

        ctk.CTkLabel(self, text="Desenvolvido por Cb Pacífico", font=ctk.CTkFont(size=10, slant="italic")).pack()

    def selecionar_arquivo(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.arquivo = filepath
            self.status_label.configure(text=f"📄 Arquivo selecionado: {os.path.basename(filepath)}", text_color="blue")

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_destino = pasta
            self.status_label.configure(text=f"📂 Pasta selecionada: {pasta}", text_color="blue")

    def format_identificador(self, val):
        val_str = str(int(val)).zfill(14)
        stripped = str(int(val))
        if len(stripped) <= 11:
            cpf = val_str[-11:]
            return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
        else:
            cnpj = val_str
            return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

    def converter_mapa(self):
        try:
            if not self.arquivo:
                raise Exception("Nenhum arquivo selecionado.")

            mapa_df = pd.read_excel(self.arquivo, dtype={"CNPJ": str, "CPF": str, "Fatura": str})
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
            resultado['CNPJ/CPF'] = resultado['CNPJ/CPF'].apply(self.format_identificador)
            resultado['Valor'] = resultado['Valor'].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"relatorio_por_cnpj_{timestamp}.xlsx"
            caminho_completo = os.path.join(self.pasta_destino, nome_arquivo)

            resultado.to_excel(caminho_completo, index=False)
            self.status_label.configure(text=f"✅ Arquivo salvo em: {nome_arquivo}", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"❌ Erro: {str(e)}", text_color="red")

if __name__ == "__main__":
    app = ConversorApp()
    app.mainloop()
