import flet as ft
import pandas as pd
import os
from datetime import datetime

def main(page: ft.Page):
    page.title = "Conversor de Mapas - PMGU/GNA_COPESP"
    page.window_width = 600
    page.window_height = 520
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.vertical_alignment = ft.MainAxisAlignment.CENTER

    arquivo = ft.Text(value="", visible=False)
    status_text = ft.Text("📂 Selecione um mapa para gerar o relatório.", size=14, color=ft.colors.BLUE)

    def pick_file_result(e: ft.FilePickerResultEvent):
        if e.files:
            arquivo.value = e.files[0].path
            status_text.value = f"📄 Arquivo selecionado: {os.path.basename(arquivo.value)}"
            status_text.color = ft.colors.BLUE
            page.update()

    def converter_click(e):
     try:
        mapa_df = pd.read_excel(arquivo.value)

        # Tratar CNPJ inválido (0, "0", "", espaço, etc.)
        mapa_df["CNPJ"] = mapa_df["CNPJ"].replace([0, "0", "0.0", "", " "], pd.NA)
        mapa_df["Identificador"] = mapa_df["CNPJ"].fillna(mapa_df["CPF"])

        resultado = mapa_df.groupby(['Identificador', 'Plano Interno']).agg({
            'Nome': 'first',
            'Fatura': lambda x: ', '.join(
                map(lambda v: str(int(v)) if pd.notna(v) else '', x.unique())
            ),
            'Valor': 'sum'
        }).reset_index()

        resultado = resultado[['Nome', 'Identificador', 'Plano Interno', 'Fatura', 'Valor']]
        resultado.rename(columns={'Identificador': 'CNPJ/CPF'}, inplace=True)

        # Formatar CNPJ ou CPF com pontuação
        def format_identificador(val):
            val_str = str(int(val)).zfill(14)
            if len(val_str.strip("0")) <= 11:
                # CPF
                cpf = val_str[-11:]
                return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
            else:
                # CNPJ
                cnpj = val_str
                return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
        
        resultado['CNPJ/CPF'] = resultado['CNPJ/CPF'].apply(format_identificador)

        resultado['Valor'] = resultado['Valor'].apply(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"relatorio_por_cnpj_{timestamp}.xlsx"
        caminho_completo = os.path.join(pasta_downloads, nome_arquivo)

        resultado.to_excel(caminho_completo, index=False)

        status_text.value = f"✅ Arquivo salvo em: {nome_arquivo}"
        status_text.color = ft.colors.GREEN
     except Exception as ex:
            status_text.value = f"❌ Erro: {str(ex)}"
            status_text.color = ft.colors.RED
    page.update()



    file_picker = ft.FilePicker(on_result=pick_file_result)
    page.overlay.append(file_picker)

    page.add(
        ft.Column([
            ft.Text("Conversor de Execução Orçamentária", size=22, weight=ft.FontWeight.BOLD),
            ft.Image(
                src="logo_pmgu.png",
                width=80,
                height=80,
                fit=ft.ImageFit.CONTAIN
            ),
            ft.Divider(thickness=1),
            ft.Row([
                ft.ElevatedButton(
                    "Anexar Mapa",
                    icon=ft.icons.FOLDER_OPEN,
                    on_click=lambda _: file_picker.pick_files(allow_multiple=False),
                    style=ft.ButtonStyle(
                        padding=20,
                        shape=ft.RoundedRectangleBorder(radius=6),
                        text_style=ft.TextStyle(size=16)
                    )
                ),
                ft.ElevatedButton(
                    "Converter",
                    icon=ft.icons.UPLOAD_FILE,
                    on_click=converter_click,
                    style=ft.ButtonStyle(
                        padding=20,
                        shape=ft.RoundedRectangleBorder(radius=6),
                        text_style=ft.TextStyle(size=16)
                    )
                )
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=20),
            ft.Container(height=10),
            status_text,
            ft.Divider(thickness=1),
            ft.Container(
                content=ft.Text(
                    "Este sistema transforma mapas disponibilizados pelo SIPEO/DPGO em relatórios organizados por CNPJ e plano Interno, "
                    "facilitando a solicitação de notas fiscais.\n"
                    " Uso exclusivo da Seção Administrativa da Base Administrativa do COPESP.",
                    size=12,
                    text_align=ft.TextAlign.LEFT
                ),
                padding=10,
                alignment=ft.alignment.center_left
            ),
            ft.Container(height=10),
            ft.Row([
                ft.Text("Desenvolvido por ", size=12, italic=True, color=ft.colors.GREY),
                ft.TextButton(
                    text="Cb Pacífico",
                    url="https://github.com/Jedaigodx",  
                    style=ft.ButtonStyle(
                        padding=0,
                        overlay_color=ft.colors.with_opacity(0.1, ft.colors.BLUE_400),
                        text_style=ft.TextStyle(
                            size=12,
                            weight=ft.FontWeight.BOLD,
                            italic=True,
                            color=ft.colors.GREY,
                            decoration=ft.TextDecoration.UNDERLINE
                        )
                    )
                )
            ], alignment=ft.MainAxisAlignment.START)
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=20
        )
    )

ft.app(target=main)

