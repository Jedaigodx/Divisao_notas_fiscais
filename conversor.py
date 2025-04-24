import flet as ft
import pandas as pd
import os
from datetime import datetime

def main(page: ft.Page):
    page.title = "Conversor de Mapas - PMGU/GNA_COPESP"
    page.window_width = 600
    page.window_height = 540
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.vertical_alignment = ft.MainAxisAlignment.CENTER

    arquivo = ft.Text(value="", visible=False)
    pasta_destino = ft.Text(value=os.path.join(os.path.expanduser("~"), "Downloads"), visible=False)
    status_text = ft.Text("📂 Selecione um mapa para gerar o relatório.", size=14, color=ft.colors.BLUE)

    def selecionar_pasta(e: ft.FilePickerResultEvent):
        if e.path:
            pasta_destino.value = e.path
            status_text.value = f"📂 Pasta selecionada: {pasta_destino.value}"
            status_text.color = ft.colors.BLUE
            page.update()

    def pick_file_result(e: ft.FilePickerResultEvent):
        if e.files:
            arquivo.value = e.files[0].path
            status_text.value = f"📄 Arquivo selecionado: {os.path.basename(arquivo.value)}"
            status_text.color = ft.colors.BLUE
            page.update()

    def converter_click(e):
        try:
            mapa_df = pd.read_excel(arquivo.value, dtype={"CNPJ": str, "CPF": str, "Fatura": str})

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

            def format_identificador(val):
                val_str = str(val).zfill(14)
                val_str = ''.join(filter(str.isdigit, val_str))
                if len(val_str) == 11:
                    return f"{val_str[:3]}.{val_str[3:6]}.{val_str[6:9]}-{val_str[9:]}"
                elif len(val_str) == 14:
                    return f"{val_str[:2]}.{val_str[2:5]}.{val_str[5:8]}/{val_str[8:12]}-{val_str[12:]}"
                else:
                    return val_str

            resultado['CNPJ/CPF'] = resultado['CNPJ/CPF'].apply(format_identificador)

            resultado['Valor'] = resultado['Valor'].apply(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"relatorio_por_cnpj_{timestamp}.xlsx"
            caminho_completo = os.path.join(pasta_destino.value, nome_arquivo)

            resultado.to_excel(caminho_completo, index=False)

            status_text.value = f"✅ Arquivo salvo em: {nome_arquivo}"
            status_text.color = ft.colors.GREEN

        except Exception as ex:
            status_text.value = f"❌ Erro: {str(ex)}"
            status_text.color = ft.colors.RED
        page.update()

    file_picker = ft.FilePicker(on_result=pick_file_result)
    folder_picker = ft.FilePicker(on_result=selecionar_pasta)
    page.overlay.extend([file_picker, folder_picker])

    page.add(
        ft.Column([
            ft.Text("Conversor de Execução Orçamentária", size=22, weight=ft.FontWeight.BOLD),
            ft.Image(src="logo_pmgu.png", width=80, height=80, fit=ft.ImageFit.CONTAIN),
            ft.Divider(thickness=1),

            # Botão para anexar mapa
            ft.Row([
                ft.ElevatedButton(
                    "Anexar Mapa",
                    icon=ft.icons.FOLDER_OPEN,
                    on_click=lambda _: file_picker.pick_files(allow_multiple=False),
                )
            ], alignment=ft.MainAxisAlignment.CENTER),

            # Separação visual
            ft.Container(height=10),

            # Botões de selecionar pasta e converter
            ft.Row([
                ft.ElevatedButton(
                    "Selecionar Pasta",
                    icon=ft.icons.FOLDER,
                    on_click=lambda _: folder_picker.get_directory_path(),
                ),
                ft.ElevatedButton(
                    "Converter",
                    icon=ft.icons.UPLOAD_FILE,
                    on_click=converter_click,
                )
            ], alignment=ft.MainAxisAlignment.CENTER, spacing=20),

            ft.Container(height=10),
            status_text,
            ft.Divider(thickness=1),
            ft.Container(
                content=ft.Text(
                    "Este sistema transforma mapas disponibilizados pelo SIPEO/DPGO em relatórios organizados por CNPJ e plano Interno, "
                    "facilitando a solicitação de notas fiscais.\n"
                    "Uso exclusivo da Seção Administrativa da Base Administrativa do COPESP.",
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
        spacing=20)
    )

ft.app(target=main)

