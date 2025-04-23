import flet as ft
import pandas as pd
import os

def main(page: ft.Page):
    page.title = "Conversor de Planilhas por CNPJ"
    page.window_width = 500
    page.window_height = 300
    page.vertical_alignment = ft.MainAxisAlignment.CENTER

    arquivo = ft.Text(value="", visible=False)
    status_text = ft.Text("Selecione uma planilha para começar.", color=ft.colors.BLUE)

    def pick_file_result(e: ft.FilePickerResultEvent):
        if e.files:
            arquivo.value = e.files[0].path
            status_text.value = f"Arquivo selecionado: {os.path.basename(arquivo.value)}"
            page.update()

    def converter_click(e):
        try:
            mapa_df = pd.read_excel(arquivo.value)

            resultado = mapa_df.groupby(['CNPJ', 'Plano Interno']).agg({
                'Nome': 'first',
                'Fatura': lambda x: ', '.join(map(str, x.unique())),
                'Valor': 'sum'
            }).reset_index()

            resultado = resultado[['Nome', 'CNPJ', 'Plano Interno', 'Fatura', 'Valor']]
            output_file = "relatorio_por_cnpj.xlsx"
            resultado.to_excel(output_file, index=False)

            status_text.value = f"✅ Conversão concluída! Arquivo salvo como: {output_file}"
            status_text.color = ft.colors.GREEN
        except Exception as ex:
            status_text.value = f"❌ Erro: {str(ex)}"
            status_text.color = ft.colors.RED
        page.update()

    file_picker = ft.FilePicker(on_result=pick_file_result)
    page.overlay.append(file_picker)

    page.add(
        ft.Column([
            ft.Text("Conversor de Planilhas", size=25, weight=ft.FontWeight.BOLD),
            ft.ElevatedButton("Selecionar Planilha", on_click=lambda _: file_picker.pick_files(allow_multiple=False)),
            ft.ElevatedButton("Converter", on_click=converter_click, icon=ft.icons.UPLOAD_FILE),
            status_text
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )
    )

ft.app(target=main)

