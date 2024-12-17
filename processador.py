import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from kivy.app import App
from kivy.uix.floatlayout import FloatLayout  # Importa FloatLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.popup import Popup
from kivy.uix.image import Image
from typing import Optional
import traceback


class VendasApp(App):
    def build(self) -> FloatLayout:  # Altera para FloatLayout
        # Inicializa o DataFrame global como None
        self.df_global: Optional[pd.DataFrame] = None
        # Cria um layout flutuante
        layout = FloatLayout()

        # Adiciona a imagem de fundo
        self.add_background_image(layout)

        # Adiciona um título ao layout
        layout.add_widget(self.criar_titulo("Processador de Vendas", size_hint=(
            None, None), size=(300, 50), pos_hint={'center_x': 0.5, 'top': 1}))

        # Cria um spinner para selecionar a categoria
        self.combo_categoria = Spinner(size_hint_y=None, height=40, background_color=(
            0, 0, 1, 1), pos_hint={'center_x': 0.5, 'top': 0.20})
        layout.add_widget(Label(text="Selecione a Categoria para Filtragem:",
                          size_hint_y=None, height=30, pos_hint={'center_x': 0.5, 'top': 0.20}))
        layout.add_widget(self.combo_categoria)

        # Adiciona botões de ação ao layout
        layout.add_widget(self.criar_botao("Selecionar Arquivo de Vendas", self.selecionar_arquivo, size_hint=(
            None, None), size=(10000, 40), pos_hint={'center_x': 0.5, 'top': 0.15}, background_color=(0, 0, 1, 1), color=(1, 1, 1, 1)))
        layout.add_widget(self.criar_botao("Filtrar", self.filtrar_categoria, size_hint=(
            None, None), size=(10000, 40), pos_hint={'center_x': 0.5, 'top': 0.10}, background_color=(0, 0, 1, 1), color=(1, 1, 1, 1)))
        return layout

    def add_background_image(self, layout: FloatLayout) -> None:
        """Adiciona uma imagem de fundo ao layout."""
        background = Image(source='c:/Users/yansa/Downloads/312349-10-erros-mais-comuns-em-gerenciamento-de-projetos-de-ti-1080x675-1.jpg',
                           allow_stretch=True, keep_ratio=False)
        layout.add_widget(background)  # Adiciona a imagem ao layout

    def criar_titulo(self, texto: str, **kwargs) -> Label:
        """Cria um rótulo de título."""
        return Label(text=texto, font_size='20sp', size_hint_y=None, height=40, **kwargs)

    def criar_botao(self, texto: str, funcao, **kwargs) -> Button:
        """Cria um botão com texto e função associada."""
        botao = Button(text=texto, size_hint_y=None, height=40, **kwargs)
        # Vincula a função ao evento de pressionar o botão
        botao.bind(on_press=funcao)
        return botao

    def processar_vendas(self, categoria: str) -> None:
        """Processa as vendas filtrando pela categoria selecionada."""
        if self.df_global is None:
            self.show_popup("Erro", "Nenhum arquivo de vendas carregado.")
            return

        try:
            df_filtrado = self.df_global[self.df_global['Categoria'] == categoria]
            self.exibir_dados_filtrados(df_filtrado)

            relatorio = self.gerar_relatorio(df_filtrado)
            self.gerar_pdf(df_filtrado)
            relatorio.to_csv('relatorio.csv', index=False)
            self.gerar_grafico(relatorio)

        except Exception as e:
            self.show_popup(
                "Erro", f"Ocorreu um erro ao processar as vendas: {str(e)}")

    def exibir_dados_filtrados(self, df: pd.DataFrame) -> None:
        """Exibe dados filtrados no console."""
        print("Dados Filtrados:")
        print(df[['ID do Produto', 'Nome do Produto', 'Categoria',
              'Preço (R$)', 'Quantidade Vendida', 'Data da Venda']])

    def gerar_relatorio(self, df: pd.DataFrame) -> pd.DataFrame:
        """Gera um relatório agregado por categoria."""
        return df.groupby('Categoria').agg({'Quantidade Vendida': 'sum', 'Preço (R$)': 'sum'})

    def gerar_pdf(self, df_filtrado: pd.DataFrame) -> None:
        """Gera um arquivo PDF com os dados filtrados."""
        c = canvas.Canvas("relatorio.pdf", pagesize=letter)
        c.drawString(100, 750, "Relatório de Dados Processados")
        c.drawString(100, 730, "----------------------------------------")
        self.adicionar_dados_pdf(c, df_filtrado)
        c.save()

    def adicionar_dados_pdf(self, c: canvas.Canvas, df_filtrado: pd.DataFrame) -> None:
        """Adiciona dados ao PDF de forma organizada."""
        y = 710
        # Cabeçalho do PDF
        c.drawString(
            100, y, "ID | Nome do Produto         | Categoria         | Preço (R$) | Qtd | Data")
        c.drawString(
            100, y - 15, "------------------------------------------------------------")
        y -= 30  # Espaçamento após o cabeçalho

        # Adiciona cada linha de dados ao PDF
        for index, row in df_filtrado.iterrows():
            c.drawString(100, y, f"{row['ID do Produto']: <3} | {row['Nome do Produto']: <25} | "
                         f"{row['Categoria']: <18} | {
                             row['Preço (R$)']: <10} | "
                         f"{row['Quantidade Vendida']: <3} | {row['Data da Venda']}")
            y -= 20  # Espaçamento entre as linhas

            if y < 50:  # Se o espaço estiver acabando, crie uma nova página
                c.showPage()
                y = 750  # Reinicia a posição y

        c.drawString(
            100, y, "------------------------------------------------------------")

    def gerar_grafico(self, relatorio: pd.DataFrame) -> None:
        """Gera um gráfico a partir do relatório agregado."""
        relatorio.plot(kind='bar', y=[
                       'Quantidade Vendida', 'Preço (R$)'], title='Vendas por Categoria')
        plt.ylabel('Total')
        plt.xlabel('Categoria')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('grafico_vendas.png')
        plt.show()

    def selecionar_arquivo(self, instance) -> None:
        """Abre um popup para selecionar o arquivo de vendas."""
        filechooser = FileChooserIconView(path="data/")
        popup = Popup(title=" Selecionar arquivo de vendas",
                      content=filechooser, size_hint=(0.9, 0.9))
        # Vincula a função ao evento de seleção do arquivo
        filechooser.bind(
            on_submit=lambda *args: self.carregar_arquivo(filechooser, popup))
        popup.open()

    def carregar_arquivo(self, filechooser: FileChooserIconView, popup: Popup) -> None:
        """Carrega o arquivo selecionado e atualiza o spinner de categorias."""
        selection = filechooser.selection
        if selection:
            try:
                # Carrega o arquivo de vendas em um DataFrame
                self.df_global = pd.read_excel(selection[0])
                self.atualizar_spinner_categorias()
                popup.dismiss()  # Fecha o popup
            except FileNotFoundError:
                # Exibe um popup de erro se o arquivo não for encontrado
                self.show_popup("Erro", "Arquivo não encontrado")
            except Exception as e:
                # Exibe um popup de erro se ocorrer uma exceção
                if "GetFileAttributesEx" in str(e):
                    print(f"Ignorando erro de acesso a arquivo: {str(e)}")
                else:
                    self.show_popup(
                        "Erro", f"Ocorreu um erro ao carregar o arquivo: {str(e)}")

    def atualizar_spinner_categorias(self) -> None:
        """Atualiza as opções do spinner de categorias."""
        categorias = self.df_global['Categoria'].unique()
        self.combo_categoria.values = ['Selecionar'] + categorias.tolist()
        self.combo_categoria.text = 'Selecionar'

    def filtrar_categoria(self, instance) -> None:
        """Filtra as vendas pela categoria selecionada."""
        categoria = self.combo_categoria.text
        if categoria != 'Selecionar':
            self.processar_vendas(categoria)
        else:
            # Exibe um popup de erro se nenhuma categoria for selecionada
            self.show_popup("Erro", "Selecione uma categoria para filtrar")

    def show_popup(self, title: str, message: str) -> None:
        """Exibe um popup com uma mensagem de erro."""
        popup = Popup(title=title, content=Label(
            text=message), size_hint=(0.9, 0.9))
        popup.open()


def main():
    try:
        VendasApp().run()
    except Exception as e:
        with open("error_log.txt", "w") as f:
            f.write(str(e) + "\n")
            f.write(traceback.format_exc())


if __name__ == "__main__":
    main()
