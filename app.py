import sys
import os
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, 
                             QMessageBox, QComboBox, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QFrame, QStatusBar, QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QBrush

class ProcessThread(QThread):
    """Thread para processar os dados sem congelar a interface."""
    finished = pyqtSignal(pd.DataFrame, str)
    error = pyqtSignal(str)
    
    def __init__(self, app, output_path):
        super().__init__()
        self.app = app
        self.output_path = output_path
        
    def run(self):
        try:
            resultado_df = self.app.processar_classificacoes()
            
            # Tentar salvar com tratamento de erro específico para permissão
            try:
                resultado_df.to_excel(self.output_path, sheet_name='Classificação', index=False)
            except PermissionError:
                # Se falhar por causa de permissão, tente salvar em um local alternativo
                home_dir = os.path.expanduser("~")
                fallback_path = os.path.join(home_dir, "resultado_monitoria.xlsx")
                resultado_df.to_excel(fallback_path, sheet_name='Classificação', index=False)
                self.output_path = fallback_path  # Atualiza o caminho
                
            self.finished.emit(resultado_df, self.output_path)
        except Exception as e:
            self.error.emit(str(e))

class MonitoriaApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Variáveis para armazenar os dataframes
        self.notas_df = None
        self.inscricoes_df = None
        self.vagas_df = None
        self.resultado_df = None
        self.todas_candidaturas = None  # Nova variável para armazenar todas as candidaturas
        self.disciplinas = []  # Lista de disciplinas disponíveis
        
        # Variáveis para widgets críticos
        self.excel_path_entry = None
        self.output_path_entry = None
        self.process_btn = None
        self.process_info = None
        self.highlight_legend = None  # Inicializar explicitamente para evitar erros
        
        # Configuração da janela principal
        self.setWindowTitle("Podium - Sistema de Classificação para Monitoria")
        self.setGeometry(100, 100, 1000, 800)
        
        # Criar o widget de abas
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Criar as abas
        self.import_tab = QWidget()
        self.view_tab = QWidget()
        self.ranking_tab = QWidget()  # Nova aba para classificação por disciplina
        
        self.tabs.addTab(self.import_tab, "Importar Dados")
        self.tabs.addTab(self.view_tab, "Visualizar Dados")
        self.tabs.addTab(self.ranking_tab, "Classificação por Disciplina")
        
        # Configurar as abas
        self.setup_import_tab()
        self.setup_view_tab()
        self.setup_ranking_tab()
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Pronto")
        
        # Rastrear labels adicionais para gerenciá-los corretamente
        self.limit_label = None
        
        # Desabilitar o botão de processamento até que os dados sejam carregados
        if self.process_btn:
            self.process_btn.setEnabled(False)
            # Estilo para botão desabilitado (cor neutra)
            self.process_btn.setStyleSheet("background-color: #cccccc; color: #666666;")

    def setup_import_tab(self):
        # Criar layout principal
        main_layout = QVBoxLayout()
        
        # Título principal
        title_label = QLabel("Sistema de Classificação para Monitoria")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # ==== SEÇÃO DE IMPORTAÇÃO ====
        import_section = QFrame()
        import_layout = QVBoxLayout(import_section)
        
        # Título da seção
        import_title = QLabel("Importar Dados")
        import_title.setFont(QFont("Arial", 14, QFont.Bold))
        import_layout.addWidget(import_title)
        
        # Seleção de arquivo Excel
        excel_label = QLabel("Arquivo Excel com as planilhas (notas, inscricoes, vagas):")
        excel_label.setFont(QFont("Arial", 12))
        import_layout.addWidget(excel_label)
        
        # Layout para entrada de arquivo
        file_layout = QHBoxLayout()
        self.excel_path_entry = QLineEdit()
        excel_btn = QPushButton("Selecionar Arquivo")
        excel_btn.clicked.connect(self.load_excel_file)
        
        file_layout.addWidget(self.excel_path_entry)
        file_layout.addWidget(excel_btn)
        import_layout.addLayout(file_layout)
        
        # Botão de carregar dados
        load_btn = QPushButton("Carregar Dados")
        load_btn.setFont(QFont("Arial", 12, QFont.Bold))
        load_btn.setMinimumHeight(40)
        load_btn.clicked.connect(self.load_data)
        import_layout.addWidget(load_btn)
        
        main_layout.addWidget(import_section)
        
        # Separador
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(separator)
        
        # ==== SEÇÃO DE PROCESSAMENTO ====
        process_section = QFrame()
        process_layout = QVBoxLayout(process_section)
        
        # Título da seção
        process_title = QLabel("Processamento da Classificação")
        process_title.setFont(QFont("Arial", 14, QFont.Bold))
        process_layout.addWidget(process_title)
        
        # Descrição
        process_description = QLabel("Configure o arquivo de saída e clique em Processar para gerar a classificação.")
        process_description.setWordWrap(True)
        process_layout.addWidget(process_description)
        
        # Configuração do arquivo de saída
        output_label = QLabel("Arquivo de saída:")
        self.output_path_entry = QLineEdit("resultado_monitoria.xlsx")
        output_btn = QPushButton("Selecionar")
        output_btn.clicked.connect(self.select_output_file)
        
        output_layout = QHBoxLayout()
        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_path_entry)
        output_layout.addWidget(output_btn)
        process_layout.addLayout(output_layout)
        
        # Botão de processamento - inicialmente em cor neutra quando desabilitado
        self.process_btn = QPushButton("Processar Classificação")
        self.process_btn.setFont(QFont("Arial", 12, QFont.Bold))
        self.process_btn.setMinimumHeight(50)
        self.process_btn.setStyleSheet("background-color: #cccccc; color: #666666;")
        self.process_btn.clicked.connect(self.process_data)
        self.process_btn.setEnabled(False)  # Desabilitado inicialmente
        process_layout.addWidget(self.process_btn)
        
        # Informações de processamento
        self.process_info = QLabel("")
        self.process_info.setAlignment(Qt.AlignCenter)
        process_layout.addWidget(self.process_info)
        
        main_layout.addWidget(process_section)
        
        # Adicionar espaçamento
        main_layout.addStretch()
        
        self.import_tab.setLayout(main_layout)

    def setup_view_tab(self):
        layout = QVBoxLayout()
        
        # ComboBox para selecionar o dataset a visualizar
        combo_layout = QHBoxLayout()
        combo_label = QLabel("Selecione os dados para visualizar:")
        self.data_selector = QComboBox()
        self.data_selector.addItems(["Notas", "Inscrições", "Vagas", "Resultado"])
        self.data_selector.currentTextChanged.connect(self.change_dataset_view)
        
        combo_layout.addWidget(combo_label)
        combo_layout.addWidget(self.data_selector)
        combo_layout.addStretch()
        
        layout.addLayout(combo_layout)
        
        # Área de rolagem para a tabela
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        # Widget para conter a tabela
        self.table_container = QWidget()
        self.table_layout = QVBoxLayout(self.table_container)
        
        # Criar a tabela inicial (vazia)
        self.table = QTableWidget(0, 0)
        self.table_layout.addWidget(self.table)
        
        # Legenda para as cores (adicionada uma vez e mantida oculta até ser necessária)
        self.legend_label = QLabel("As cores de fundo indicam diferentes disciplinas para facilitar a visualização.")
        self.table_layout.addWidget(self.legend_label)
        self.legend_label.setVisible(False)
        
        scroll_area.setWidget(self.table_container)
        layout.addWidget(scroll_area)
        
        self.view_tab.setLayout(layout)

    def setup_ranking_tab(self):
        layout = QVBoxLayout()
        
        # Título
        title_label = QLabel("Classificação por Disciplina")
        title_label.setFont(QFont("Arial", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # Seletor de disciplina
        disc_layout = QHBoxLayout()
        disc_label = QLabel("Selecione a disciplina:")
        self.disc_selector = QComboBox()
        self.disc_selector.currentTextChanged.connect(self.show_discipline_ranking)
        
        disc_layout.addWidget(disc_label)
        disc_layout.addWidget(self.disc_selector)
        disc_layout.addStretch()
        
        layout.addLayout(disc_layout)
        
        # Informação sobre vagas
        self.vagas_info = QLabel("")
        self.vagas_info.setAlignment(Qt.AlignCenter)
        self.vagas_info.setStyleSheet("font-weight: bold; color: #1976D2;")
        layout.addWidget(self.vagas_info)
        
        # Container para a tabela de classificação
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        self.ranking_container = QWidget()
        self.ranking_layout = QVBoxLayout(self.ranking_container)
        
        # Criar a tabela inicial (vazia)
        self.ranking_table = QTableWidget(0, 0)
        self.ranking_layout.addWidget(self.ranking_table)
        
        # Legenda para estudantes já classificados - IMPORTANTE: garantir que não é None
        self.highlight_legend = QLabel("Destacado em amarelo: Estudantes já classificados em outras disciplinas")
        self.highlight_legend.setStyleSheet("color: #FF8C00; font-style: italic;")
        self.highlight_legend.setVisible(False)
        self.ranking_layout.addWidget(self.highlight_legend)
        
        scroll_area.setWidget(self.ranking_container)
        layout.addWidget(scroll_area)
        
        # Informação sobre a classificação
        self.ranking_info = QLabel("Para visualizar a classificação, carregue os dados e processe a classificação primeiro.")
        self.ranking_info.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.ranking_info)
        
        self.ranking_tab.setLayout(layout)

    def load_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione o arquivo Excel",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        if file_path:
            self.excel_path_entry.setText(file_path)

    def select_output_file(self):
        home_dir = os.path.expanduser("~")
        documents_dir = os.path.join(home_dir, "Documentos" if os.name == "nt" else "Documents")
        
        # Garantir que temos um caminho inicial válido
        initial_dir = documents_dir if os.path.exists(documents_dir) else home_dir
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Selecione onde salvar o arquivo de resultado",
            os.path.join(initial_dir, "resultado_monitoria.xlsx"),
            "Excel files (*.xlsx);;All files (*.*)"
        )
        
        if file_path:
            self.output_path_entry.setText(file_path)

    def load_data(self):
        try:
            self.status_bar.showMessage("Carregando dados...")
            
            # Verificar se o arquivo Excel foi selecionado
            if not self.excel_path_entry.text():
                QMessageBox.critical(self, "Erro", "Por favor, selecione o arquivo Excel com os dados.")
                self.status_bar.showMessage("Erro ao carregar dados.")
                return
                
            # Carregamento de arquivo único
            excel_path = self.excel_path_entry.text()
            self.notas_df = pd.read_excel(excel_path, sheet_name='notas')
            self.inscricoes_df = pd.read_excel(excel_path, sheet_name='inscricoes')
            self.vagas_df = pd.read_excel(excel_path, sheet_name='vagas')
            
            # Verificar se os dados foram carregados corretamente
            if self.notas_df is None or self.inscricoes_df is None or self.vagas_df is None:
                QMessageBox.critical(self, "Erro", "Falha ao carregar um ou mais conjuntos de dados.")
                self.status_bar.showMessage("Erro ao carregar dados.")
                return
            
            # Atualizar a lista de disciplinas para o combobox
            self.disciplinas = sorted(self.vagas_df['DISCIPLINA'].unique())
            self.disc_selector.clear()
            self.disc_selector.addItems(self.disciplinas)
            
            # Pré-calcular todas as candidaturas para a aba de classificação
            self.todas_candidaturas = self.criar_candidaturas()
            
            # Exibir os dados iniciais (notas)
            self.data_selector.setCurrentText("Notas")
            self.change_dataset_view("Notas")
            
            # ALTERAÇÃO: Não mudar para a aba de visualização, apenas habilitar o botão de processamento
            # e mudar sua cor para verde
            self.process_btn.setEnabled(True)
            self.process_btn.setStyleSheet("background-color: #4CAF50; color: white;")
            
            QMessageBox.information(self, "Sucesso", "Dados carregados com sucesso!")
            self.status_bar.showMessage("Dados carregados. Pronto para processar.")
            
            # Atualizar a mensagem na aba de classificação
            self.ranking_info.setText("Dados carregados. Selecione uma disciplina para ver a classificação.")
            
            # Mostrar a classificação da primeira disciplina na lista
            if self.disciplinas:
                self.disc_selector.setCurrentText(self.disciplinas[0])
                self.show_discipline_ranking(self.disciplinas[0])
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar os dados: {str(e)}")
            self.status_bar.showMessage(f"Erro: {str(e)}")

    def create_table(self, df):
        # Limpar a tabela anterior
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        
        # Esconder a legenda de cores
        self.legend_label.setVisible(False)
        
        # Remover a label de limitação de linhas se existir
        if self.limit_label is not None:
            self.limit_label.setVisible(False)
        
        if df is None:
            self.table.setRowCount(1)
            self.table.setColumnCount(1)
            item = QTableWidgetItem("Nenhum dado disponível para visualização.")
            self.table.setItem(0, 0, item)
            return
        
        # Configurar a tabela com base no DataFrame
        rows, cols = df.shape
        self.table.setRowCount(min(rows, 100))  # Limitar a 100 linhas
        self.table.setColumnCount(cols)
        
        # Definir cabeçalhos das colunas
        self.table.setHorizontalHeaderLabels(df.columns)
        
        # Preencher a tabela com os dados
        for i in range(min(rows, 100)):
            for j in range(cols):
                value = df.iloc[i, j]
                if pd.isna(value):
                    item = QTableWidgetItem("")
                else:
                    item = QTableWidgetItem(str(value))
                
                self.table.setItem(i, j, item)
        
        # Ajustar o tamanho das colunas ao conteúdo
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        # Adicionar nota sobre a limitação
        if rows > 100:
            if self.limit_label is None:
                self.limit_label = QLabel(f"Nota: Exibindo apenas 100 de {rows} linhas para melhor desempenho.")
                self.table_layout.addWidget(self.limit_label)
            else:
                self.limit_label.setText(f"Nota: Exibindo apenas 100 de {rows} linhas para melhor desempenho.")
                self.limit_label.setVisible(True)

    def create_result_table_with_colors(self, df):
        # Limpar a tabela anterior
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        
        # Esconder a legenda por padrão
        self.legend_label.setVisible(False)
        
        if df is None or df.empty:
            self.table.setRowCount(1)
            self.table.setColumnCount(1)
            item = QTableWidgetItem("Nenhum resultado disponível para visualização.")
            self.table.setItem(0, 0, item)
            return
        
        # Configurar a tabela com base no DataFrame
        rows, cols = df.shape
        self.table.setRowCount(min(rows, 100))  # Limitar a 100 linhas
        self.table.setColumnCount(cols)
        
        # Definir cabeçalhos das colunas
        self.table.setHorizontalHeaderLabels(df.columns)
        
        # Gerar cores para cada disciplina
        disciplinas_unicas = df['Disciplina'].unique()
        num_disciplinas = len(disciplinas_unicas)
        
        # Lista de cores para disciplinas (cores suaves)
        cores = [
            QColor(230, 230, 250),  # Lavender
            QColor(255, 245, 238),  # Seashell
            QColor(240, 248, 255),  # Alice Blue
            QColor(245, 245, 220),  # Beige
            QColor(255, 240, 245),  # Lavender Blush
            QColor(240, 255, 240),  # Honeydew
            QColor(255, 250, 240),  # Floral White
            QColor(245, 255, 250),  # Mint Cream
            QColor(248, 248, 255),  # Ghost White
            QColor(255, 228, 225),  # Misty Rose
            QColor(245, 222, 179),  # Wheat
            QColor(255, 255, 224),  # Light Yellow
            QColor(220, 220, 220),  # Gainsboro
            QColor(240, 255, 255),  # Azure
            QColor(211, 211, 211)   # Light Gray
        ]
        
        # Se houver mais disciplinas que cores, repetimos as cores
        if num_disciplinas > len(cores):
            cores = cores * (num_disciplinas // len(cores) + 1)
        
        # Mapear disciplinas para cores
        disciplina_para_cor = {disciplina: cores[i % len(cores)] for i, disciplina in enumerate(disciplinas_unicas)}
        
        # Preencher a tabela com os dados
        for i in range(min(rows, 100)):
            disciplina_atual = df.iloc[i]['Disciplina']
            cor_fundo = disciplina_para_cor[disciplina_atual]
            
            for j in range(cols):
                value = df.iloc[i, j]
                if pd.isna(value):
                    item = QTableWidgetItem("")
                else:
                    item = QTableWidgetItem(str(value))
                
                # Aplicar cor de fundo baseada na disciplina
                item.setBackground(cor_fundo)
                # Forçar texto preto para todas as células, independente do tema
                item.setForeground(QColor(0, 0, 0))
                self.table.setItem(i, j, item)
        
        # Ajustar o tamanho das colunas ao conteúdo
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        # Mostrar legenda de cores
        self.legend_label.setVisible(True)
        
        # Adicionar nota sobre a limitação
        if rows > 100:
            if self.limit_label is None:
                self.limit_label = QLabel(f"Nota: Exibindo apenas 100 de {rows} linhas para melhor desempenho.")
                self.table_layout.addWidget(self.limit_label)
            else:
                self.limit_label.setText(f"Nota: Exibindo apenas 100 de {rows} linhas para melhor desempenho.")
                self.limit_label.setVisible(True)
        elif self.limit_label is not None:
            self.limit_label.setVisible(False)

    def change_dataset_view(self, selection):
        if selection == "Notas":
            self.create_table(self.notas_df)
        elif selection == "Inscrições":
            self.create_table(self.inscricoes_df)
        elif selection == "Vagas":
            self.create_table(self.vagas_df)
        elif selection == "Resultado":
            # Aqui aplicamos cores de fundo diferentes por disciplina
            if self.resultado_df is not None:
                self.create_result_table_with_colors(self.resultado_df)
            else:
                self.create_table(None)

    def show_discipline_ranking(self, disciplina):
        if not disciplina or self.todas_candidaturas is None:
            return
        
        # Verificar se o widget highlight_legend existe
        if self.highlight_legend is None:
            # Se por algum motivo não existe, crie-o novamente
            self.highlight_legend = QLabel("Destacado em amarelo: Estudantes já classificados em outras disciplinas")
            self.highlight_legend.setStyleSheet("color: #FF8C00; font-style: italic;")
            self.ranking_layout.addWidget(self.highlight_legend)
        
        # Filtrar candidatos para a disciplina selecionada
        candidatos_disciplina = [c for c in self.todas_candidaturas if c['DISCIPLINA'] == disciplina]
        
        # Ordenar por média classificatória (do maior para o menor)
        candidatos_ordenados = sorted(candidatos_disciplina, 
                                    key=lambda x: x['MEDIA_CLASSIFICATORIA'], 
                                    reverse=True)
        
        # Criar dataframe para exibição
        df = pd.DataFrame(candidatos_ordenados)
        
        if df.empty:
            self.ranking_table.setRowCount(1)
            self.ranking_table.setColumnCount(1)
            item = QTableWidgetItem("Nenhum candidato para esta disciplina.")
            self.ranking_table.setItem(0, 0, item)
            self.vagas_info.setText(f"Disciplina: {disciplina} - Sem candidatos")
            self.highlight_legend.setVisible(False)
            return
        
        # Renomear colunas para melhor visualização
        df = df.rename(columns={
            'NOME': 'Nome',
            'MATRICULA': 'Matrícula',
            'MEDIA_CLASSIFICATORIA': 'Média Classificatória',
            'OPCAO': 'Opção',
            'NOTA_DISCIPLINA': 'Nota na Disciplina',
            'MEDIA_GLOBAL': 'Média Global'
        })
        
        # Formatar média classificatória para 4 casas decimais
        df['Média Classificatória'] = df['Média Classificatória'].apply(lambda x: round(x, 4))
        
        # Substituir OPCAO por um texto mais amigável
        df['Opção'] = df['Opção'].replace({
            'PRIMEIRA OPCAO': '1ª OPÇÃO',
            'SEGUNDA OPCAO': '2ª OPÇÃO',
            'TERCEIRA OPCAO': '3ª OPÇÃO'
        })
        
        # Obter número de vagas para a disciplina
        try:
            num_vagas = self.vagas_df[self.vagas_df['DISCIPLINA'] == disciplina]['VAGAS'].iloc[0]
            self.vagas_info.setText(f"Disciplina: {disciplina} - Vagas: {num_vagas} - Candidatos: {len(df)}")
        except:
            num_vagas = "Desconhecido"
            self.vagas_info.setText(f"Disciplina: {disciplina} - Vagas: {num_vagas} - Candidatos: {len(df)}")
        
        cols = list(df.columns)
        
        # Configurar a tabela
        self.ranking_table.setRowCount(len(df))
        self.ranking_table.setColumnCount(len(cols))
        self.ranking_table.setHorizontalHeaderLabels(cols)
        
        # Verificar se já temos o resultado processado para destacar estudantes classificados em outras disciplinas
        estudantes_classificados_por_disciplina = {}
        tem_destaques = False
        
        if self.resultado_df is not None:
            for _, row in self.resultado_df.iterrows():
                disc = row['Disciplina']
                nome = row['Nome']
                if disc not in estudantes_classificados_por_disciplina:
                    estudantes_classificados_por_disciplina[disc] = []
                estudantes_classificados_por_disciplina[disc].append(nome)
        
        # Preencher a tabela
        for i in range(len(df)):
            nome_estudante = df.iloc[i]['Nome']
            
            # Verificar se o estudante está classificado em outra disciplina que não a atual
            classificado_em_outra_disciplina = False
            for disc, estudantes in estudantes_classificados_por_disciplina.items():
                if disc != disciplina and nome_estudante in estudantes:
                    classificado_em_outra_disciplina = True
                    tem_destaques = True
                    break
            
            # Coloca os dados do DataFrame
            for j, col in enumerate(df.columns):
                value = df.iloc[i, j]
                if pd.isna(value):
                    item = QTableWidgetItem("")
                else:
                    item = QTableWidgetItem(str(value))
                
                # Destacar estudantes já classificados em outras disciplinas com fundo amarelo
                if classificado_em_outra_disciplina:
                    item.setBackground(QBrush(QColor(255, 255, 153)))  # Amarelo claro
                    item.setForeground(QBrush(QColor(0, 0, 0)))  # Texto na cor preta
                
                self.ranking_table.setItem(i, j, item)
        
        # Mostrar ou esconder a legenda de destaque
        self.highlight_legend.setVisible(tem_destaques)
        
        # Ajustar o tamanho das colunas ao conteúdo
        self.ranking_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        # Atualizar a informação de classificação
        self.ranking_info.setText("A tabela mostra todos os candidatos inscritos para esta disciplina, " +
                                "ordenados por média classificatória (independentemente da prioridade de opção).")

    def process_data(self):
        # Verificar se os dados foram carregados
        if self.notas_df is None or self.inscricoes_df is None or self.vagas_df is None:
            QMessageBox.critical(self, "Erro", "Por favor, carregue todos os dados primeiro.")
            return
        
        # Obter o caminho de saída
        output_path = self.output_path_entry.text()
        
        # Verificar se o diretório de saída tem permissão de escrita
        output_dir = os.path.dirname(output_path)
        if output_dir == "":  # Se não há diretório especificado (só nome de arquivo)
            # Usar a pasta de documentos do usuário
            home_dir = os.path.expanduser("~")
            documents_dir = os.path.join(home_dir, "Documentos" if os.name == "nt" else "Documents")
            output_dir = documents_dir if os.path.exists(documents_dir) else home_dir
            output_path = os.path.join(output_dir, output_path)
            self.output_path_entry.setText(output_path)
        
        # Verificar permissão de escrita
        try:
            if not os.access(output_dir, os.W_OK):
                QMessageBox.warning(
                    self, 
                    "Atenção", 
                    f"Sem permissão de escrita no diretório {output_dir}.\n"
                    "Por favor, escolha outro local para salvar o arquivo."
                )
                self.select_output_file()
                return
        except Exception as e:
            QMessageBox.warning(
                self,
                "Atenção",
                f"Não foi possível verificar permissões no diretório.\n"
                "Por favor, escolha explicitamente um local para salvar o arquivo.\n"
                f"Erro: {str(e)}"
            )
            self.select_output_file()
            return
            
        # Se chegou até aqui, temos permissão para escrever
        # Iniciar processamento em uma thread separada
        self.status_bar.showMessage("Processando dados...")
        self.process_info.setText("Processamento iniciado... Aguarde...")
        
        # Desabilitar o botão durante o processamento e mudar para cor cinza
        self.process_btn.setEnabled(False)
        self.process_btn.setStyleSheet("background-color: #cccccc; color: #666666;")
        
        self.process_thread = ProcessThread(self, output_path)
        self.process_thread.finished.connect(self.on_process_finished)
        self.process_thread.error.connect(self.on_process_error)
        self.process_thread.start()

    def on_process_finished(self, resultado_df, output_path):
        self.resultado_df = resultado_df
        self.data_selector.setCurrentText("Resultado")
        self.change_dataset_view("Resultado")
        
        self.status_bar.showMessage("Processamento concluído com sucesso.")
        self.process_info.setText(f"Processamento concluído!\nArquivo salvo em: {output_path}")
        
        # Atualizar a classificação por disciplina também
        current_disc = self.disc_selector.currentText()
        if current_disc:
            self.show_discipline_ranking(current_disc)
            
        # Mudar para a aba de visualização para mostrar o resultado
        self.tabs.setCurrentIndex(1)  # Índice da aba de visualização
        
        # Reativar o botão de processamento e voltar à cor verde
        self.process_btn.setEnabled(True)
        self.process_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        
        QMessageBox.information(self, "Sucesso", f"Processamento concluído com sucesso!\nArquivo salvo em: {output_path}")

    def on_process_error(self, error_msg):
        self.status_bar.showMessage(f"Erro no processamento: {error_msg}")
        self.process_info.setText(f"Erro durante o processamento: {error_msg}")
        
        # Reativar o botão de processamento e voltar à cor verde
        self.process_btn.setEnabled(True)
        self.process_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        
        QMessageBox.critical(self, "Erro", f"Erro durante o processamento: {error_msg}")

    def calcular_media_classificatoria(self, nota_disciplina, media_global):
        return (2 * nota_disciplina + media_global) / 3

    def criar_candidaturas(self):
        todas_candidaturas = []
        for _, inscricao in self.inscricoes_df.iterrows():
            aluno_nome = inscricao['ESTUDANTE']
            notas_aluno = self.notas_df[self.notas_df['ESTUDANTE'] == aluno_nome].iloc[0]

            for opcao in ['PRIMEIRA OPCAO', 'SEGUNDA OPCAO', 'TERCEIRA OPCAO']:
                if pd.isna(inscricao[opcao]):
                    continue

                disciplina = inscricao[opcao]
                nota_disciplina = notas_aluno[disciplina]
                media_global = notas_aluno['Média Global']
                media_class = self.calcular_media_classificatoria(nota_disciplina, media_global)

                todas_candidaturas.append({
                    'NOME': aluno_nome,
                    'MATRICULA': inscricao['MATRICULA'],
                    'DISCIPLINA': disciplina,
                    'MEDIA_CLASSIFICATORIA': media_class,
                    'OPCAO': opcao,
                    'NOTA_DISCIPLINA': nota_disciplina,
                    'MEDIA_GLOBAL': media_global
                })
        return todas_candidaturas

    def get_ranking_disciplina(self, candidaturas, disciplina, classificados):
        # Pega todos os candidatos não classificados para a disciplina
        candidatos_disciplina = [
            c for c in candidaturas
            if c['DISCIPLINA'] == disciplina and c['NOME'] not in classificados
        ]

        # Ordena por média classificatória
        return sorted(candidatos_disciplina,
                     key=lambda x: x['MEDIA_CLASSIFICATORIA'],
                     reverse=True)

    def processar_classificacoes(self):
        todas_candidaturas = self.criar_candidaturas()
        classificados = set()  # conjunto de alunos já classificados
        resultado_final = {disc['DISCIPLINA']: [] for _, disc in self.vagas_df.iterrows()}
        vagas_restantes = {row['DISCIPLINA']: row['VAGAS'] for _, row in self.vagas_df.iterrows()}

        # FASE 1: Classificar apenas candidatos de 1ª opção que estejam bem ranqueados
        for disciplina in vagas_restantes.keys():
            if vagas_restantes[disciplina] > 0:
                # Obtém ranking completo da disciplina
                ranking = self.get_ranking_disciplina(todas_candidaturas, disciplina, classificados)

                # Classifica apenas candidatos de 1ª opção dentro do número de vagas
                for candidato in ranking[:vagas_restantes[disciplina]]:
                    if candidato['OPCAO'] == 'PRIMEIRA OPCAO':
                        resultado_final[disciplina].append(candidato)
                        classificados.add(candidato['NOME'])
                        vagas_restantes[disciplina] -= 1

        # FASE 2: Classificar candidatos de 1ª e 2ª opção nas vagas remanescentes
        for disciplina in vagas_restantes.keys():
            if vagas_restantes[disciplina] > 0:
                # Obtém novo ranking (sem os já classificados)
                ranking = self.get_ranking_disciplina(todas_candidaturas, disciplina, classificados)

                # Classifica candidatos de 1ª e 2ª opção
                for candidato in ranking[:vagas_restantes[disciplina]]:
                    if candidato['OPCAO'] in ['PRIMEIRA OPCAO', 'SEGUNDA OPCAO']:
                        resultado_final[disciplina].append(candidato)
                        classificados.add(candidato['NOME'])
                        vagas_restantes[disciplina] -= 1

        # FASE 3: Preencher vagas restantes com qualquer opção
        for disciplina in vagas_restantes.keys():
            if vagas_restantes[disciplina] > 0:
                # Obtém ranking final (sem os já classificados)
                ranking = self.get_ranking_disciplina(todas_candidaturas, disciplina, classificados)

                # Classifica candidatos restantes
                for candidato in ranking[:vagas_restantes[disciplina]]:
                    resultado_final[disciplina].append(candidato)
                    classificados.add(candidato['NOME'])
                    vagas_restantes[disciplina] -= 1

        # Criar DataFrame com resultados
        linhas = []
        for disciplina, classificados in resultado_final.items():
            # Ordenar os classificados por média classificatória antes de adicionar ao DataFrame
            classificados_ordenados = sorted(classificados,
                                           key=lambda x: x['MEDIA_CLASSIFICATORIA'],
                                           reverse=True)

            for pos, aluno in enumerate(classificados_ordenados, 1):
                linhas.append({
                    'Disciplina': disciplina,
                    'Posição': pos,
                    'Nome': aluno['NOME'],
                    'Matrícula': aluno['MATRICULA'],
                    'Média Classificatória': round(aluno['MEDIA_CLASSIFICATORIA'], 4),
                    'Opção': aluno['OPCAO'].replace(' OPCAO', ' OPÇÃO'),
                    'Nota na Disciplina': aluno['NOTA_DISCIPLINA'],
                    'Média Global': aluno['MEDIA_GLOBAL']
                })

        resultado_final_df = pd.DataFrame(linhas)

        # Ordenar o DataFrame por Disciplina e Posição
        resultado_final_df = resultado_final_df.sort_values(['Disciplina', 'Posição'])
        
        return resultado_final_df

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MonitoriaApp()
    window.show()
    sys.exit(app.exec_())