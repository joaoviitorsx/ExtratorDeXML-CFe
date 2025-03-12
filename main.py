import xmltodict
import os
import sys
import glob
import pandas as pd
from PySide6 import QtWidgets, QtGui, QtCore
from PySide6.QtWidgets import QFileDialog, QMessageBox, QLabel, QVBoxLayout, QPushButton, QProgressBar, QHBoxLayout
from utils.mensagem import mensagem_aviso, mensagem_sucesso, mensagem_error
from utils.icone import usar_icone

def extrair_dados_cfe(nome_arquivo, valores, valores_erros):
    try:
        with open(nome_arquivo, "rb") as arquivo_xml:
            dic_arquivo = xmltodict.parse(arquivo_xml)

        if 'CFe' not in dic_arquivo:
            valores_erros.append(nome_arquivo)
            return False

        info_cfe = dic_arquivo['CFe']['infCFe']
        chave_acesso = info_cfe.get('@Id', 'Não informado')
        cnpj_emitente = info_cfe['emit'].get('CNPJ', 'Não informado')
        nome_emitente = info_cfe['emit'].get('xNome', 'Não informado')
        numero_cupom = info_cfe['ide'].get('nCFe', 'Não informado')
        modelo = info_cfe['ide'].get('mod', 'Não informado')
        serie = info_cfe['ide'].get('serie', 'Não informado')
        data_emissao = info_cfe['ide'].get('dEmi', 'Não informado')
        hora_emissao = info_cfe['ide'].get('hEmi', 'Não informado')
        nome_destinatario = info_cfe.get('dest', {}).get('xNome', 'Não informado')
        cpf_cnpj_cliente = info_cfe.get('dest', {}).get('CNPJ', info_cfe.get('dest', {}).get('CPF', 'Não informado'))
        total_venda = info_cfe.get('total', {}).get('vCFe', 'Não informado')
        forma_pagamento = info_cfe.get('pgto', {}).get('MP', {}).get('cMP', 'Não informado')
        qr_code = info_cfe.get('infAdic', {}).get('qrcode', 'Não informado')
        nome_arquivo = os.path.basename(nome_arquivo)
        
        produtos = info_cfe['det']
        if not isinstance(produtos, list):
            produtos = [produtos]

        for produto in produtos:
            codigo_produto = produto['prod'].get('cProd', 'Não informado')
            descricao = produto['prod'].get('xProd', 'Não informado')
            ncm = produto['prod'].get('NCM', 'Não informado')
            cfop = produto['prod'].get('CFOP', 'Não informado')
            quantidade = produto['prod'].get('qCom', 'Não informado')
            unidade = produto['prod'].get('uCom', 'Não informado')
            valor_unitario = produto['prod'].get('vUnCom', 'Não informado')
            valor_total = produto['prod'].get('vItem', 'Não informado')
            
            cst_icms = produto.get('imposto', {}).get('ICMS', {}).get('@CST', 'Não informado')
            aliquota_icms = produto.get('imposto', {}).get('ICMS', {}).get('pICMS', 'Não informado')
            valor_icms = produto.get('imposto', {}).get('ICMS', {}).get('vICMS', 'Não informado')
            cst_pis = produto.get('imposto', {}).get('PIS', {}).get('@CST', 'Não informado')
            valor_pis = produto.get('imposto', {}).get('PIS', {}).get('vPIS', 'Não informado')
            cst_cofins = produto.get('imposto', {}).get('COFINS', {}).get('@CST', 'Não informado')
            valor_cofins = produto.get('imposto', {}).get('COFINS', {}).get('vCOFINS', 'Não informado')
            
            valores.append([
                chave_acesso, cnpj_emitente, nome_emitente, numero_cupom, modelo, serie, data_emissao, hora_emissao,
                nome_destinatario, cpf_cnpj_cliente, total_venda, forma_pagamento, qr_code, codigo_produto, descricao, ncm, cfop, quantidade, unidade,
                valor_unitario, valor_total, cst_icms, aliquota_icms, valor_icms, cst_pis, valor_pis, cst_cofins, valor_cofins, nome_arquivo
            ])
        return True
    except Exception as e:
        valores_erros.append([nome_arquivo, str(e)])
        return False

def selecionar_pasta(progresso):
    progresso.setValue(0)
    folder_path = QFileDialog.getExistingDirectory(None, "Selecione a pasta com os arquivos XML de CF-e")
    if not folder_path:
        mensagem_aviso("Nenhuma pasta foi selecionada.")
        return

    lista_arquivos = glob.glob(os.path.join(folder_path, "*.xml"))
    if not lista_arquivos:
        mensagem_aviso("A pasta selecionada não contém arquivos XML de CF-e.")
        return

    colunas = ["Chave de Acesso", "CNPJ Emitente", "Nome Emitente", "Número Cupom", "Modelo", "Série",
               "Data Emissão", "Hora Emissão", "Nome Cliente", "CPF/CNPJ Cliente", "Total Venda", "Forma de Pagamento", 
               "QR Code", "Código Produto", "Descrição","NCM", "CFOP", "Quantidade", "Unidade", "Valor Unitário", "Valor Total", 
               "CST ICMS", "Alíquota ICMS","Valor ICMS", "CST PIS", "Valor PIS", "CST COFINS", "Valor COFINS", "Arquivo"]
    valores = []
    valores_erros = []

    sucesso = sum([extrair_dados_cfe(arquivo, valores, valores_erros) for arquivo in lista_arquivos])
    falha = len(lista_arquivos) - sucesso
    mensagem_sucesso(f"Arquivos processados: {len(lista_arquivos)} | Sucesso: {sucesso} | Falha: {falha}")
    progresso.setValue(100)

    resposta = QMessageBox.question(None, "Salvar arquivo", "Deseja salvar os dados extraídos em um arquivo Excel?", QMessageBox.Yes | QMessageBox.No)
    if resposta == QMessageBox.No:
        return
    
    save_path, _ = QFileDialog.getSaveFileName(None, "Salvar arquivo Excel", "", "Arquivos Excel (*.xlsx)")
    if not save_path:
        mensagem_aviso("Nenhum caminho de destino foi selecionado.")
        return
    if not save_path.endswith('.xlsx'):
        save_path += '.xlsx'

    tabela = pd.DataFrame(columns=colunas, data=valores)
    tabela.to_excel(save_path, index=False, sheet_name="Dados CF-e")
    mensagem_sucesso(f"Arquivo Excel salvo com sucesso em:\n{save_path}")

    resposta = QMessageBox.question(None, "Abrir arquivo", "Deseja abrir o arquivo salvo?", QMessageBox.Yes | QMessageBox.No)
    if resposta == QMessageBox.Yes:
        os.startfile(save_path) if os.name == "nt" else os.system(f"open {save_path}")

def main():
    app = QtWidgets.QApplication(sys.argv)

    janela = QtWidgets.QMainWindow()
    usar_icone(janela)
    janela.setWindowTitle("Extrator de XMLs de CF-e")
    janela.setGeometry(100, 100, 650, 650)
    janela.setStyleSheet("background-color: #1B1B2F; color: white;")

    widget_central = QtWidgets.QWidget()
    janela.setCentralWidget(widget_central)
    layout = QVBoxLayout(widget_central)

    titulo = QLabel("Extrator de XMLs de CF-e")
    titulo.setAlignment(QtCore.Qt.AlignCenter)
    titulo.setFont(QtGui.QFont("Arial", 22, QtGui.QFont.Bold))
    titulo.setStyleSheet("color: #bf1d4b;")
    layout.addWidget(titulo)

    imagem_placeholder = QLabel()
    imagem_placeholder.setPixmap(QtGui.QPixmap("images/logo.png").scaled(400, 400, QtCore.Qt.KeepAspectRatio))
    imagem_placeholder.setAlignment(QtCore.Qt.AlignCenter)
    layout.addWidget(imagem_placeholder)

    botao_frame = QHBoxLayout()
    layout.addLayout(botao_frame)

    botoes = [
        ("Selecionar Pasta com XMLs dos CF-e", lambda: selecionar_pasta(progresso)),
    ]

    for texto, funcao in botoes:
        botao = QPushButton(texto)
        botao.clicked.connect(funcao)
        botao.setFont(QtGui.QFont("Arial", 14))
        botao.setStyleSheet("""
            QPushButton {
                background-color: #b6bab5;
                color: #1B1B2F;
                border: none;
                padding: 12px 18px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #bf1d4b;
            }
        """)
        botao.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        botao_frame.addWidget(botao)

    progresso = QProgressBar()
    progresso.setRange(0, 100)
    progresso.setValue(0)
    progresso.setStyleSheet("""
        QProgressBar {
            border: 2px solid #bf1d4b;
            border-radius: 5px;
            background-color: #333;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #40a12a;
            width: 10px;
        }
    """)
    layout.addWidget(progresso)

    janela.showMaximized()
    app.exec()

if __name__ == "__main__":
    main()