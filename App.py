#! /usr/bin/env python
# -*- coding: utf-8 -*-

"""Leitor-e-Digitador-Etiquetas.py: 
        Desenvolvido para realizar a leitura e digitação de etiquetas."""

__author__ = "Diogo Dias"
__version__ = "1.0.0"
__email__ = "diogo@frangoamericano.com"
__status__ = "Em desenvolvimento"


from Relatorio import *
from PyQt5 import uic,QtWidgets
from PyQt5.QtWidgets import QApplication, QMessageBox, QShortcut, QFileDialog
from PyQt5.QtGui import QKeySequence
#from PyQt5.QtWidgets import QApplication,  QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QMenuBar, QMenu, QAction
#from PyKDE4.kdeui import KDateComboBox
#import sqlite3
#import datetime
import os
import sys

class Main_Window(QtWidgets.QMainWindow):
    def __init__(self):
        super(Main_Window, self).__init__()
        uic.loadUi("GUI\Leitor Etiquetas Python.ui", self)
        self.btnProcurar.setFocus()

        #TECLAS DE ATALHOS
        self.shortcut_sair = QShortcut(QKeySequence("Esc"), self)
        self.shortcut_sair.activated.connect(exit)
        
        #BOTÕES
        self.btnSair.clicked.connect(exit)
        self.btnProcurar.clicked.connect(self.Procurar)
        self.btnProcessar.clicked.connect(self.Processar)
        
        #LISTAS
        self._listaEtiquetasGeral = list()
        self._listaEtiquetasFiltro = list()
        
        #ENTER
        #self.btnSair.released.connect(exit)
    
    def ativar_botoes(self):
        """Verificar se foi digitado o espelho
        se SIM, ativar ativa os botões
        """
        if self.txt_espelho.text() == "":
            self.btn_cadastrar.setEnabled(False)
            self.btn_abrir.setEnabled(False)
            self.btn_exportar.setEnabled(False)
        else:
            self.btn_cadastrar.setEnabled(True)
            self.btn_abrir.setEnabled(True)
            self.btn_exportar.setEnabled(True)
            
    
    def verificarArquivo(self, arquivo):
        """Verifica se o arquivo existe

        Args:
            arquivo (str): path e nome do Arquivo

        Returns:
            bolean: Se o arquivo existe, retorna True
        """
        if os.path.isfile(arquivo):
            return True
        else:
            return False
    
    def Procurar(self):
        """Abre a janela de interação com o usuario para a escolha
        do arquivo contendo as etiqueta em estrutura compartivel.
        Formatos aceitos:
            '##R': "Arq. separado por vírgula (*.##R)",
            'csv': "Arq. separado por vírgula (*.csv)",
            'xlsx': "Arquivo excel (*.xlsx)",
            'txt': "Arquivo de texto (*.txt)"
        
        HABILITA O BOTÃO PROCESSAR E SALVA O CAMINHA DO ARQUIVO NO LABEL 'lblPathArquivo'
        """
        
        formatos = {
            '##R': "Arq. separado por vírgula (*.##R)",
            'csv': "Arq. separado por vírgula (*.csv)",
            'xlsx': "Arquivo excel (*.xlsx)",
            'txt': "Arquivo de texto (*.txt)"
        }
        try:
            arquivo = QFileDialog.getOpenFileName(self, "Escolha um arquivo compativel", "",\
                (f"{formatos['xlsx']};;{formatos['csv']};;{formatos['txt']};;{formatos['##R']}"))
            
            if self.verificarArquivo(arquivo[0]):
                self.lblPathArquivo.setText(arquivo[0])
                self.btnProcessar.setEnabled(True)
                
        except Exception as e:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle("ERRO")
            msg_box.setText("Erro ao abrir o arquivo!")
            msg_box.setDetailedText(str(e))
            return_value = msg_box.exec()

    def Processar(self):  
        self.btnFiltrar.setEnabled(True)
        
        self._listaEtiquetasGeral = Relatorio.extrairReltorioExpedicaoLocal(self.lblPathArquivo.text())
        
        self.PreencherTabela(self._listaEtiquetasGeral)
        self.lblTotalEtiquetas.setText(f"{len(self._listaEtiquetasGeral)} Etiquetas Encontradas")
        
    def PreencherTabela(self, dados):
        header = self.tbl_resumo.horizontalHeader()       
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)

        rowCount = self.tbl_resumo.rowCount()
        cont = 0
        while cont < len(dados):
            self.tbl_resumo.insertRow(rowCount)
            for coluna in range(len(dados[0])):
                #print(dados[cont][coluna])
                self.tbl_resumo.setItem(rowCount, coluna, QtWidgets.QTableWidgetItem(str(dados[cont][coluna]))) 
            cont += 1
        

        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = Main_Window()
    w.show()
    sys.exit(app.exec_())
