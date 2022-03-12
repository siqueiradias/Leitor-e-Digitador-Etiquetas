#! /usr/bin/env python
# -*- coding: utf-8 -*-


"""Relatorio.py: Lê os relatórios TOTVS
                e Expedição_Local."""

__author__ = "Diogo Dias"
__version__ = "0.0.1"
__email__ = "diogo@frangoamericano.com"
__status__ = "Em desenvolvimento"


import pandas as pd
import os


class Relatorio():
    @staticmethod
    def extracaoRelatorioTOTVS(file, tipo):
        """
        Metodo responsavel por fazer a extração dos
        Dados de Relatórios do Sistema TOTVS
        """
        print("-"*20+" Extraindo dados "+"-"*20)
        arquivo = open(file, "r", encoding= "iso8859-1") #encoding= "iso8859-1" compativel o Relatório TOTVS
        totalCaixa = 0
        listDados = list()
        if tipo == "devolucao" or tipo == 'expedicao':
            for linha in arquivo:
                dados = linha.split()
                if len(dados)>4 and dados[0].isnumeric() and dados[-4].replace(",","").isnumeric():
                    print("{} = {} cx".format(dados[0], dados[-4].replace(",","")))
                    totalCaixa += int(dados[-4].replace(",",""))
                    listDados.append((dados[0], int(dados[-4].replace(",",""))))
        print("Total de Caixas: {}".format(totalCaixa))
        return listDados
    
    @staticmethod
    def extrairReltorioExpedicaoLocal(file):
        lista_etiquetas = list()
        
        try:
            df = pd.read_excel(file)
            
            lista_etiquetas = df.values.tolist()
            
            print("::::: Lista de etiquetas extraidas :::::")
            for item in lista_etiquetas:
                print(item)
            return lista_etiquetas
        except ImportError as e:
            print(f"Falha de importação/dependecia:\n {e} ")
            if str(e).find('xlrd'):
                try:
                    print("Será instalado a dependencia 'xlrd'.\n")
                    os.system('pip install xlrd==1.2.0')
                except Exception as e:
                    print("Erro na instalação da dependencia 'xlrd'.")
                
        except Exception as e:
            print(f"Erro na função de extrair etiquetas: {e}")
    