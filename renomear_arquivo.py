import os

FOLDER = r'/Users/tatianeyukarikawakami/Desktop/Sumire/Produtos Bloqueados/'

def rename(document,fabricante):
    folder = FOLDER
    for file_name in os.listdir(folder):
        if file_name == document:
            old_name = folder + document
            nem_name = folder + fabricante + '.xlsx'
            os.rename(old_name,nem_name)

tuple = (["ProdutosBloqCmpSaldo.xlsx","Loja-1.xlsx"],["ProdutosBloqCmpSaldo (1).xlsx","Loja-3.xlsx"],["ProdutosBloqCmpSaldo (2).xlsx","Loja-4.xlsx"],["ProdutosBloqCmpSaldo (3).xlsx","Loja-5.xlsx"],["ProdutosBloqCmpSaldo (4).xlsx","Loja-7.xlsx"],["ProdutosBloqCmpSaldo (5).xlsx","Loja-8.xlsx"],["ProdutosBloqCmpSaldo (6).xlsx","Loja-9.xlsx"],["ProdutosBloqCmpSaldo (7).xlsx","Loja-10.xlsx"],["ProdutosBloqCmpSaldo (8).xlsx","Loja-11.xlsx"],["ProdutosBloqCmpSaldo (9).xlsx","Loja-12.xlsx"],["ProdutosBloqCmpSaldo (10).xlsx","Loja-13.xlsx"],["ProdutosBloqCmpSaldo (11).xlsx","Loja-14.xlsx"],["ProdutosBloqCmpSaldo (12).xlsx","Loja-17.xlsx"],["ProdutosBloqCmpSaldo (13).xlsx","Loja-20.xlsx"],["ProdutosBloqCmpSaldo (14).xlsx","Loja-21.xlsx"],["ProdutosBloqCmpSaldo (15).xlsx","Loja-22.xlsx"],["ProdutosBloqCmpSaldo (16).xlsx","Loja-25.xlsx"],["ProdutosBloqCmpSaldo (17).xlsx","Loja-26.xlsx"],["ProdutosBloqCmpSaldo (18).xlsx","Loja-28.xlsx"],["ProdutosBloqCmpSaldo (19).xlsx","Loja-30.xlsx"],["ProdutosBloqCmpSaldo (20).xlsx","Loja-31.xlsx"],["ProdutosBloqCmpSaldo (21).xlsx","Loja-32.xlsx"],["ProdutosBloqCmpSaldo (22).xlsx","Loja-33.xlsx"],["ProdutosBloqCmpSaldo (23).xlsx","Loja-34.xlsx"],["ProdutosBloqCmpSaldo (24).xlsx","Loja-40.xlsx"],["ProdutosBloqCmpSaldo (25).xlsx","Loja-41.xlsx"],["ProdutosBloqCmpSaldo (26).xlsx","Loja-42.xlsx"],["ProdutosBloqCmpSaldo (27).xlsx","Loja-43.xlsx"],["ProdutosBloqCmpSaldo (28).xlsx","Loja-44.xlsx"],["ProdutosBloqCmpSaldo (29).xlsx","Loja-45.xlsx"],["ProdutosBloqCmpSaldo (30).xlsx","Loja-46.xlsx"],["ProdutosBloqCmpSaldo (31).xlsx","Loja-47.xlsx"],["ProdutosBloqCmpSaldo (32).xlsx","Loja-48.xlsx"],["ProdutosBloqCmpSaldo (33).xlsx","Loja-49.xlsx"],["ProdutosBloqCmpSaldo (34).xlsx","Loja-50.xlsx"],["ProdutosBloqCmpSaldo (35).xlsx","Loja-51.xlsx"],["ProdutosBloqCmpSaldo (36).xlsx","Loja-55.xlsx"],["ProdutosBloqCmpSaldo (37).xlsx","Loja-56.xlsx"],["ProdutosBloqCmpSaldo (38).xlsx","Loja-57.xlsx"],["ProdutosBloqCmpSaldo (39).xlsx","Loja-58.xlsx"],["ProdutosBloqCmpSaldo (40).xlsx","Loja-61.xlsx"],["ProdutosBloqCmpSaldo (41).xlsx","Loja-62.xlsx"],["ProdutosBloqCmpSaldo (42).xlsx","Loja-63.xlsx"])

for t in tuple:
    documento = t[0]
    fabricante = t[1]
    rename(document=documento, fabricante=fabricante)
    