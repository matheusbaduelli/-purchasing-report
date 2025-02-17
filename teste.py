import pandas as pd
import sqlite3



planilha1 = pd.read_excel('Estrutura_Produtos_30_01_2025.xls',usecols="A:AL")
planilha2 = pd.read_excel('Estrutura_Produtos_30_01_20252.xls',usecols="A:AL")


planilhaConcatenada = pd.concat([planilha1,planilha2],ignore_index=False)

print(planilhaConcatenada)

planilhaConcatenada.to_excel('resultado.xlsx',index=False)

lista_tecnica = planilhaConcatenada
lista_pedidos = pd.read_excel('Lista de pedidos.xlsx')



conn = sqlite3.connect('Lista.db')
cursor = conn.cursor()

lista_tecnica.to_sql('lista_tecnica', conn, if_exists='replace', index=False)
lista_pedidos.to_sql('lista_de_pedidos',conn,if_exists='replace',index=False)

cursor.execute("SELECT lista_de_pedidos.Pedido,lista_tecnica.Produto,lista_tecnica.Descricao,lista_tecnica.Grupo,lista_tecnica.Cod,lista_tecnica.Codigo,lista_tecnica.Qtde,lista_tecnica.Descricao_comp,lista_de_pedidos.Qtd FROM lista_de_pedidos INNER JOIN lista_tecnica ON lista_de_pedidos.Codigo = lista_tecnica.Produto;")



nova_lista = []
for linha in cursor.fetchall():
    concatenado = f'{linha[0]}{linha[1]}{linha[4]}'
    lista = {"Pedido":linha[0],"Produto":linha[1],"Descricao":linha[2],"Grupo":linha[3],"Cod":linha[4],"Codigo":linha[5],"Qtd":linha[6],"Descricao_comp":linha[7],"Qtd total":float(linha[8]*linha[6]),"CodConcatenado":concatenado}
    nova_lista.append(lista)
    
df =  pd.DataFrame(nova_lista)
df.drop_duplicates(subset=['CodConcatenado']).to_excel("resultado.xlsx",index=False)



conn.commit()
conn.close()