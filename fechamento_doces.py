import pandas as pd
import os

# Leitura da planilha na aba "Vendas"
file_path = "base/Polá_Doces_2024.xlsx"
df = pd.read_excel(file_path, sheet_name="Vendas")

# Verificar nomes das colunas
print("Colunas disponíveis na planilha:", df.columns)

# Remover espaços extras dos nomes das colunas
df.columns = df.columns.str.strip()

# Filtrando por "Pago" igual a "Não"
df_nao_Pago = df[df['Pago'] == 'Não']

# Diretório para salvar os arquivos de saída
output_dir = "resultado"
os.makedirs(output_dir, exist_ok=True)

# Iterando por cada cliente
for cliente, df_cliente in df_nao_Pago.groupby('Cliente'):
    mensagem = f"Bom dia, tudo bem?\nQuinto dia útil chegouuu ksksksks\nTo aqui pra passar quanto ficou sua conta no *Polá Doces* este mês!\n\n"

    valor_total = 0
    for produto, df_produto in df_cliente.groupby('Produto'):
        quantidade_total = df_produto['Quantidade'].sum()
        valor_produto = df_produto['Valor'].sum()
        mensagem += f"{quantidade_total} {produto} = R${valor_produto:.2f}\n"
        valor_total += valor_produto
    
    mensagem += f"\n*Total=* R${valor_total:.2f}\n\n"
    mensagem += "Chave Pix: *549.249.088-51* Gustavo Polastrini da Cruz - Nubank\n"

    # Salvando a mensagem em um arquivo .txt
    file_name = f"{output_dir}/{cliente}.txt"
    with open(file_name, "w", encoding="utf-8") as file:
        file.write(mensagem)

print("Mensagens geradas e salvas na pasta 'resultado'.")
