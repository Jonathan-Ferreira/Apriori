import pandas as pd
from mlxtend.preprocessing import TransactionEncoder
from mlxtend.frequent_patterns import apriori, association_rules

path = 'C:/Users/SDE/Desktop/Codes/Python/Apriori/apriori.xlsx'
df = pd.read_excel(path)

# Agrupar os segmentos por Nro. Único
df_agrupado = df.groupby(df.iloc[:,0])[df.columns[1]].apply(list).reset_index()

# TransactionEncoder para transformar a lista em uma matriz binária (pré-requisito do algoritmo Apriori)
te = TransactionEncoder()
df_transformado = te.fit_transform(df_agrupado.iloc[:,1])
df_binario = pd.DataFrame(df_transformado, columns=te.columns_)
print(df_binario)

# Aplicação do algoritmo apriori
itemset_frequente = apriori(df_binario, min_support=0.01, use_colnames=True)

# Geração das Regras de associação
regra = association_rules(itemset_frequente, metric="confidence", min_threshold=0.8)

# Print dos resultados
print(itemset_frequente)
print(regra)

with pd.ExcelWriter("resultado_apriori.xlsx") as writer:
    # Salvando os itemsets
    itemset_frequente.to_excel(writer, sheet_name="Itemset", index=False)
    
    # Salvando as regras de associação com as métricas de avaliação de resultado
    regra.to_excel(writer, sheet_name="Regras de Associação", index=False)

print("Resultado exportado para 'resultado_apriori.xlsx'")

