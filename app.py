import streamlit as st
import pandas as pd

# Título da aplicação
st.title("Sistema de De-Para de Pedidos")

# Função para normalizar nomes de colunas (ignorar maiúsculas/minúsculas e espaços)
def normalizar_nome_coluna(colunas):
    return [col.strip().lower() for col in colunas]

# Função para limpar valores (remover espaços extras e caracteres indesejados)
def limpar_valor(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

# Função para verificar se o valor é "não nulo" (não vazio após limpeza)
def valor_nao_nulo(valor):
    return limpar_valor(valor) != ""

# Função para carregar e processar os arquivos
def processar_arquivos(file1, file2):
    # Ler os arquivos Excel
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    
    # Normalizar os nomes das colunas para comparação
    colunas_df1 = normalizar_nome_coluna(df1.columns)
    colunas_df2 = normalizar_nome_coluna(df2.columns)
    
    # Verificar se as colunas esperadas existem (ignorando maiúsculas/minúsculas e espaços)
    if 'pedido print one' not in colunas_df1:
        raise ValueError("A planilha do Sistema 1 não contém a coluna 'Pedido Print One' (verifique maiúsculas/minúsculas e espaços).")
    if 'número do pedido' not in colunas_df2:
        raise ValueError("A planilha do Sistema 2 não contém a coluna 'Número do Pedido' (verifique maiúsculas/minúsculas e espaços).")
    
    # Procurar a coluna que contém "cliente" no nome, mas não "cliente - nome"
    coluna_cliente = None
    for col, col_norm in zip(df1.columns, colunas_df1):
        if 'cliente' in col_norm and 'cliente - nome' not in col_norm:
            coluna_cliente = col
            break
    
    if coluna_cliente is None:
        raise ValueError("A planilha do Sistema 1 não contém uma coluna com 'Cliente' no nome (exceto 'Cliente - Nome').")
    
    # Mapear os nomes originais das colunas para os normalizados
    colunas_originais_df1 = dict(zip(colunas_df1, df1.columns))
    colunas_originais_df2 = dict(zip(colunas_df2, df2.columns))
    
    # Renomear as colunas para um nome padrão
    df1 = df1.rename(columns={colunas_originais_df1['pedido print one']: 'Numero_Pedido'})
    df2 = df2.rename(columns={colunas_originais_df2['número do pedido']: 'Numero_Pedido'})
    
    # Verificar se a renomeação foi bem-sucedida
    if 'Numero_Pedido' not in df1.columns or 'Numero_Pedido' not in df2.columns:
        raise ValueError("Erro na renomeação das colunas. Verifique os nomes das colunas nas planilhas.")
    
    # Limpar os valores das colunas antes da comparação
    df1['Numero_Pedido'] = df1['Numero_Pedido'].apply(limpar_valor)
    df2['Numero_Pedido'] = df2['Numero_Pedido'].apply(limpar_valor)
    
    # Filtrar apenas os pedidos com Numero_Pedido não nulo no Sistema 1
    df1_filtrado = df1[df1['Numero_Pedido'].apply(valor_nao_nulo)]
    
    # Criar conjuntos para comparação
    pedidos_sistema1 = set(df1_filtrado['Numero_Pedido'])
    pedidos_sistema2 = set(df2['Numero_Pedido'])
    
    # Comparação dos pedidos
    pedidos_corretos = pedidos_sistema1.intersection(pedidos_sistema2)
    pedidos_inconsistentes = pedidos_sistema1.symmetric_difference(pedidos_sistema2)
    
    # Pedidos que estão apenas no Sistema 1 (precisam de revisão)
    pedidos_apenas_sistema1 = pedidos_sistema1 - pedidos_sistema2
    # Pedidos que estão apenas no Sistema 2 (precisam de revisão)
    pedidos_apenas_sistema2 = pedidos_sistema2 - pedidos_sistema1
    
    # Atualizar DataFrame com status (aplicar a todos os registros, mesmo os filtrados)
    df1['Status_Comparacao'] = df1['Numero_Pedido'].apply(
        lambda x: 'OK' if x in pedidos_corretos else 'Não Lançado' if x in pedidos_inconsistentes else 'Sem Pedido'
    )
    
    # Filtrar clientes com "THE BEST" no nome e que NÃO possuem Numero_Pedido
    df1[coluna_cliente] = df1[coluna_cliente].astype(str)  # Garantir que a coluna seja string
    clientes_the_best = df1[
        (df1[coluna_cliente].str.contains('THE BEST', case=False, na=False)) & 
        (~df1['Numero_Pedido'].apply(valor_nao_nulo))
    ][[coluna_cliente]].drop_duplicates()
    
    # Renomear a coluna de volta ao nome original
    df1 = df1.rename(columns={'Numero_Pedido': colunas_originais_df1['pedido print one']})
    
    return (df1, len(pedidos_corretos), len(pedidos_inconsistentes), 
            pedidos_corretos, pedidos_apenas_sistema1, pedidos_apenas_sistema2, clientes_the_best, coluna_cliente)

# Upload dos arquivos
st.subheader("Upload das Planilhas")
col1, col2 = st.columns(2)

with col1:
    arquivo1 = st.file_uploader("Planilha Sistema 1", type=['xlsx'])
    
with col2:
    arquivo2 = st.file_uploader("Planilha Sistema 2", type=['xlsx'])

# Variáveis para armazenar resultados
resultado_df = None
contagem_corretos = 0
contagem_inconsistentes = 0
pedidos_corretos = set()
pedidos_apenas_sistema1 = set()
pedidos_apenas_sistema2 = set()
clientes_the_best = pd.DataFrame()
coluna_cliente = None

# Processar quando ambos arquivos forem carregados
if arquivo1 and arquivo2:
    with st.spinner('Processando os dados...'):
        try:
            (resultado_df, contagem_corretos, contagem_inconsistentes, 
             pedidos_corretos, pedidos_apenas_sistema1, pedidos_apenas_sistema2, 
             clientes_the_best, coluna_cliente) = processar_arquivos(arquivo1, arquivo2)
            
            # Mostrar resultados
            st.success('Processamento concluído!')
            
            # Botões com contagens
            col3, col4 = st.columns(2)
            with col3:
                st.button(f"Pedidos Corretos: {contagem_corretos}", disabled=True)
            with col4:
                st.button(f"Pedidos Inconsistentes: {contagem_inconsistentes}", disabled=True)
            
            # Exibir pedidos que constam em ambas as planilhas
            st.subheader("Pedidos que Constam em Ambas as Planilhas (OK)")
            if pedidos_corretos:
                st.write(list(pedidos_corretos))
            else:
                st.write("Nenhum pedido consta em ambas as planilhas.")
            
            # Exibir pedidos que precisam de revisão
            st.subheader("Pedidos que Precisam de Revisão")
            
            st.write("**Pedidos apenas no Sistema 1 (não estão no Sistema 2):**")
            if pedidos_apenas_sistema1:
                st.write(list(pedidos_apenas_sistema1))
            else:
                st.write("Nenhum pedido exclusivo no Sistema 1.")
            
            st.write("**Pedidos apenas no Sistema 2 (não estão no Sistema 1):**")
            if pedidos_apenas_sistema2:
                st.write(list(pedidos_apenas_sistema2))
            else:
                st.write("Nenhum pedido exclusivo no Sistema 2.")
            
            # Exibir clientes com "THE BEST" no nome e sem Pedido Print One
            st.subheader("Clientes com 'THE BEST' no Nome e Sem 'Pedido Print One' (Sistema 1)")
            if not clientes_the_best.empty:
                st.write(clientes_the_best[coluna_cliente].tolist())
            else:
                st.write("Nenhum cliente com 'THE BEST' no nome e sem 'Pedido Print One' foi encontrado.")
            
            # Mostrar tabela resultante
            st.subheader("Resultado da Comparação (Planilha Sistema 1)")
            st.dataframe(resultado_df)
            
            # Download do resultado
            st.subheader("Download do Resultado")
            # Exportar como Excel
            output = pd.ExcelWriter('resultado_depara_pedidos.xlsx')
            resultado_df.to_excel(output, index=False)
            output.close()
            
            with open('resultado_depara_pedidos.xlsx', 'rb') as f:
                st.download_button(
                    label="Baixar Planilha Atualizada",
                    data=f,
                    file_name="resultado_depara_pedidos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro ao processar os arquivos: {str(e)}")

# Instruções
st.sidebar.header("Instruções")
st.sidebar.write("""
1. Faça upload das duas planilhas Excel.
2. A planilha do Sistema 1 deve ter a coluna 'Pedido Print One' e uma coluna com 'Cliente' no nome (exceto 'Cliente - Nome').
3. A planilha do Sistema 2 deve ter a coluna 'Número do Pedido'.
4. O sistema comparará os pedidos entre as duas colunas, considerando apenas pedidos com 'Pedido Print One' preenchido.
5. Resultados:
   - 'OK' para pedidos presentes em ambos os sistemas.
   - 'Não Lançado' para pedidos ausentes em um dos sistemas.
   - 'Sem Pedido' para registros sem 'Pedido Print One'.
6. Veja os pedidos corretos, os que precisam de revisão e os clientes com 'THE BEST' no nome (sem 'Pedido Print One') na tela.
7. Use os botões para ver as contagens.
8. Baixe a planilha atualizada com a nova coluna 'Status_Comparacao'.
""")