import pandas as pd
import os

# Carregar a planilha Excel
file_path = 'C:/Users/lucas.nogueira/Downloads/rel2203.xlsx'

# Verificar se o arquivo existe
if not os.path.exists(file_path):
    print(f"Erro: O arquivo '{file_path}' não foi encontrado.")
else:
    try:
        df = pd.read_excel(file_path)

        # Remover linhas completamente vazias
        df = df.dropna(how='all')

        # Converter colunas necessárias para string para evitar erros com métodos de string
        df['Contrato Gerado'] = df['Contrato Gerado'].astype(str)
        df['Contrato Assinado'] = df['Contrato Assinado'].astype(str)

        # Converter a coluna de data para datetime, assumindo que a coluna de data se chama 'Data'
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')

        # Remover linhas onde a data é nula
        df = df.dropna(subset=['Data'])

        # Agrupar dados por dia
        df['Dia'] = df['Data'].dt.date

        # Função para calcular o número total de CPF/CNPJ que se repetem
        def calcular_cpf_cnpj_repetidos(sub_df):
            cpf_cnpj_ocorrencias = sub_df.groupby(['CPF/CNPJ', 'Codigo Bem']).size().reset_index(name='Ocorrências')
            return cpf_cnpj_ocorrencias[cpf_cnpj_ocorrencias['Ocorrências'] > 1].shape[0]

        # Função para listar CPF/CNPJ que se repetem e a quantidade de ocorrências
        def listar_cpf_cnpj_ocorrencias(sub_df):
            cpf_cnpj_ocorrencias = sub_df.groupby(['CPF/CNPJ', 'Codigo Bem']).size().reset_index(name='Ocorrências')
            return cpf_cnpj_ocorrencias[cpf_cnpj_ocorrencias['Ocorrências'] > 1]

        # Função para calcular a quantidade de CPF/CNPJ que se repetem e tem pelo menos um contrato gerado
        def calcular_cpf_cnpj_contrato_gerado(sub_df):
            cpf_cnpj_ocorrencias_com_contrato = sub_df[sub_df.duplicated(subset=['CPF/CNPJ', 'Codigo Bem'], keep=False) & sub_df['Contrato Gerado'].notna() & (sub_df['Contrato Gerado'].str.strip() != '')]
            cpf_cnpj_repetidos_contrato_gerado = cpf_cnpj_ocorrencias_com_contrato[['CPF/CNPJ', 'Codigo Bem']].drop_duplicates()
            return cpf_cnpj_repetidos_contrato_gerado['CPF/CNPJ'].nunique()

        # Função para listar CPF/CNPJ com ocorrência repetida que possuem pelo menos um contrato gerado
        def listar_cpf_cnpj_repetidos_contrato_gerado(sub_df):
            cpf_cnpj_ocorrencias_com_contrato = sub_df[sub_df.duplicated(subset=['CPF/CNPJ', 'Codigo Bem'], keep=False) & sub_df['Contrato Gerado'].notna() & (sub_df['Contrato Gerado'].str.strip() != '')]
            return cpf_cnpj_ocorrencias_com_contrato[['CPF/CNPJ', 'Codigo Bem']].drop_duplicates()

        # Inicializar DataFrame para resultados
        resultados = pd.DataFrame(columns=[
            'Dia', 'Total de Vendas', 'Contratos Gerados', 'Contratos Assinados', 
            'CPF/CNPJ Repetidos', 'Quantidade de CPF/CNPJ com Contrato Gerado'
        ])

        # Processar dados por dia
        for dia, sub_df in df.groupby('Dia'):
            total_vendas = sub_df.shape[0]
            contratos_gerados = sub_df[sub_df['Contrato Gerado'].notna() & (sub_df['Contrato Gerado'].str.strip() != '')].shape[0]
            contratos_assinados = sub_df[sub_df['Contrato Assinado'].notna() & (sub_df['Contrato Assinado'].str.strip() != '')].shape[0]
            cpf_cnpj_repetidos = calcular_cpf_cnpj_repetidos(sub_df)
            cpf_cnpj_contrato_gerado = calcular_cpf_cnpj_contrato_gerado(sub_df)

            resultados = pd.concat([resultados, pd.DataFrame([{
                'Dia': dia,
                'Total de Vendas': total_vendas,
                'Contratos Gerados': contratos_gerados,
                'Contratos Assinados': contratos_assinados,
                'CPF/CNPJ Repetidos': cpf_cnpj_repetidos,
                'Quantidade de CPF/CNPJ com Contrato Gerado': cpf_cnpj_contrato_gerado
            }])], ignore_index=True)

        # Salvar resultados em uma nova planilha Excel
        output_path = file_path.replace('.xlsx', '_consolidado_diario.xlsx')
        resultados.to_excel(output_path, index=False)

        print("Dados consolidados salvos em:", output_path)

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
