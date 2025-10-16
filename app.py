# app.py
from flask import Flask, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import io
import numpy as np 
from decimal import Decimal, ROUND_HALF_UP 

app = Flask(__name__)
# Aumenta o limite para múltiplos arquivos
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 * 5 

# Mapeamento de UF para Nome Completo (Padrão IBGE)
UF_MAP = {
    'AC': 'Acre', 'AL': 'Alagoas', 'AP': 'Amapá', 'AM': 'Amazonas', 'BA': 'Bahia', 
    'CE': 'Ceará', 'DF': 'Distrito Federal', 'ES': 'Espírito Santo', 'GO': 'Goiás', 
    'MA': 'Maranhão', 'MT': 'Mato Grosso', 'MS': 'Mato Grosso do Sul', 'MG': 'Minas Gerais', 
    'PA': 'Pará', 'PB': 'Paraíba', 'PR': 'Paraná', 'PE': 'Pernambuco', 'PI': 'Piauí', 
    'RJ': 'Rio de Janeiro', 'RN': 'Rio Grande do Norte', 'RS': 'Rio Grande do Sul', 
    'RO': 'Rondônia', 'RR': 'Roraima', 'SC': 'Santa Catarina', 'SP': 'São Paulo', 
    'SE': 'Sergipe', 'TO': 'Tocantins'
}


@app.route('/')
def index():
    """Renderiza a página inicial para upload (index.html)."""
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar_arquivos():
    
    # 1. MAPEAMENTO DE CONTAS (Lógica de preenchimento da coluna 'conta')
    MAPA_CONTA = {
        'Decarion (Monaco Metais)': 'DECARION TORNEIRAS',
        'Gs Torneiras': 'GS TORNEIRAS',
        'Via Flix (Matriz)': 'VIA FLIX'
    }

    # Lista para armazenar todos os DataFrames processados
    all_processed_dfs = []
    
    entry_index = 0
    
    # Colunas finais necessárias no gabarito
    COLUNAS_GABARITO_FINAL = [
        'SKU PRINCIPAL', 'SKU', 'Data da venda', 'EMISSAO', 'N.º de venda', 
        'origem', '# de anúncio', 'tipo de anuncio', 'Venda por publicidade', 
        'Forma de entrega', 'Preço unitário de venda do anúncio (BRL)', 
        'Unidades', 'Receita por produtos (BRL)', 'Envio Seller', 'TARIFA', 
        'Tarifa de venda e impostos (BRL)', 'conta', 'Estado'
    ]
    
    try:
        # 2. PROCESSA TODOS OS GRUPOS DE UPLOAD ENVIADOS PELO JAVASCRIPT
        while request.form.get(f'marketplace_{entry_index}') is not None:
            
            marketplace = request.form.get(f'marketplace_{entry_index}')
            conta_selecionada = request.form.get(f'conta_selecionada_{entry_index}')
            
            conta_final = MAPA_CONTA.get(conta_selecionada, 'OUTRAS')
            
            files = request.files.getlist(f'files_{entry_index}[]')
            
            if not files or files[0].filename == '':
                entry_index += 1
                continue

            # 3. PROCESSA CADA ARQUIVO NA LISTA
            for file in files:
                if not file.filename.endswith('.xlsx'):
                    continue 

                data_io = io.BytesIO(file.read())
                
                # --- DEFINIÇÃO DE HEADERS E COLUNAS POR MARKETPLACE ---
                
                # VARIÁVEIS DE CONFIGURAÇÃO POR MARKETPLACE
                
                if marketplace == 'Mercado Livre':
                    header_row = 5 # Linha 6
                    
                    # Nomes das colunas no arquivo de ORIGEM (ML) - Usando os nomes EXATOS fornecidos
                    MAPPING = {
                        'N.º de venda': 'N.º de venda',
                        'Data da venda': 'Data da venda',
                        'Unidades': 'Unidades',
                        'Preço unitário de venda do anúncio (BRL)': 'Preço unitário de venda do anúncio (BRL)',
                        'Tarifas de envio (BRL)': 'Tarifas de envio (BRL)',
                        'Tarifa de venda e impostos (BRL)': 'Tarifa de venda e impostos (BRL)',
                        'Estado': 'Estado',
                        'Forma de entrega': 'Forma de entrega',
                        'SKU': 'SKU', 
                        'Estado.1': 'Estado',
                    }
                    
                    # Coluna do original que deve ser renomeada para 'tipo de anuncio'
                    RENAME_ORIGINAL = {'Tipo de anúncio': 'tipo de anuncio'}
                    
                    # Coluna do ML para pegar a Receita por produtos (será recalculada depois)
                    RECEITA_COL = 'Receita por produtos (BRL)' 
                    
                elif marketplace in ['Shein', 'Shopee']:
                    header_row = 0 # Linha 1
                    
                    # Mapeamento Shein/Shopee (baseado nas imagens)
                    MAPPING = {
                        'CODIGOSKU': 'SKU', # CHAVE PRINCIPAL (nome limpo)
                        'DATADAVEN': 'Data da venda', # Coluna A
                        'DATADEEMIS': 'EMISSAO_TEMP', # Coluna B (Novo)
                        'NUMERODOPEDIDONOECOMERCE': 'N.º de venda',
                        'PRECOUNIT': 'Preço unitário de venda do anúncio (BRL)',
                        'QUANTIDAD': 'Unidades',
                        'PRECOTOTAL': 'Receita por produtos (BRL)', # Usaremos esse valor diretamente
                        'FRETEPAGOPELOCLI': 'Tarifas de envio (BRL)', # Custo de Frete (Tarifa de envio)
                        'COMISSAOECOMERCE': 'Tarifa de venda e impostos (BRL)', # Comissão
                        'VALORDEDESCONTO': 'Valor_Desconto_Temp', # Novo para Shein/Shopee
                        'UF': 'Estado',
                        'ECOMMERCE': 'origem' # NOVO: Mapeia o nome do Marketplace do arquivo para 'origem'
                    }
                    
                    RENAME_ORIGINAL = {} # Não precisa de renomeação no corpo principal
                    RECEITA_COL = 'Preço total' # Usamos o total que já vem na planilha
                    
                else:
                    entry_index += 1
                    continue
                
                # 5. LÊ O ARQUIVO
                try:
                    data_frame_raw = pd.read_excel(data_io, header=header_row)
                    
                    # 6. LIMPEZA E PADRONIZAÇÃO DOS NOMES DAS COLUNAS (CORREÇÃO DO KEYERROR)
                    
                    # Função para limpar nomes: remove espaços, caracteres especiais e coloca em maiúsculas
                    def clean_col_name(col):
                        if isinstance(col, str):
                            # Remove parênteses, hífens, pontos, etc., e espaços
                            col = col.replace(' ', '').replace('(', '').replace(')', '').replace('.', '').replace('-', '').upper()
                            return col
                        return col
                    
                    # Aplica a limpeza APENAS se não for Mercado Livre (ML precisa dos nomes exatos)
                    if marketplace in ['Shein', 'Shopee']:
                         data_frame_raw.columns = [clean_col_name(col) for col in data_frame_raw.columns]


                    # 7. Renomeia e Seleciona Apenas as Colunas Mapeadas
                    
                    if marketplace in ['Shein', 'Shopee']:
                        # Para Shein/Shopee, as chaves do MAPPING JÁ ESTÃO LIMPAS.
                        df = data_frame_raw.rename(columns=MAPPING, inplace=False)
                    else:
                        # ML: Aplica renomeações ML
                        df = data_frame_raw.rename(columns=RENAME_ORIGINAL, inplace=False)
                        df = data_frame_raw.rename(columns=MAPPING, inplace=False)


                    # --- Correção do KeyError: Se a coluna Forma de entrega não for encontrada no ML, ela é adicionada como vazia.
                    if marketplace == 'Mercado Livre' and 'Forma de entrega' not in df.columns:
                        df['Forma de entrega'] = ''
                    # --- Fim da Correção ---
                    
                    # ADICIONA ORIGEM VAZIA PARA ML ANTES DE SELECIONAR (RESOLVE KEYERROR DE ORIGEM)
                    if marketplace == 'Mercado Livre' and 'origem' not in df.columns:
                         df['origem'] = ''


                    # Seleciona apenas as colunas mapeadas
                    cols_needed = list(MAPPING.values())
                    df = df[cols_needed] 
                    
                except KeyError as e:
                    # Isso deve capturar erros de colunas não encontradas nos arquivos
                    return f"Erro de coluna no arquivo '{file.filename}'. Coluna chave não encontrada: {e}. Verifique se o cabeçalho está na linha correta ({header_row + 1}).", 500
                except Exception as e:
                    return f"Erro ao ler o arquivo '{file.filename}': {e}", 500


                # CORREÇÃO CRÍTICA DO NÚMERO DE VENDA
                df['N.º de venda'] = df['N.º de venda'].apply(
                    lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() else x
                )

                # 8. CONVERSÕES E CÁLCULOS BASE
                df['Unidades'] = pd.to_numeric(df['Unidades'], errors='coerce').fillna(0)
                df['Preço unitário de venda do anúncio (BRL)'] = pd.to_numeric(
                    df['Preço unitário de venda do anúncio (BRL)'], errors='coerce'
                ).fillna(0)
                
                # Prepara colunas para o cálculo Decimal
                original_freight = pd.to_numeric(df.get('Tarifas de envio (BRL)', 0), errors='coerce').fillna(0)
                original_freight_decimal = original_freight.astype(str).apply(Decimal)
                unidades_decimal = df['Unidades'].apply(Decimal)
                
                df['Tarifas de envio (BRL)'] = original_freight 
                df['Tarifa de venda e impostos (BRL)'] = pd.to_numeric(df['Tarifa de venda e impostos (BRL)'], errors='coerce').fillna(0)

                # --- LÓGICA POR MARKETPLACE (APLICAÇÃO DE REGRAS) ---
                
                if marketplace == 'Mercado Livre':
                    # LÓGICA DE DATAS (ML)
                    try:
                        month_map = { 'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04', 'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08', 'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12' }
                        dates = df['Data da venda'].astype(str).str.replace(r'\s\d{1,2}:\d{2}\s*hs\.$', '', regex=True).str.strip()
                        for pt_month, num_month in month_map.items():
                            dates = dates.str.replace(f' de {pt_month} de ', f'/{num_month}/', regex=False)
                        df['Data da venda'] = pd.to_datetime(dates, format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y')
                    except Exception as e:
                        print(f"Aviso: Falha na formatação da coluna 'Data da venda' (ML). Erro: {e}")

                    # LÓGICA DE CUSTO COMPLEXA (ML)
                    
                    unit_price = df['Preço unitário de venda do anúncio (BRL)']
                    
                    # CORREÇÃO: Verifica se 'Forma de entrega' existe antes de usar
                    if 'Forma de entrega' in df.columns:
                        df['Forma de entrega'] = df['Forma de entrega'].astype(str).str.strip()
                        is_flex = df['Forma de entrega'] == 'Mercado Envios Flex'
                        shipping_group_logic = ['Mercado Envios Full', 'Correios e pontos de envio', 'Pontos de Envio']
                        is_full_logic_group = df['Forma de entrega'].isin(shipping_group_logic)

                        # --- DEFINIÇÕES DE CUSTO FIXO FLEX (CORREÇÃO DE NAMEROR) ---
                        cost_flex_high_calc = Decimal("-9.11") 
                        cost_flex_low_calc = Decimal("-1.10") 
                        
                        # CUSTO UNITÁRIO HISTÓRICO (Minimização) - PREVENÇÃO CONTRA DIVISÃO POR ZERO
                        # A divisão só é feita se 'Unidades' > 0 e for item com custo original > 0
                        is_original_cost = (original_freight.abs() > 0) & (df['Unidades'] > 0)
                        is_valid_cost = (original_freight.abs() < 100) # Máscara de custo válido
                        
                        df['Unit_Cost_Temp'] = np.where(
                            is_original_cost & is_valid_cost & (df['Unidades'] > 0), # Condição de divisão por zero
                            (original_freight_decimal / unidades_decimal).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)),
                            np.nan 
                        )
                        max_cost_group = df.groupby(['SKU', 'Forma de entrega'])['Unit_Cost_Temp'].transform('max')
                        max_cost_total = (max_cost_group.apply(Decimal) * unidades_decimal).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
                        
                        sale_counts = df.groupby('N.º de venda')['N.º de venda'].transform('count')
                        is_sub_item_to_calculate = (sale_counts > 1) & (original_freight == 0)

                        
                        cond_flex_high = is_flex & (unit_price >= 79.00) 
                        cond_flex_low = is_flex & (unit_price <= 78.99) 
                        flex_calculated_cost_unit = np.select(
                            [cond_flex_high, cond_flex_low],
                            [cost_flex_high_calc, cost_flex_low_calc],
                            default=Decimal("0.00")
                        ) 
                        
                        cond_full_force_zero = is_full_logic_group & (unit_price <= 78.99)
                        cond_flex_apply_fixed = is_flex & (original_freight == 0)
                        cond_apply_best_cost = (max_cost_total.astype(float) > original_freight) & (original_freight < 0)
                        cond_full_apply_min = is_sub_item_to_calculate & is_full_logic_group
                        
                        final_cost_decimal_array = np.select(
                            [
                                cond_flex_apply_fixed,      
                                cond_apply_best_cost,       
                                cond_full_apply_min,        
                            ], 
                            [
                                flex_calculated_cost_unit,         
                                max_cost_total,                    
                                max_cost_total,                    
                            ],
                            default=original_freight_decimal 
                        )
                        
                        # Sobrescrita para o Zeramento (FULL <= 78.99)
                        final_cost_series = pd.Series(final_cost_decimal_array, index=df.index)
                        final_cost_series = np.where(
                            cond_full_force_zero,
                            Decimal("0.00"),
                            final_cost_series
                        )
                        
                        df['Tarifas de envio (BRL)'] = pd.Series(final_cost_series, index=df.index).apply(lambda x: Decimal(str(x))).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)).astype(float)
                        df.drop(columns=['Unit_Cost_Temp'], inplace=True) 
                    
                    # FIM DA LÓGICA COMPLEXA (ML)
                    
                    # CALCULA RECEITA ML: Unidades * Preço Unitário
                    df['Receita por produtos (BRL)'] = (df['Unidades'] * df['Preço unitário de venda do anúncio (BRL)']).round(2)
                    
                    # Colunas específicas do ML que não vêm no arquivo, mas precisam existir no DF
                    df['# de anúncio'] = ''
                    df['tipo de anuncio'] = ''
                    df['Venda por publicidade'] = ''
                    
                
                elif marketplace in ['Shein', 'Shopee']:
                    # LÓGICA DE DATAS (Shein/Shopee)
                    # Formata as datas de venda e emissão para DD/MM/AAAA
                    df['Data da venda'] = pd.to_datetime(df['Data da venda'], errors='coerce').dt.strftime('%d/%m/%Y')
                    df['EMISSAO'] = pd.to_datetime(df['EMISSAO_TEMP'], errors='coerce').dt.strftime('%d/%m/%Y')
                    df.drop(columns=['EMISSAO_TEMP'], inplace=True) 

                    # -----------------------------------------------------------
                    # INÍCIO DA LÓGICA ESPECÍFICA SHEIN/SHOPEE
                    # -----------------------------------------------------------

                    # 1. CONVERSÃO DE UF PARA ESTADO COMPLETO
                    df['Estado'] = df['Estado'].astype(str).str.upper().map(UF_MAP).fillna(df['Estado'])

                    # 2. NEGATIVAR COMISSÃO (COLUNA 'Tarifa de venda e impostos (BRL)')
                    # Deve ser numérico e negativo
                    df['Tarifa de venda e impostos (BRL)'] = pd.to_numeric(df['Tarifa de venda e impostos (BRL)'], errors='coerce').fillna(0) * -1
                    
                    # 3. DESCONTAR 'Valor de desconto' DA RECEITA
                    df['Valor_Desconto_Temp'] = pd.to_numeric(df['Valor_Desconto_Temp'], errors='coerce').fillna(0)
                    df['Receita por produtos (BRL)'] = pd.to_numeric(df['Receita por produtos (BRL)'], errors='coerce').fillna(0) - df['Valor_Desconto_Temp']

                    # Garante que a Receita fique em 2 casas decimais
                    df['Receita por produtos (BRL)'] = df['Receita por produtos (BRL)'].round(2)
                    
                    # 4. LIMPAR COLUNAS TEMPORÁRIAS E INSERIR COLUNAS VAZIAS (Gabarito)
                    df.drop(columns=['Valor_Desconto_Temp'], inplace=True) 
                    
                    # Colunas específicas que não vêm no arquivo, mas precisam existir no DF
                    df['# de anúncio'] = ''
                    df['tipo de anuncio'] = ''
                    df['Venda por publicidade'] = ''
                    df['Tarifas de envio (BRL)'] = df.get('Tarifas de envio (BRL)', pd.Series(0.0, index=df.index)) # Mantém a coluna de frete (se existir)
                    
                # --- FIM DA LÓGICA POR MARKETPLACE ---
                
                # --- REGRAS COMUNS (APLICADAS A TODOS) ---

                # CORREÇÃO: Duplicação do SKU para SKU PRINCIPAL
                df['SKU PRINCIPAL'] = df['SKU'] 
                
                df['Envio Seller'] = df.get('Tarifas de envio (BRL)', df.get('Envio Seller', pd.Series(0.0, index=df.index))) # Se o custo foi recalculado ou mantido, é atribuído aqui
                df['EMISSAO'] = df['Data da venda'] # O valor da emissão já foi tratado acima, mas garantimos que exista
                
                # CRUCIAL: Se for ML (origem vazia), preenche com o marketplace selecionado.
                df['origem'] = df['origem'].apply(lambda x: marketplace if pd.isna(x) or x == '' else x)
                
                df['conta'] = conta_final # Coluna 'conta' preenchida com a lógica do MAPA_CONTA

                df['TARIFA'] = 0 # Valor Fixo
                
                # SELEÇÃO FINAL DAS COLUNAS E REORDENAMENTO
                df = df[COLUNAS_GABARITO_FINAL]
                
                all_processed_dfs.append(df)
            
            # Passa para o próximo grupo de upload
            entry_index += 1
            
        # 4. CONCATENAÇÃO FINAL
        if not all_processed_dfs:
            return "Nenhum arquivo válido foi processado.", 400

        final_df = pd.concat(all_processed_dfs, ignore_index=True)

        # NOVA LÓGICA: Ordenar por 'conta' e depois por 'SKU PRINCIPAL'
        final_df = final_df.sort_values(by=['conta', 'SKU PRINCIPAL'], ascending=[True, True])
        
        # 5. SALVAR E FORMATAR (UM ÚNICO ARQUIVO DE SAÍDA)
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        final_df.to_excel(writer, index=False, sheet_name='Planilha_Combinada')
        writer.close()
        output.seek(0)

        # 6. AJUSTAR LARGURA DE COLUNAS E APLICAR FILTROS (No arquivo final)
        wb = load_workbook(output)
        ws = wb.active 
        ws.auto_filter.ref = ws.dimensions
        
        for col in ws.columns:
            max_length = 0
            column = get_column_letter(col[0].column) 
            
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        # --- FIM DO PROCESSAMENTO ---

        # 7. Envia o arquivo processado para o download do usuário
        return send_file(
            final_output,
            as_attachment=True,
            download_name=f'planilha_combinada.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
            
    except Exception as e:
        # Se ocorrer um erro durante o processamento de qualquer arquivo
        return f"Ocorreu um erro no processamento. Detalhes: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)
