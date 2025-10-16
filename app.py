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

@app.route('/')
def index():
    """Renderiza a página inicial para upload (index.html)."""
    # Renderiza o frontend com o layout simplificado
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

    # *** NOVO: DataFrame de referência para custos de SKU/Preço únicos processados ***
    cost_reference_df = pd.DataFrame()
    
    try:
        # 2. PROCESSA TODOS OS GRUPOS DE UPLOAD ENVIADOS PELO JAVASCRIPT
        while request.form.get(f'marketplace_{entry_index}') is not None:
            
            marketplace = request.form.get(f'marketplace_{entry_index}')
            conta_selecionada = request.form.get(f'conta_selecionada_{entry_index}')
            
            # Aplica a lógica de preenchimento da coluna 'conta'
            conta_final = MAPA_CONTA.get(conta_selecionada, 'OUTRAS')
            
            # Pega a lista de arquivos para este grupo específico
            files = request.files.getlist(f'files_{entry_index}[]')
            
            if not files or files[0].filename == '':
                entry_index += 1
                continue

            # 3. PROCESSA CADA ARQUIVO NA LISTA
            for file in files:
                if not file.filename.endswith('.xlsx'):
                    continue 

                data_io = io.BytesIO(file.read())
                
                # --- INÍCIO DO PROCESSAMENTO INDIVIDUAL ---
                
                # LÊ O ARQUIVO, COM O CABEÇALHO NA LINHA 6 (header=5), 'N.º de venda' como string
                df = pd.read_excel(data_io, header=5, dtype={'N.º de venda': str})
                
                
                # *** AJUSTE: Renomeação Flexível de Colunas CHAVE (SKU) ANTES DO FILTRO ***
                df = df.rename(columns={'Tipo de anúncio': 'tipo de anuncio'}, inplace=False) 
                
                # Mapeamento de possíveis nomes de SKU para o nome padronizado 'SKU'
                sku_col_map = {
                    'Cód. item': 'SKU',
                    'Cód. do item': 'SKU',
                    'SKU item': 'SKU',
                    'SKU da variação': 'SKU', # Típico para Mercado Livre
                    'Código': 'SKU' # Se for um nome genérico
                }
                
                # Aplica o mapeamento
                for original, new in sku_col_map.items():
                    if original in df.columns and new not in df.columns:
                        df.rename(columns={original: new}, inplace=True)
                
                # Garante que a coluna SKU existe para o filtro
                if 'SKU' not in df.columns:
                     df['SKU'] = ''
                # Fim da Renomeação Flexível
                
                
                # *** NOVO FILTRO: REMOVER LINHAS COM STATUS 'CANCELADO' (Reforçado para o nome da coluna) ***
                
                # 1. Normaliza os nomes das colunas para minúsculas e remove caracteres especiais
                df.columns = df.columns.str.strip()
                # Cria um mapeamento de nomes de colunas normalizados (sem espaços/pontos e em minúsculo)
                normalized_columns = {col.lower().replace(' ', '').replace('.', '').replace('_', '').replace('#', ''): col for col in df.columns}
                
                # Chaves de busca normalizadas
                STATUS_KEYS_NORMALIZED = ['statusdavenda', 'descriçãodostatus', 'status']
                
                # Procura o nome real da coluna no DataFrame
                status_col_name = next((normalized_columns[key] for key in STATUS_KEYS_NORMALIZED if key in normalized_columns), None)

                if status_col_name:
                    # Filtra: mantém apenas linhas onde o status NÃO contém 'cancel' (case-insensitive)
                    df = df[~df[status_col_name].astype(str).str.contains('cancel', case=False, na=False)].copy()
                # Fim do Novo Filtro
                
                
                # *** AJUSTE CRÍTICO: REMOVE A LINHA DE RESUMO DO CARRINHO (SKU VAZIO, MAS N.º DE VENDA PREENCHIDO) ***
                if not df.empty and 'SKU' in df.columns and 'N.º de venda' in df.columns:
                    # Filtra: mantém apenas linhas onde o SKU não é vazio/NaN (remove a linha de resumo)
                    df = df[df['SKU'].astype(str).str.strip().ne('')].copy()
                
                # Se o DataFrame ficar vazio após o filtro, pula.
                if df.empty:
                    continue

                # CORREÇÃO CRÍTICA DO NÚMERO DE VENDA
                df['N.º de venda'] = df['N.º de venda'].apply(
                    lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() else x
                )
                
                # 4. DEFINIÇÃO DO GABARITO E TRATAMENTO DE RENOMEAÇÕES
                COLUNAS_GABARITO_FINAL = [
                    'SKU PRINCIPAL', 'SKU', 'Data da venda', 'EMISSAO', 'N.º de venda', 
                    'origem', '# de anúncio', 'tipo de anuncio', 'Venda por publicidade', 
                    'Forma de entrega', 'Preço unitário de venda do anúncio (BRL)', 
                    'Unidades', 'Receita por produtos (BRL)', 'Envio Seller', 'TARIFA', 
                    'Tarifa de venda e impostos (BRL)', 'conta', 'Estado'
                ]
                
                df['Forma de entrega'] = df['Forma de entrega'].astype(str).str.strip() 
                
                # *** AJUSTE: Propagar dados de cabeçalho do carrinho (BFILL e FFILL) ***
                cols_to_fill = ['Forma de entrega', 'tipo de anuncio', 'Venda por publicidade', 'Estado']
                
                for col in cols_to_fill:
                    if col in df.columns:
                        # Passo 1: Converter strings vazias para NaN, garantindo que bfill/ffill funcione.
                        df[col] = df[col].replace(r'^\s*$', np.nan, regex=True)
                        
                        # 2. BFILL: Preenche NaN/vazio para trás (pega o valor de baixo)
                        df[col] = df.groupby('N.º de venda')[col].bfill()
                        # 3. FFILL: Preenche NaN/vazio para frente (pega o valor de cima)
                        df[col] = df.groupby('N.º de venda')[col].ffill()


                # DATA
                try:
                    month_map = { 'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04', 'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08', 'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12' }
                    dates = df['Data da venda'].astype(str).str.replace(r'\s\d{1,2}:\d{2}\s*hs\.$', '', regex=True).str.strip()
                    for pt_month, num_month in month_map.items():
                        dates = dates.str.replace(f' de {pt_month} de ', f'/{num_month}/', regex=False)
                    df['Data da venda'] = pd.to_datetime(dates, format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y')
                except Exception as e:
                    print(f"Aviso: Falha na formatação da coluna 'Data da venda'. Erro: {e}")

                # CONVERSÕES E CÁLCULOS BASE
                df['Unidades'] = pd.to_numeric(df['Unidades'], errors='coerce').fillna(0)
                
                # Cria uma chave arredondada para a Herança de Custos
                df['Price_Key'] = pd.to_numeric(df['Preço unitário de venda do anúncio (BRL)'], errors='coerce').fillna(0).round(2)
                df['Preço unitário de venda do anúncio (BRL)'] = df['Price_Key'] # Usa a chave arredondada
                
                original_freight = pd.to_numeric(df['Tarifas de envio (BRL)'], errors='coerce').fillna(0)
                original_tax_fee = pd.to_numeric(df['Tarifa de venda e impostos (BRL)'], errors='coerce').fillna(0)
                
                # --- NOVO: Herança de Custos (Passo 3.1) - Baseada APENAS em SKU e Preço Unitário ***
                if not cost_reference_df.empty:
                    # Prepara a referência para merge
                    reference = cost_reference_df[['SKU', 'Preço unitário de venda do anúncio (BRL)', 'Envio Seller', 'Tarifa de venda e impostos (BRL)']].copy()
                    
                    # Junta para encontrar custos previamente calculados (Chave: SKU e Price_Key)
                    df = df.merge(
                        reference,
                        left_on=['SKU', 'Price_Key'],
                        right_on=['SKU', 'Preço unitário de venda do anúncio (BRL)'],
                        how='left',
                        suffixes=('', '_ref')
                    )
                    
                    # *** ROBUSTEZ: Garantir que as colunas de referência existam após o merge ***
                    if 'Envio Seller_ref' not in df.columns:
                        df['Envio Seller_ref'] = np.nan
                    if 'Tarifa de venda e impostos (BRL)_ref' not in df.columns:
                        df['Tarifa de venda e impostos (BRL)_ref'] = np.nan
                        
                    # *** AJUSTE CRÍTICO NA HERANÇA ***
                    # Se o frete original for 0 OU o valor herdado for melhor que o atual (para evitar custos zerados em itens principais)
                    
                    # Condição para aplicar o custo de Envio Seller herdado:
                    # 1. Há um valor de referência (Envio Seller_ref não é NaN)
                    # 2. O frete original é zero (caso de item secundário de carrinho ou item sem custo)
                    condition_apply_freight_ref = pd.notna(df['Envio Seller_ref']) & (df['Tarifas de envio (BRL)'] == 0)

                    df['Tarifas de envio (BRL)'] = np.where(
                        condition_apply_freight_ref,
                        df['Envio Seller_ref'],
                        df['Tarifas de envio (BRL)']
                    )
                    
                    # Aplica o custo de Tarifa de Venda herdado, se existir
                    df['Tarifa de venda e impostos (BRL)'] = np.where(
                        pd.notna(df['Tarifa de venda e impostos (BRL)_ref']),
                        df['Tarifa de venda e impostos (BRL)_ref'],
                        df['Tarifa de venda e impostos (BRL)']
                    )
                    
                    # Remove colunas de merge
                    df.drop(columns=['Preço unitário de venda do anúncio (BRL)_ref', 'Envio Seller_ref', 'Tarifa de venda e impostos (BRL)_ref'], inplace=True, errors='ignore')
                
                df.drop(columns=['Price_Key'], inplace=True, errors='ignore')

                # Se houve herança, atualiza original_freight e original_tax_fee
                original_freight = df['Tarifas de envio (BRL)']
                original_tax_fee = df['Tarifa de venda e impostos (BRL)']
                # Fim da Herança de Custos
                
                # Continua o processamento normal a partir daqui, usando os valores ATUALIZADOS/HERDADOS
                
                # Conversão para Decimal para maior precisão
                original_freight_decimal = original_freight.astype(str).apply(Decimal)
                unidades_decimal = df['Unidades'].apply(Decimal)
                
                df['Tarifas de envio (BRL)'] = original_freight 
                df['Tarifa de venda e impostos (BRL)'] = original_tax_fee
                
                # LÓGICAS COMPLEXAS DE CUSTO (Somente se os valores não foram herdados e precisarem de cálculo)
                unit_price = df['Preço unitário de venda do anúncio (BRL)']
                is_flex = df['Forma de entrega'] == 'Mercado Envios Flex'
                
                # Condição de Custo Original (apenas quando o custo de envio não é zero E unidades > 0)
                is_original_cost = (original_freight.abs() > 0) & (df['Unidades'] > 0)
                
                # NOVO: Define um conjunto de envios com lógica de custo variável (FULL + Correios/Pontos)
                shipping_group_logic = ['Mercado Envios Full', 'Correios e pontos de envio', 'Pontos de Envio']
                is_full_logic_group = df['Forma de entrega'].isin(shipping_group_logic)
                
                # MINIMIZAÇÃO DE CUSTO (Custo MÁXIMO = Mais Vantajoso)
                is_valid_cost = (original_freight.abs() < 100)
                
                # *** CORREÇÃO CRÍTICA DO DIVISIONBYZERO (Revisada) ***
                valid_rows_for_division = is_original_cost & is_valid_cost & (df['Unidades'] > 0)
                df['Unit_Cost_Temp'] = np.nan 

                safe_division_result = (
                    original_freight_decimal[valid_rows_for_division] / unidades_decimal[valid_rows_for_division]
                ).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
                
                df.loc[valid_rows_for_division, 'Unit_Cost_Temp'] = safe_division_result.astype(float)
                # Fim da correção do DivisionByZero

                max_cost_group = df.groupby(['SKU', 'Forma de entrega'])['Unit_Cost_Temp'].transform('max')
                
                # Converto max_cost_group de volta para Decimal para a próxima multiplicação
                max_cost_group_decimal = max_cost_group.apply(lambda x: Decimal(str(x)) if pd.notna(x) else Decimal('0'))

                max_cost_total = (max_cost_group_decimal * unidades_decimal).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
                
                sale_counts = df.groupby('N.º de venda')['N.º de venda'].transform('count')
                
                # Um sub-item é parte de um carrinho (sale_counts > 1) E tem frete original zero.
                is_sub_item_to_calculate = (sale_counts > 1) & (original_freight == 0)

                # DEFINIÇÕES FLEX/FULL
                cost_flex_high_calc = Decimal("-9.11") 
                cond_flex_high = is_flex & (unit_price >= 79.00) 
                cost_flex_low_calc = Decimal("-1.10") 
                cond_flex_low = is_flex & (unit_price <= 78.99) 
                
                flex_calculated_cost_unit = np.select(
                    [cond_flex_high, cond_flex_low],
                    [cost_flex_high_calc, cost_flex_low_calc],
                    default=Decimal("0.00")
                ) 
                
                # CORREÇÃO CRÍTICA DO ZERAMENTO: Se o preço for <= 78.99 E pertencer ao grupo FULL/Logística, FORÇA ZERO.
                cond_full_force_zero = is_full_logic_group & (unit_price <= 78.99)
                
                # *** AJUSTE CRÍTICO DO FLEX: Aplica a lógica FLEX sempre que for Flex E não for um sub-item. ***
                cond_flex_apply_fixed = is_flex & (~is_sub_item_to_calculate)
                
                # *** AJUSTE: Minimização só se aplica se NÃO for um sub-item de carrinho E houver um custo melhor ***
                cond_apply_best_cost = (~is_sub_item_to_calculate) & (max_cost_total.astype(float) > original_freight) & (original_freight < 0)
                
                # Aplica o max_cost_total ao sub-item
                cond_full_apply_min = is_sub_item_to_calculate & is_full_logic_group

                # ATRIBUIÇÃO FINAL DE CUSTO (np.select)
                final_cost_decimal_array = np.select(
                    [
                        cond_flex_apply_fixed,      # 2. FLEX (Custo Fixo Unitário) - PRIO MÁXIMA PARA FLEX
                        cond_apply_best_cost,       # 3. CORRIGE O ITEM PRINCIPAL (Minimização)
                        cond_full_apply_min,        # 4. FULL/GRUPO (Custo Mínimo Total) - APLICA AO SUB-ITEM
                    ], 
                    [
                        flex_calculated_cost_unit,         # Resultado 2: Custo FLEX Unitário (Decimal)
                        max_cost_total,                    # Resultado 3: Aplica o Custo Mínimo (Decimal)
                        max_cost_total,                    # Resultado 4: Máximo Histórico Calculado (TOTAL: Decimal) - Aplica ao sub-item
                    ],
                    # Fallback: Mantém o valor original (Decimal)
                    default=original_freight_decimal 
                )
                
                # CORREÇÃO DEFINITIVA DO ZERAMENTO: APLICA-SE SE O ITEM ESTAVA NO GRUPO DE ZERAMENTO E O PREÇO ERA BAIXO
                final_cost_series = pd.Series(final_cost_decimal_array, index=df.index)
                
                final_cost_series = np.where(
                    cond_full_force_zero,
                    Decimal("0.00"),
                    final_cost_series
                )
                
                # 5. LÓGICA DE PRORRATEIO DE FRETE E TARIFA (CARRINHO) - SOBRESCREVE O CUSTO FINAL CALCULADO
                
                # Calcula a Receita por Produtos
                df['Receita por produtos (BRL)'] = (df['Unidades'] * df['Preço unitário de venda do anúncio (BRL)']).round(2)
                
                # Converte para Decimal para cálculos de proporção
                receita_decimal = df['Receita por produtos (BRL)'].apply(Decimal)
                
                # Cria colunas temporárias para os totais de custo (Frete e Tarifa de Venda)
                df['Frete_Calculado_Temp'] = pd.Series(final_cost_series).astype(float).values
                df['Tarifa_Calculada_Temp'] = df['Tarifa de venda e impostos (BRL)'].astype(float).values
                
                # *** AJUSTE DE ROBUSTEZ: Inicializa as colunas de merge para evitar KeyError ***
                df['total_receita_carrinho'] = np.nan
                df['total_frete_carrinho'] = np.nan
                df['total_tax_carrinho'] = np.nan 
                df['item_count_carrinho'] = np.nan
                
                # 5.1. Calcula o total de frete, tarifa e receita por N.º de venda (carrinho)
                agg_group = df.groupby('N.º de venda').agg(
                    total_receita=('Receita por produtos (BRL)', 'sum'),
                    total_frete=('Frete_Calculado_Temp', 'sum'),
                    total_tax=('Tarifa_Calculada_Temp', 'sum'),
                    item_count=('N.º de venda', 'count')
                )
                
                # Junta os totais de volta ao DataFrame.
                df = df.merge(agg_group, on='N.º de venda', suffixes=('', '_carrinho'), how='left')
                
                # Remove as colunas temporárias após a agregação
                df.drop(columns=['Frete_Calculado_Temp', 'Tarifa_Calculada_Temp'], inplace=True, errors='ignore')

                # Prorrateio (aplicado SÓ se for um carrinho E se houver custo/tarifa a ser distribuído)
                df['item_count_carrinho'] = df['item_count_carrinho'].fillna(0)
                df['total_frete_carrinho'] = df['total_frete_carrinho'].fillna(0)
                df['total_receita_carrinho'] = df['total_receita_carrinho'].fillna(0)
                df['total_tax_carrinho'] = df['total_tax_carrinho'].fillna(0)

                # Condição para aplicar o PRORRATEIO: Mais de um item no carrinho
                is_cart_sale = (df['item_count_carrinho'] > 1)
                
                if is_cart_sale.any():
                    # 5.2. Define o fator de prorrogação (Proporcional à Receita)
                    denominator = df['total_receita_carrinho'].apply(lambda x: x if x != 0 else 1)
                    prorate_factor = receita_decimal / denominator.apply(Decimal)
                    
                    # --- PRORRATEIO DE FRETE ---
                    total_frete_decimal = df['total_frete_carrinho'].apply(Decimal)

                    prorated_freight = (
                        prorate_factor * total_frete_decimal
                    ).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))

                    # --- PRORRATEIO DE TARIFA DE VENDA E IMPOSTOS ---
                    total_tax_decimal = df['total_tax_carrinho'].apply(Decimal)
                    
                    prorated_tax_fee = (
                        prorate_factor * total_tax_decimal
                    ).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
                    
                    # 5.4. Aplica os valores prorrateados nas linhas do carrinho
                    final_cost_series = np.where(
                        is_cart_sale,
                        prorated_freight,
                        final_cost_series # Mantém o custo complexo/minimizado original em vendas unitárias
                    )
                    
                    # Aplica a tarifa prorrateada na coluna final de Tarifa de Venda e Impostos
                    df['Tarifa de venda e impostos (BRL)'] = np.where(
                        is_cart_sale,
                        prorated_tax_fee.apply(lambda x: float(x)),
                        df['Tarifa de venda e impostos (BRL)']
                    )
                    
                df.drop(columns=['total_receita_carrinho', 'total_frete_carrinho', 'total_tax_carrinho', 'item_count_carrinho'], inplace=True, errors='ignore')
                
                # 6. CONVERSÃO E ARREDONDAMENTO FINAL (para precisão)
                final_cost_series = pd.Series(final_cost_series, index=df.index)

                def to_float_exact(val):
                    if pd.isna(val) or val is None: return 0.0
                    if isinstance(val, (float, np.float64)): val = Decimal(str(val))
                    if not isinstance(val, Decimal): return 0.0
                    
                    quantized = val.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    return float(quantized)

                df['Tarifas de envio (BRL)'] = final_cost_series.apply(to_float_exact)
                df.drop(columns=['Unit_Cost_Temp'], inplace=True, errors='ignore') 

                # --- REGRAS DE PREENCHIMENTO DO GABARITO ---
                
                df['Receita por produtos (BRL)'] = (df['Unidades'] * df['Preço unitário de venda do anúncio (BRL)']).round(2)
                
                df['SKU PRINCIPAL'] = df['SKU']
                
                df['Envio Seller'] = df['Tarifas de envio (BRL)']
                df['EMISSAO'] = df['Data da venda']
                
                df['origem'] = marketplace 
                df['conta'] = conta_final 

                df['TARIFA'] = 0 
                if 'Estado.1' in df.columns:
                    df['Estado'] = df['Estado.1']
                
                COLUNAS_GABARITO_FINAL = [
                    'SKU PRINCIPAL', 'SKU', 'Data da venda', 'EMISSAO', 'N.º de venda', 
                    'origem', '# de anúncio', 'tipo de anuncio', 'Venda por publicidade', 
                    'Forma de entrega', 'Preço unitário de venda do anúncio (BRL)', 
                    'Unidades', 'Receita por produtos (BRL)', 'Envio Seller', 'TARIFA', 
                    'Tarifa de venda e impostos (BRL)', 'conta', 'Estado'
                ]
                df = df[COLUNAS_GABARITO_FINAL]
                
                # 6.1. NOVO: Atualiza a Referência de Custos
                # Seleciona os custos finais calculados para este arquivo e os adiciona à referência.
                unique_costs = df.drop_duplicates(subset=['SKU', 'Preço unitário de venda do anúncio (BRL)'])[
                    ['SKU', 'Preço unitário de venda do anúncio (BRL)', 'Envio Seller', 'Tarifa de venda e impostos (BRL)']
                ].copy()

                cost_reference_df = pd.concat([cost_reference_df, unique_costs], ignore_index=True).drop_duplicates(
                    subset=['SKU', 'Preço unitário de venda do anúncio (BRL)'],
                    keep='first' # Mantém o primeiro custo encontrado para o par (SKU, Preço)
                )

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
                    content = str(cell.value)
                    if len(content) > max_length:
                        max_length = len(content)
                except:
                    pass
            
            adjusted_width = (max_length + 2) 
            if adjusted_width < 10: adjusted_width = 10 
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
        return f"<h1>Erro no Processamento</h1><p>Ocorreu um erro inesperado. Detalhes: <b>{e}</b></p><p>Verifique o formato dos arquivos e tente novamente.</p>", 500

if __name__ == '__main__':
    # Inicia o servidor Flask
    app.run(debug=True, host='0.0.0.0')
