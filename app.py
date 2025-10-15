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
    
    try:
        # 2. PROCESSA TODOS OS GRUPOS DE UPLOAD ENVIADOS PELO JAVASCRIPT
        # O loop continua enquanto houver a chave 'marketplace_X' no formulário.
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
                
                df = df.rename(columns={'Tipo de anúncio': 'tipo de anuncio'}, inplace=False) 
                df['Forma de entrega'] = df['Forma de entrega'].astype(str).str.strip() 

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
                df['Preço unitário de venda do anúncio (BRL)'] = pd.to_numeric(
                    df['Preço unitário de venda do anúncio (BRL)'], errors='coerce'
                ).fillna(0)
                
                original_freight = pd.to_numeric(df['Tarifas de envio (BRL)'], errors='coerce').fillna(0)
                original_freight_decimal = original_freight.astype(str).apply(Decimal)
                unidades_decimal = df['Unidades'].apply(Decimal)
                
                df['Tarifas de envio (BRL)'] = original_freight 
                df['Tarifa de venda e impostos (BRL)'] = pd.to_numeric(df['Tarifa de venda e impostos (BRL)'], errors='coerce').fillna(0)
                
                # LÓGICAS COMPLEXAS DE CUSTO
                unit_price = df['Preço unitário de venda do anúncio (BRL)']
                is_flex = df['Forma de entrega'] == 'Mercado Envios Flex'
                is_original_cost = (original_freight.abs() > 0) & (df['Unidades'] > 0)
                
                # NOVO: Define um conjunto de envios com lógica de custo variável (FULL + Correios/Pontos)
                shipping_group_logic = ['Mercado Envios Full', 'Correios e pontos de envio', 'Pontos de Envio'] # Corrigido o nome da string
                is_full_logic_group = df['Forma de entrega'].isin(shipping_group_logic)
                
                # MINIMIZAÇÃO DE CUSTO (Custo MÁXIMO = Mais Vantajoso)
                is_valid_cost = (original_freight.abs() < 100)
                df['Unit_Cost_Temp'] = np.where(
                    is_original_cost & is_valid_cost,
                    (original_freight_decimal / unidades_decimal).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)),
                    np.nan 
                )

                max_cost_group = df.groupby(['SKU', 'Forma de entrega'])['Unit_Cost_Temp'].transform('max')
                max_cost_total = (max_cost_group.apply(Decimal) * unidades_decimal).apply(lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
                
                sale_counts = df.groupby('N.º de venda')['N.º de venda'].transform('count')
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
                
                # CORREÇÃO CRÍTICA DO ZERAMENTO: Se o preço for <= 78.99 E pertencer ao grupo FULL/Logística, FORÇA ZERO,
                # independentemente do custo original, pois é um custo que deve ser zerado por regra.
                cond_full_force_zero = is_full_logic_group & (unit_price <= 78.99)
                
                cond_flex_apply_fixed = is_flex & (original_freight == 0)
                cond_apply_best_cost = (max_cost_total.astype(float) > original_freight) & (original_freight < 0)
                cond_full_apply_min = is_sub_item_to_calculate & is_full_logic_group

                # ATRIBUIÇÃO FINAL DE CUSTO (np.select)
                final_cost_decimal_array = np.select(
                    [
                        cond_flex_apply_fixed,      # 2. FLEX (Custo Fixo Unitário) - Se custo original for 0
                        cond_apply_best_cost,       # 3. CORRIGE O ITEM PRINCIPAL (Minimização) - Se o custo histórico for melhor
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
                # Esta sobrescrita tem PRIORIDADE MÁXIMA sobre o np.select, resolvendo o problema de fallback.
                final_cost_series = pd.Series(final_cost_decimal_array, index=df.index)
                
                final_cost_series = np.where(
                    cond_full_force_zero,
                    Decimal("0.00"),
                    final_cost_series
                )
                
                # 6. CONVERSÃO E ARREDONDAMENTO FINAL (para precisão)
                final_cost_series = pd.Series(final_cost_series, index=df.index)

                def to_float_exact(val):
                    if pd.isna(val) or val is None: return 0.0
                    if isinstance(val, float): val = Decimal(str(val))
                    quantized = val.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    return float(quantized)

                df['Tarifas de envio (BRL)'] = final_cost_series.apply(to_float_exact)
                df.drop(columns=['Unit_Cost_Temp'], inplace=True) 

                # --- REGRAS DE PREENCHIMENTO DO GABARITO ---
                
                df['Receita por produtos (BRL)'] = (df['Unidades'] * df['Preço unitário de venda do anúncio (BRL)']).round(2)
                df['SKU PRINCIPAL'] = df['SKU']
                df['Envio Seller'] = df['Tarifas de envio (BRL)']
                df['EMISSAO'] = df['Data da venda']
                
                df['origem'] = marketplace # Marketplace selecionado
                df['conta'] = conta_final # Coluna 'conta' preenchida com a lógica do MAPA_CONTA

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
