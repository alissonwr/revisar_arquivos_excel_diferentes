import pandas as pd
import unidecode
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def normalize_column_names(df):
    df.columns = [unidecode.unidecode(col.upper().strip()) for col in df.columns]
    return df

def find_similar_city(city, city_list, threshold=0.8):
    matches = get_close_matches(city, city_list, n=1, cutoff=threshold)
    return matches[0] if matches else None

def merge_excel_files(file1, file2, output_file):
    # Leitura dos arquivos
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    
    # Normalização dos nomes das colunas
    df1 = normalize_column_names(df1)
    df2 = normalize_column_names(df2)
    
    # Normalização dos nomes das cidades
    df1['MUNICIPIO'] = df1['MUNICIPIO'].apply(lambda x: unidecode.unidecode(x.upper().strip()))
    df2['MUNICIPIO'] = df2['MUNICIPIO'].apply(lambda x: unidecode.unidecode(x.upper().strip()))
    
    # Encontrar cidades semelhantes e comuns
    df1['MUNICIPIO_FILE2'] = df1['MUNICIPIO'].apply(lambda x: find_similar_city(x, df2['MUNICIPIO'].tolist()))
    df1 = df1.dropna(subset=['MUNICIPIO_FILE2'])
    
    df2_common = df2[df2['MUNICIPIO'].isin(df1['MUNICIPIO_FILE2'])]
    
    # Combinar os dataframes
    merged_df = pd.merge(df1, df2_common, left_on='MUNICIPIO_FILE2', right_on='MUNICIPIO', suffixes=('_FILE1', '_FILE2'))
    
    # Reordenar colunas para colocar as semelhantes lado a lado
    df1_columns = [col for col in merged_df.columns if '_FILE1' in col]
    df2_columns = [col for col in merged_df.columns if '_FILE2' in col]
    
    ordered_columns = []
    for col in df1_columns:
        col_base = col.replace('_FILE1', '')
        if f'{col_base}_FILE2' in df2_columns:
            ordered_columns.append(col)
            ordered_columns.append(f'{col_base}_FILE2')
    
    # Adicionar colunas que não têm correspondência
    remaining_columns = [col for col in merged_df.columns if col not in ordered_columns]
    ordered_columns.extend(remaining_columns)
    
    merged_df = merged_df[ordered_columns]
    
    # Salvar o resultado em um arquivo Excel
    merged_df.to_excel(output_file, index=False)
    
    # Aplicar a formatação condicional para células diferentes
    highlight_differences(output_file)
    
    return merged_df

def highlight_differences(output_file):
    # Carregar o arquivo Excel
    wb = load_workbook(output_file)
    ws = wb.active
    
    # Definir o preenchimento para as células com diferenças
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    # Obter o número de colunas e linhas
    col_count = ws.max_column
    row_count = ws.max_row
    
    # Iterar sobre as colunas e comparar os valores
    for col in range(1, col_count, 2):  # Considera pares de colunas
        col_name1 = ws.cell(row=1, column=col).value
        col_name2 = ws.cell(row=1, column=col + 1).value

        # Verificar se os nomes das colunas são idênticos (removendo os sufixos "_FILE1" e "_FILE2")
        if col_name1 and col_name2 and col_name1.split('_FILE')[0] == col_name2.split('_FILE')[0]:
            for row in range(2, row_count + 1):
                cell1 = ws.cell(row=row, column=col)
                cell2 = ws.cell(row=row, column=col + 1)
                
                # Comparar valores e aplicar formatação se diferentes
                if cell1.value != cell2.value:
                    cell1.fill = yellow_fill
                    cell2.fill = yellow_fill
    
    # Salvar o arquivo com as diferenças destacadas
    wb.save(output_file)

# Flask para a interface web
from flask import Flask, request, render_template, send_file

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file1 = request.files['file1']
        file2 = request.files['file2']
        output_file = "merged_output.xlsx"
        
        # Salvar os arquivos no servidor
        file1_path = f"./{file1.filename}"
        file2_path = f"./{file2.filename}"
        file1.save(file1_path)
        file2.save(file2_path)
        
        # Fazer a fusão dos arquivos
        merge_excel_files(file1_path, file2_path, output_file)
        
        # Retornar o arquivo Excel resultante para download
        return send_file(output_file, as_attachment=True)
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
