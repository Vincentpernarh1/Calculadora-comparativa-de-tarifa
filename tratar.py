import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from itertools import groupby

def select_folder():
    """Abre uma caixa de diálogo para selecionar uma pasta e retorna seu caminho."""
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal do tkinter
    folder_path = filedialog.askdirectory(title="Selecione a pasta com os arquivos Excel")
    return folder_path

def process_and_pivot_data(file_path):
    """
    Lê um único arquivo Excel, transforma-o de formato largo para longo,
    e então o pivota para o formato largo final, conforme solicitado.
    Retorna o DataFrame pivotado.
    """
    try:
        # Prioriza a leitura da primeira planilha com cabeçalhos na 1ª linha (índice 0)
        df = pd.read_excel(file_path, header=0, sheet_name=0)
    except Exception:
        # Se isso falhar, tenta ler com cabeçalhos na 3ª linha (índice 2)
        try:
            df = pd.read_excel(file_path, header=2, sheet_name=0)
        except Exception as e2:
            raise Exception(f"Falha ao ler o arquivo. Verifique se o formato está correto. Erro: {e2}")

    # --- Identifica colunas pela posição (índice) ---
    veiculo_col_name = df.columns[0]
    motoristas_col_name = df.columns[1]
    
    df.rename(columns={
        veiculo_col_name: 'Veiculo',
        motoristas_col_name: 'Motoristas'
    }, inplace=True)
    
    df['Veiculo'] = df['Veiculo'].ffill()

    # --- Limpa a coluna 'Veiculo' ---
    df['Veiculo'] = df['Veiculo'].astype(str)
    text_to_remove_veiculo = ['- Um motorista', 'Um motorista', '-']
    for text in text_to_remove_veiculo:
        df['Veiculo'] = df['Veiculo'].str.replace(text, '', regex=False)
    df['Veiculo'] = df['Veiculo'].str.strip()


    # Identifica as colunas de ID e as colunas de rota para a operação de melt
    id_cols = ['Veiculo', 'Motoristas']
    route_cols = df.columns[2:].tolist()

    # Transforma o dataframe de colunas para linhas (melt)
    melted_df = pd.melt(
        df,
        id_vars=id_cols,
        value_vars=route_cols,
        var_name='RotaCompleta',
        value_name='Tarifa'
    )

    # Limpa os dados
    melted_df.dropna(subset=['Tarifa'], inplace=True)
    melted_df['Tarifa'] = pd.to_numeric(melted_df['Tarifa'], errors='coerce')
    melted_df.dropna(subset=['Tarifa'], inplace=True)
    melted_df = melted_df[melted_df['Tarifa'] > 0]

    # Limpa a string da rota antes de tentar dividi-la

    print("Antes da limpeza, RotaCompleta:", melted_df['RotaCompleta'].unique())
    melted_df['RotaCompleta'] = melted_df['RotaCompleta'].astype(str).str.replace('até', '-', regex=False)

    # Divide a coluna 'RotaCompleta' em 'Origem' e 'Destino'
    try:
        split_rota = melted_df['RotaCompleta'].str.split(' X ', n=1, expand=True)
        
        melted_df['Origem'] = split_rota[0]
        # Apenas adiciona a coluna de destino se a divisão foi bem-sucedida
        if split_rota.shape[1] > 1:
            melted_df['Destino'] = split_rota[1]
        else:
            melted_df['Destino'] = None # Preenche com nulo se não houver separador

    except Exception as e:
        raise Exception(f"Ocorreu um erro inesperado ao dividir a coluna de rotas. Erro: {e}")

    # Agora, remove todas as linhas onde a coluna 'Destino' não pôde ser criada
    melted_df.dropna(subset=['Destino'], inplace=True)

    # Adiciona uma verificação para garantir que há dados para processar
    if melted_df.empty:
        # Se não houver rotas válidas, retorna um DataFrame vazio para evitar erros
        return pd.DataFrame()
            
    melted_df['Origem'] = melted_df['Origem'].str.strip()
    melted_df['Destino'] = melted_df['Destino'].str.strip()

    # --- Limpa a coluna 'Destino' ---
    text_to_remove_dest = ['(R$/KM) vice-versa', '(R$/KM)', 'R$/KM', 'vice-versa']
    for text in text_to_remove_dest:
        melted_df['Destino'] = melted_df['Destino'].str.replace(text, '', regex=False)
    melted_df['Destino'] = melted_df['Destino'].str.strip()


    # Pivota a tabela longa para obter o formato largo final desejado
    pivoted_df = melted_df.pivot_table(
        index=['Origem', 'Destino'],
        columns=['Veiculo', 'Motoristas'],
        values='Tarifa'
    )
    
    # Reseta o índice para tornar 'Origem' e 'Destino' colunas regulares
    pivoted_df.reset_index(inplace=True)

    return pivoted_df

def save_with_merged_headers(pivoted_df, output_path):
    """
    Salva o DataFrame pivotado em um arquivo Excel com cabeçalhos mesclados
    personalizados usando o motor XlsxWriter.
    """
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    workbook = writer.book
    # Adiciona manualmente a planilha para evitar o KeyError 'Sheet1'
    worksheet = workbook.add_worksheet()

    # Define formatos de célula
    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'top',
        'fg_color': '#D7E4BC', 'border': 1, 'align': 'center'
    })
    cell_format = workbook.add_format({'border': 1})
    
    # --- Escreve manualmente os cabeçalhos complexos e mesclados ---
    worksheet.write('A3', 'Origem', header_format)
    worksheet.write('B3', 'Destino', header_format)
    
    col_offset = 2 
    headers = pivoted_df.columns[col_offset:]
    current_col = col_offset
    
    for veiculo, group in groupby(headers, key=lambda x: x[0]):
        group_list = list(group)
        num_motoristas = len(group_list)
        
        if num_motoristas > 1:
            worksheet.merge_range(0, current_col, 0, current_col + num_motoristas - 1, veiculo, header_format)
        else:
            worksheet.write(0, current_col, veiculo, header_format)
        
        worksheet.merge_range(1, current_col, 1, current_col + num_motoristas - 1, 'Motorista(s)', header_format)

        for i, header_tuple in enumerate(group_list):
            motorista_value = header_tuple[1]
            worksheet.write(2, current_col + i, motorista_value, header_format)
        
        current_col += num_motoristas

    # --- Escreve manualmente os dados linha por linha para evitar o erro ---
    start_row = 3
    for row_num, row_data in pivoted_df.iterrows():
        row_values = row_data.values.tolist()
        for col_num, cell_value in enumerate(row_values):
            if pd.isna(cell_value):
                cell_value = None
            worksheet.write(start_row + row_num, col_num, cell_value, cell_format)

    # Fecha o writer para salvar o arquivo
    writer.close()


def main():
    """
    Função principal para orquestrar a seleção de pastas, processamento de arquivos
    e salvamento dos resultados arquivo por arquivo com formatação personalizada.
    """
    try:
        folder = select_folder()
        if not folder:
            print("Nenhuma pasta foi selecionada. O programa será encerrado.")
            messagebox.showinfo("Encerrado", "Nenhuma pasta foi selecionada.")
            return

        files_in_folder = os.listdir(folder)
        excel_files = [f for f in files_in_folder if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]

        if not excel_files:
            messagebox.showwarning("Aviso", "Nenhum arquivo Excel encontrado na pasta selecionada.")
            return

        print(f"Pasta selecionada: {folder}\n" + "-" * 30)
        processed_count, error_count = 0, 0

        for filename in excel_files:
            if '_trans' in filename:
                continue

            file_path = os.path.join(folder, filename)
            try:
                print(f"Processando arquivo: {filename}...")
                pivoted_df = process_and_pivot_data(file_path)
                
                # Se o DataFrame estiver vazio após o processamento, pula para o próximo arquivo
                if pivoted_df.empty:
                    print(f"-> Nenhuma rota válida encontrada em {filename}. Arquivo ignorado.")
                    continue

                base_name, extension = os.path.splitext(filename)
                output_filename = f"{base_name}_trans{extension}"
                output_path = os.path.join(folder, output_filename)
                
                save_with_merged_headers(pivoted_df, output_path)
                
                print(f"-> Arquivo transformado salvo como: {output_filename}")
                processed_count += 1

            except Exception as e:
                error_count += 1
                print(f"-> ERRO ao processar o arquivo {filename}: {e}")
                messagebox.showerror("Erro de Processamento", f"Ocorreu um erro ao processar o arquivo:\n\n{filename}\n\nErro: {e}\n\nEste arquivo será ignorado.")

        print("-" * 30)
        if processed_count > 0:
            success_message = f"{processed_count} arquivo(s) foram transformados com sucesso!"
            if error_count > 0:
                success_message += f"\n{error_count} arquivo(s) falharam."
            messagebox.showinfo("Concluído", success_message)
        else:
            messagebox.showerror("Nenhum Resultado", "Nenhum arquivo foi processado com sucesso.")

    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
        messagebox.showerror("Erro Inesperado", f"Ocorreu um erro inesperado no programa:\n\n{e}")

if __name__ == "__main__":
    main()
