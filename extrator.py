import pandas as pd
import calendar

def compare_and_merge(file1_path, file2_path, output_path):
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)
    df2['EID'] = df2['EID'].astype(str)
    
    matched_rows = df2[df2['EID'].isin(df1['Email'])]
    
    result = pd.merge(
        matched_rows, df1, left_on='EID', right_on='Email', how='left', suffixes=('', '_old')
    )
    
    for col in result.columns:
        if col.endswith('_old'):
            base_col = col[:-4]
            if base_col in result.columns:
                result[base_col] = result[base_col].combine_first(result[col])
            result.drop(columns=[col], inplace=True)

    if 'Email' in result.columns:
        result.drop(columns=['Email'], inplace=True)

    result.to_excel(output_path, index=False)
    print(f"Arquivo resultante salvo em: {output_path}")

def tratar_accenture(file_path):
    df = pd.read_excel(file_path)

    # Renomear colunas para facilitar o acesso
    df = df.rename(columns={
        'Email': 'Email',
        'Profile -  Write a brief description about yourself, highlighting your main skills, experiences, and professional goals. ': 'Profile',
        'Languages [English]': 'Language_English',
        'Languages [Portuguese]': 'Language_Portuguese',
        'Languages [French]': 'Language_French',
        'Languages [German]': 'Language_German',
        'Picture – Please attach a professional and up-to-date photograph that will be used in your CV. The image should be of good quality, have a neutral background, and clearly show your face. ': 'Photo'
    })

    # Unir as colunas de idiomas em uma só coluna 'Language'
    def idioma_valido(val):
        return str(val).strip().lower() == 'coluna 1'

    idiomas = {
        'English': 'Language_English',
        'Portuguese': 'Language_Portuguese',
        'French': 'Language_French',
        'German': 'Language_German'
    }

    def get_idiomas(row):
        return ', '.join([lang for lang, col in idiomas.items() if idioma_valido(row.get(col, ''))])

    df['Language'] = df.apply(get_idiomas, axis=1)

    # Montar campo Education no formato desejado
    education_list = []
    for idx, row in df.iterrows():
        educations = []
        for i in range(1, 4):  # até 3 cursos

            # Busca flexível pelo nome das colunas (ignora espaços extras)
            def find_col(prefix, idx):
                for col in row.index:
                    if prefix in col and col.strip().endswith(f"- {idx}"):
                        return col
                return None

            course_col = find_col("Course Name", i)
            institution_col = find_col("Course Institution", i)
            end_col = find_col("Course Completion Date", i)

            course = row.get(course_col, '') if course_col else ''
            institution = row.get(institution_col, '') if institution_col else ''
            end = row.get(end_col, '') if end_col else ''

            def clean(val):
                if pd.isnull(val):
                    return ''
                val = str(val).strip()
                return '' if val.lower() in ['nan', 'nat', ''] else val

            course = clean(course)
            institution = clean(institution)
            end_clean = clean(end)
            end_year = ''
            if len(end_clean) >= 4 and end_clean[:4].isdigit():
                end_year = end_clean[:4]

            if course and institution:
                if end_year:
                    educations.append(f"{institution} - {course} - {end_year}")
                else:
                    educations.append(f"{institution} - {course}")
        education_list.append('; '.join(educations) if educations else '')
    df['Education'] = education_list

    # Projetos: pegar nomes, datas e descrições (até 5 projetos)
    project_names = []
    project_dates = []
    project_descs = []

    def format_project_date(start, end):
        def format_date(date_str):
            try:
                date = pd.to_datetime(date_str, errors='coerce')
                if pd.isnull(date):
                    return ''
                return f"{calendar.month_name[date.month]}/{date.year}"
            except Exception:
                return ''
        start_fmt = format_date(start)
        end_fmt = format_date(end)
        if start_fmt and end_fmt:
            return f"{start_fmt} - {end_fmt}"
        elif start_fmt:
            return start_fmt
        elif end_fmt:
            return end_fmt
        else:
            return ''

    for i in range(1, 6):
        name_col = f'Project {i} -  Name of the Project' if i == 1 else f'Project {i} -  Project Name '
        desc_col = f'Project {i} - Brief Project Description '
        start_col = f'Project {i} Start Date '
        end_col = f'Project {i} Completion Date '

        df[f'Project_{i}_Name'] = df.get(name_col, '')
        df[f'Project_{i}_Description'] = df.get(desc_col, '')
        # Use a função para formatar as datas
        df[f'Project_{i}_Date'] = df.apply(
            lambda row: format_project_date(row.get(start_col, ''), row.get(end_col, '')), axis=1
        )

        project_names.append(f'Project_{i}_Name')
        project_descs.append(f'Project_{i}_Description')
        project_dates.append(f'Project_{i}_Date')

    # Transformar em formato "longo" (um projeto por linha)
    records = []
    for _, row in df.iterrows():
        for i in range(5):
            if pd.notnull(row[project_names[i]]) and str(row[project_names[i]]).strip() != '':
                records.append({
                    'Email': row['Email'],
                    'Language': row['Language'],
                    'Profile': row['Profile'],
                    'Photo': row['Photo'],
                    'Education': row['Education'],
                    'Project Name': row[project_names[i]],
                    'Project Date': row[project_dates[i]],
                    'Description': row[project_descs[i]]
                })

    df_final = pd.DataFrame(records)
    return df_final

# Exemplo de uso:
file1 = "accenture.xlsx"  
file2 = "Skills.xlsx"  
output = "resultado.xlsx"  

accenture_tratado = tratar_accenture(file1)
accenture_tratado.to_excel("accenture_tratado.xlsx", index=False)
print("Arquivo accenture_tratado.xlsx gerado com sucesso.")

compare_and_merge("accenture_tratado.xlsx", file2, output)