import argparse
import pandas as pd

def parse_file_name(string):
    string = string.split('.')[0]
    string = string.replace('_', ' ')
    string = string.replace('sg', '') 
    return string.title()

def get_dfs(filenames):
    dfs = []
    for f in filenames:
        d = pd.read_excel(f, sheet_name=None)
        for sheet in d.values():
            sheet['Emphasis Area'] = parse_file_name(f)
            dfs.append(sheet)
    return dfs

def make_top_n_table(dfs, n, cols_to_keep, colnames):
    df = pd.concat(dfs)
    
    # Select and rename columns
    df = df.loc[:, cols_to_keep]
    df.columns = colnames

    # Sort by p value descending
    df = df.sort_values(by=['P value'], ascending=True)

    # Filter to top n
    df = df.head(n)

    # Format columns
    df[colnames[2]] = df[colnames[2]].apply('{:.2f}'.format)
    df[colnames[4]] = df[colnames[4]].apply('{:.2f}'.format)
    df[colnames[6]] = df[colnames[6]].apply('{:.2f}'.format)
    df[colnames[7]] = df[colnames[7]].apply('{:.3e}'.format)

    return df

def make_top_n_tables(n):
    tables = []
    names = []
    
    # Table B
    filenames = ["companion_animal.xlsx",
                  "equine.xlsx",
                  "food_animal.xlsx",
                  "special_species.xlsx"]
    dfs = get_dfs(filenames)
    cols = ['Subquestion',
            'Emphasis Area', 
            'SVM: avg', 
            'SVM: num responses', 
            'WVMA: avg', 
            'WVMA: num responses', 
            'Diff Mean (SVM-WVMA)', 
            'pval', 
            'sig']
    col_names = ['Item', 
            'Emphasis Area',
            'SVM mean', 
            'SVM responses', 
            'WVMA mean', 
            'WVMA responses', 
            'Mean difference', 
            'P value', 
            'sig']
    table = make_top_n_table(dfs, n, cols_to_keep=cols, colnames=col_names)
    tables.append(table)
    names.append('Table_B')

    # Table C
    filenames = ["summary_nontechnical_allspecies.xlsx"]
    dfs = get_dfs(filenames)
    cols = ['Subquestion',
            'Emphasis Area', 
            'SVM: avg', 
            'SVM: num responses', 
            'WVMA: avg', 
            'WVMA: num responses', 
            'Diff Mean (SVM-WVMA)', 
            'pval', 
            'sig']
    col_names = ['Item', 
            'Emphasis Area',
            'SVM mean', 
            'SVM responses', 
            'WVMA mean', 
            'WVMA responses', 
            'Mean difference', 
            'P value', 
            'sig']
    table = make_top_n_table(dfs, n, cols_to_keep=cols, colnames=col_names)
    table = table.drop(columns=['Emphasis Area'])
    tables.append(table)
    names.append('Table_C')

    # Table D
    filenames = ["companion_animal_sg.xlsx",
                  "equine_sg.xlsx",
                  "food_animal_sg.xlsx",
                  "special_species_sg.xlsx"]
    dfs = get_dfs(filenames)
    cols = ['Subquestion',
            'Emphasis Area', 
            'specialist: avg', 
            'specialist: num responses', 
            'generalist: avg', 
            'generalist: num responses', 
            'Diff Mean (specialist-generalist)', 
            'pval', 
            'sig']
    col_names = ['Item', 
            'Emphasis Area',
            'Specialist mean', 
            'Specialist responses', 
            'Generalist mean', 
            'Generalist responses', 
            'Mean difference', 
            'P value', 
            'sig']
    table = make_top_n_table(dfs, n, cols_to_keep=cols, colnames=col_names)
    tables.append(table)
    names.append('Table_D')

    # Table E
    filenames = ["summary_sg_nontechnical_allspecies.xlsx"]
    dfs = get_dfs(filenames)
    cols = ['Subquestion',
            'Emphasis Area', 
            'specialist: avg', 
            'specialist: num responses', 
            'generalist: avg', 
            'generalist: num responses',
            'Diff Mean (specialist-generalist)', 
            'pval', 
            'sig']
    col_names = ['Item', 
            'Emphasis Area',
            'Specialist mean', 
            'Specialist responses', 
            'Generalist mean', 
            'Generalist responses', 
            'Mean difference', 
            'P value', 
            'sig']
    table = make_top_n_table(dfs, n, cols_to_keep=cols, colnames=col_names)
    table = table.drop(columns=['Emphasis Area'])
    tables.append(table)
    names.append('Table_E')

    return tables, names

def main(n):
    tables, names = make_top_n_tables(n)

    writer = pd.ExcelWriter('top_n_tables.xlsx', engine='xlsxwriter')

    for i,table in enumerate(tables):
        sheet_name = names[i]
        table.to_excel(writer, sheet_name=sheet_name, index=False)

        # Auto-adjust columns widths
        for column in table:
            column_width = max(table[column].astype(str).map(len).max(), len(column))
            col_idx = table.columns.get_loc(column)
            writer.sheets[sheet_name].set_column(col_idx, col_idx, column_width)

    writer.save()

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--top', type=int)
    args = parser.parse_args()

    n = args.top

    main(n)
