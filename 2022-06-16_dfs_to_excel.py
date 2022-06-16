import pandas as pd

def main():
    df = pd.DataFrame({'num_legs': [2, 4, 8, 0],
                   'num_wings': [2, 0, 0, 0],
                   'num_specimen_seen': [10, 2, 1, 8]},
                  index=['falcon', 'dog', 'spider', 'fish'])
    df1 = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})


    print(df)
    print(df1)
    
    list_of_dfs = [df1, df]

    col = 0
    with pd.ExcelWriter('output.xlsx', engine="xlsxwriter") as writer:
        for df in list_of_dfs:
            df.to_excel(writer, sheet_name="Sheet_name_1", startrow=1, startcol=col, header=False, index=False)
            (max_row, max_col) = df.shape
            column_settings = []
            for header in df.columns:
                column_settings.append({'header': header})

            worksheet = writer.sheets['Sheet_name_1']
            worksheet.add_table(0, col, max_row, col + max_col-1, {'columns': column_settings})
            col += max_col + 2


main()