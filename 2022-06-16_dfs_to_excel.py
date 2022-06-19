import pandas as pd
import numpy as np

"""
USEFUL LINKS:
https://xlsxwriter.readthedocs.io/worksheet.html#add_table
https://pandas-xlsxwriter-charts.readthedocs.io/chart_grouped_column.html#chart-grouped-column

"""


def main():
    df1 = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
    df2 = pd.DataFrame({"C": [1, 2, 3], "D": [4, 5, 6]})
    df3 = pd.DataFrame({"E": [1, 2, 3], "F": [4, 5, 6]})
    df4 = pd.DataFrame({"G": [1, 2, 3], "H": [4, 5, 6]})
    df5 = pd.DataFrame({"I": [1, 2, 3], "J": [4, 5, 6]})

    df1_1 = pd.DataFrame({"A_1": [1, 2, 3], "B_1": [4, 5, 6]})
    df1_2 = pd.DataFrame({"A_2": [1, 2, 3], "B_2": [4, 5, 6]})
    df1_3 = pd.DataFrame({"A_3": [1, 2, 3], "B_3": [4, 5, 6]})

    df2_1 = pd.DataFrame({"C_1": [1, 2, 3], "D_1": [4, 5, 6]})
    df2_2 = pd.DataFrame({"C_2": [1, 2, 3], "D_2": [4, 5, 6]})
    df2_3 = pd.DataFrame({"C_3": [1, 2, 3], "D_3": [4, 5, 6]})

    df3_1 = pd.DataFrame({"E_1": [1, 2, 3], "F_1": [4, 5, 6]})
    df3_2 = pd.DataFrame({"E_2": [1, 2, 3], "F_2": [4, 5, 6]})
    df3_3 = pd.DataFrame({"E_3": [1, 2, 3], "F_3": [4, 5, 6]})

    df4_1 = pd.DataFrame({"G_1": [1, 2, 3], "H_1": [4, 5, 6]})
    df4_2 = pd.DataFrame({"G_2": [1, 2, 3], "H_2": [4, 5, 6]})
    df4_3 = pd.DataFrame({"G_3": [1, 2, 3], "H_3": [4, 5, 6]})


    dfs = {"df1":df1, "df2":df2, "df3":df3, "df4":df4, "df5":df5, "df1_1": df1_1, 
            "df1_2": df1_2, "df1_3": df1_3, "df2_1": df2_1, "df2_2": df2_2, 
            "df2_3": df2_3, "df3_1": df3_1, "df3_2": df3_2, "df3_3": df3_3, 
            "df4_1": df4_1, "df4_2": df4_2, "df4_3": df4_3}

    fields = ["df1", "df2", "df3", "df4"]

    col = 5
    row = np.ones(len(fields))

    with pd.ExcelWriter('output.xlsx', engine="xlsxwriter") as writer:
        for df in dfs:
            for field in fields:
                if field in df:
                    dfs[df].to_excel(writer, sheet_name="Sheet_name_1", startrow=int(row[fields.index(field)]), startcol=int(fields.index(field)*col), header=False, index=False)
                    (max_row, max_col) = dfs[df].shape
                    column_settings = []
                    for header in dfs[df].columns:
                        column_settings.append({'header': header})
                    
                    worksheet = writer.sheets['Sheet_name_1']
                    worksheet.add_table(int(row[fields.index(field)]), int(fields.index(field)*col)+1, int(row[fields.index(field)]) + max_row-1 , int(fields.index(field)*col), {'columns': column_settings})
                    row[fields.index(field)] = row[fields.index(field)] + max_row + 2

    


main()
