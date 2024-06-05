import pandas as pd
from word_gen import example_contract




def singular_data_to_contract(df: pd.DataFrame, index_row:int):
    sub_df = df.iloc[index_row]
    date = sub_df['Fecha']
    rol = sub_df['Rol']
    address = sub_df['Residencia']
    rut = sub_df['RUT']
    full_name = sub_df['nombre_completo']
    nationality = sub_df['Nacionalidad']
    birth_date = sub_df['Fecha de nacimiento']
    profession = sub_df['Profesion']
    salary = sub_df['Sueldo']
    example_contract(date, rol, address, rut, full_name, nationality, birth_date, profession, str(salary))

    






