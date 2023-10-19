import pandas as pd

def Arreglar_MC(
        Archivo: str , 
        Salida:str , 
        Tolerancia: float = 0.1 ,
        ):

    df = pd.read_csv(Archivo , sep=';' , decimal=',')

    Lista_Comprobantes_C = [11 , 12 , 13 , 15 , 16 , 29 , 36 , 41 , 47 , 68 , 111 , 114 , 117 , 211 , 212 , 213]

    try:
        df['Fecha de Emisión'] = pd.to_datetime(df['Fecha de Emisión'], format='%Y-%m-%d')
    except:
        df['Fecha de Emisión'] = pd.to_datetime(df['Fecha de Emisión'], format='%d/%m/%Y')

    Lista = ['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'Otros Tributos' , 'IVA']
        
    # si la suma de todos los elemntos de la Lista es igual a la columna 'Total' o es menor a la tolerancia no hacer nada, sino la columa de 'Total' es la suma de todos los elementos de la lista
    df['Total Aux'] = df[Lista].sum(axis=1)
    df['Total Aux'] = df['Total Aux'].round(2)

    df['Diferecia'] = df['Imp. Total'] - df['Total Aux']
    df['Diferecia'] = df['Diferecia'].round(2)
    df['Diferecia'] = df['Diferecia'].abs()

    # si el 'tipo de comprobante' no esta en la Lista_Comprobantes_C y Si la diferencia es mayor a la tolerancia no hacer nada, sino la columa de 'Imp. Total' es la el valor de 'Total Aux'
    #df.loc[((df['Diferecia'] > Tolerancia) & (df['En lista']) == False), 'Imp. Total'] = df.loc[((df['Diferecia'] > Tolerancia) & (df['En lista']) == False), 'Total Aux']
    
    condition = (df['Diferecia'] > Tolerancia) & ~(df['tipo de comprobante'].isin(Lista_Comprobantes_C))
    df.loc[condition, 'Imp. Total'] = df.loc[condition, 'Total Aux']

    # Eliminar las columnas auxiliares
    #df.drop(columns=['Total Aux' , 'Diferecia'] , inplace=True)

    # Devolver la Fecha de Emisión en formato dd/mm/aaaa
    df['Fecha de Emisión'] = df['Fecha de Emisión'].dt.strftime('%d/%m/%Y')

    # Guardar el archivo
    df.to_csv(Salida , index=False , sep=';' , decimal=',')




def Arreglar_Comprobantes():

    Archivos = pd.read_excel('Lista de Archivos.xlsx')

    # Filtrar los que 'Procesar' == 'SI'
    Archivos = Archivos[Archivos['Procesar'] == 'SI']

    for Item in range(len(Archivos)):
        Archivo = Archivos['Archivo'].iloc[Item]
        Tolerancia = Archivos['Tolerancia'].iloc[Item]
        Archivo_Salida = Archivos['Archivo Salida'].iloc[Item]
        Tipo = Archivos['Tipo'].iloc[Item]

        # Arreglar el archivo
        if Tipo == 'MIS COMPROBANTES':
            Arreglar_MC(Archivo , Archivo_Salida , Tolerancia)




if __name__ == '__main__':
    Arreglar_Comprobantes()
