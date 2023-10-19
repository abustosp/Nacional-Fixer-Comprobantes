import pandas as pd
import numpy as np
import os

def Arreglar_MC(
        Archivo: str , 
        Salida:str , 
        Tolerancia: float,
        ):
    
    df = pd.read_csv(Archivo , sep=';' , decimal=',')

    # Lista de comprobantes C
    Lista_Comprobantes_C = [11 , 12 , 13 , 15 , 16 , 29 , 36 , 41 , 47 , 68 , 111 , 114 , 117 , 211 , 212 , 213]

    # Convertir la columna 'Fecha de Emisión' a formato dd/mm/aaaa
    try:
        df['Fecha de Emisión'] = pd.to_datetime(df['Fecha de Emisión'], format='%Y-%m-%d')
    except:
        df['Fecha de Emisión'] = pd.to_datetime(df['Fecha de Emisión'], format='%d/%m/%Y')

    # Reemplazalar con 0 los valores nulos de la lista
    Lista = ['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'Otros Tributos' , 'IVA']
    for i in Lista:
        df[i] = df[i].fillna(0)
        
    # si la suma de todos los elemntos de la Lista es igual a la columna 'Total' o es menor a la tolerancia no hacer nada, sino la columa de 'Total' es la suma de todos los elementos de la lista
    df['Total Aux'] = df[Lista].sum(axis=1).round(2)
    df['Diferecia'] = df['Imp. Total'] - df['Total Aux']
    df['Diferecia'] = df['Diferecia'].round(2).abs()

    condition = (df['Diferecia'] < Tolerancia) & ~(df['Tipo de Comprobante'].isin(Lista_Comprobantes_C))
    df.loc[condition, 'Imp. Total'] = df.loc[condition, 'Total Aux']

    # Eliminar las columnas auxiliares
    df.drop(columns=['Total Aux' , 'Diferecia'] , inplace=True)

    # Devolver la Fecha de Emisión en formato dd/mm/aaaa
    df['Fecha de Emisión'] = df['Fecha de Emisión'].dt.strftime('%d/%m/%Y')

    # Convertir las columnas a enteros
    Lista_Enteros = ['Tipo de Comprobante','Punto de Venta','Número Desde','Número Hasta','Tipo Doc. Receptor','Tipo Doc. Emisor','Cód. Autorización','Nro. Doc. Emisor','Nro. Doc. Receptor']

    for Item in Lista_Enteros:
        try:
            df[Item] = df[Item].fillna(np.nan).astype(np.int64)
        except:
            pass

    # Guardar el archivo
    df.to_csv(Salida , index=False , sep=';' , decimal=',')




def Arreglar_Comprobantes():

    Archivos = pd.read_excel('Lista de Archivos.xlsx')

    # Filtrar los que 'Procesar' == 'SI'
    Archivos = Archivos[Archivos['Procesar'] == 'SI']

    Archivos['Tolerancia'].fillna(0.05 , inplace=True)

    for Item in range(len(Archivos)):
        Archivo = Archivos['Archivo'].iloc[Item]
        Tolerancia = Archivos['Tolerancia'].iloc[Item]
        Archivo_Salida = Archivos['Archivo Salida'].iloc[Item]
        Tipo = Archivos['Tipo'].iloc[Item]

        # Obtener el directorio de la variable 'Archivo'
        Directorio = os.path.dirname(Archivo_Salida)
        os.makedirs(Directorio , exist_ok=True)

        # Arreglar el archivo
        if Tipo == 'MIS COMPROBANTES':
            Arreglar_MC(Archivo , Archivo_Salida , Tolerancia)
        else:
            print('No se reconoce el tipo de archivo')




if __name__ == '__main__':
    Arreglar_Comprobantes()
