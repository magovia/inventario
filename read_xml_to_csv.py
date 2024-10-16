# -*- coding: utf-8 -*-
"""
Created on Fri Sep  6 16:06:42 2024

@author: MAGOVIA
"""

import xmltodict
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import ctypes
import pyodbc

#%%
def open_file_dialog():
    # Create a Tkinter root window (hidden)
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    
    # Open file dialog to choose XML file
    file_path = filedialog.askopenfilename(
        title="Seleccione archivo XML", 
        filetypes=[("XML Files", "*.xml")]
    )
    
    print(f"File selected: {file_path}")
    
    return file_path

#%%  Function to parse xml file and update database    
def parse_xml(file_path):
    try:
    
        with open(file_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()
            
        # Parsear el XML a un diccionario
        data_dict = xmltodict.parse(xml_content)

#%% Extraer ResumenFactura
         # Extract the list of line items
        resumenfc = data_dict['FacturaElectronica']['ResumenFactura']
        
        resumenfact = {
            'TotalServGravados':resumenfc.get('TotalServGravados',0),
            'TotalServExentos':resumenfc.get('TotalServExentos',0),
            'TotalServExonerado':resumenfc.get('TotalServExonerado',0),
            'TotalMercanciasGravadas':resumenfc.get('TotalMercanciasGravadas',0),
            'TotalMercanciasExentas':resumenfc.get('TotalMercanciasExentas',0),
            'TotalMercExonerada':resumenfc.get('TotalMercExonerada',0),
            'TotalGravado':resumenfc.get('TotalGravado',0),
            'TotalExento':resumenfc.get('TotalExento',0),
            'TotalExonerado':resumenfc.get('TotalExonerado',0),
            'TotalVenta':resumenfc.get('TotalVenta',0),
            'TotalVentaNeta':resumenfc.get('TotalVentaNeta',0),
            'TotalImpuesto':resumenfc.get('TotalImpuesto',0),
            'TotalIVADevuelto':resumenfc.get('TotalIVADevuelto',0), 
            'TotalComprobante':resumenfc.get('TotalComprobante',0),
            'TotalOtrosCargos':resumenfc.get('TotalOtrosCargos',0),
            'TotalDescuentos':resumenfc.get('TotalDescuentos',0)
            }
        
        resumenfact = pd.DataFrame([resumenfact])

        # Create the full path for the Excel file
        excel_path = os.path.join(script_dir, 'tblresumenfactura.xlsx')
        
        # Save DataFrame as xlsx file in the same directory as the script
        resumenfact.to_excel(excel_path, index=False)
        
        # save DataFrame as xlsx file 
        #resumenfact.to_excel('tblresumenfactura.xlsx', index=False)   
                

#%% Extrae el Proveedor 
        emisor = data_dict['FacturaElectronica']['Emisor']
        dataEmisor = {
        'Nombre': emisor.get('Nombre',"No encontrado"),
        'personeriaJuridica': emisor.get('Identificacion',{}).get('Numero',"0000"),
        'NombreComercial': emisor.get('NombreComercial',"No encontrado"),
        'Telefono': emisor.get('Telefono',{}).get('NumTelefono',"00000"),
        'CorreoElectronico': emisor.get('CorreoElectronico',"No encontrado")
        }
        
        # Create DataFrame
        tblProveedor = pd.DataFrame([dataEmisor])
               
        # Create the full path for the Excel file
        excel_path = os.path.join(script_dir, 'proveedor.xlsx')
        
        # Save DataFrame as xlsx file in the same directory as the script
        tblProveedor.to_excel(excel_path, index=False)
       
        # Optionally, save DataFrame as xlsx file
        #tblProveedor.to_excel('proveedor.xlsx', index=False) 
        
        
#%% Extrae la factura  
        tblfactura = data_dict['FacturaElectronica']
        # type(factura)
        # print('total items en factura: ',len(factura['DetalleServicio']['LineaDetalle']))
        
        fc = {
        'FacturaID': tblfactura.get('NumeroConsecutivo',''),
        'FechaEmision': tblfactura.get('FechaEmision',''),
        'personeriaJuridica': emisor['Identificacion']['Numero'],
        'CondicionVenta': tblfactura.get('CondicionVenta',''),
        'PlazoCredito': tblfactura.get('PlazoCredito',''),
        'MedioPago': tblfactura.get('MedioPago','')      
        }
        
        # Create DataFrame
        tblfactura = pd.DataFrame([fc])
    
        # Create the full path for the Excel file
        excel_path = os.path.join(script_dir, 'factura.xlsx')
        
        # Save DataFrame as xlsx file in the same directory as the script
        tblfactura.to_excel(excel_path, index=False)
        
        #tblfactura.to_csv('factura.csv', index=False)
       # tblfactura.to_excel('factura.xlsx', index=False)

        
#%% Extrae el detalle de las lineas
        # Extract the list of line items
        line_items = data_dict['FacturaElectronica']['DetalleServicio']['LineaDetalle']
        
     
        # Iterate over the elements and check their types
        for idx, element in enumerate(line_items):
            if isinstance(element, int):
                print(f"Element at index {idx} is an integer: {element}")
            
            elif isinstance(element, str):
               # print(f"Element at index {idx} is a string: {element}")
               
               lineaUnica = pd.DataFrame.from_dict(data_dict['FacturaElectronica']['DetalleServicio'], orient='index')
               
               #Extrae el codigo
               codigoCabys = line_items.get('Codigo')
               #Crea una nueva columna
               lineaUnica['id_producto']=codigoCabys[:5] + str(len(line_items.get('Detalle', '')))
               
               #Agregar Factura ID
               lineaUnica['FacturaID']=data_dict['FacturaElectronica']['NumeroConsecutivo']
               
               
               #Eliminar columnas
               #lineaUnica.drop(columns=['MontoTotalLinea'], inplace=True)
               lineaUnica.drop(columns=['CodigoComercial'], inplace=True)
               
               if 'ImpuestoNeto' in lineaUnica.columns:
                   lineaUnica.drop(columns=['ImpuestoNeto'], inplace=True)
               
               
               if 'UnidadMedidaComercial' not in lineaUnica.columns:
                   lineaUnica['UnidadMedidaComercial']=lineaUnica['UnidadMedida']
               
               if 'Descuento' in lineaUnica.columns:
                   # Unpack the 'Descuento' column into two new columns
                   descuento_df = lineaUnica['Descuento'].apply(pd.Series)
              
                   # Rename columns if needed (optional, to match keys)
                   descuento_df.columns = ['MontoDescuento', 'NaturalezaDescuento']
                
                   # Merge the unpacked columns into the original DataFrame
                   lineaUnica = pd.concat([lineaUnica, descuento_df], axis=1)
                
                   # Drop the original 'Descuento' column (optional)
                   lineaUnica.drop(columns=['Descuento'], inplace=True)
               else:
                   # Add a 'Descuento' column with null values (using pd.NA or None)
                   lineaUnica['MontoDescuento'] = 0  # You can also use None if preferred
    
               if 'Impuesto' in lineaUnica.columns:
                   # Unpack the 'Descuento' column into two new columns
                   impuesto_df = lineaUnica['Impuesto'].apply(pd.Series)
              
                   # Rename columns if needed (optional, to match keys)
                   impuesto_df.columns = ['CodImpuesto', 'ImpuestoCodTarifa','ImpuestoPorcentaje','MontoImpuesto']
                
                   # Merge the unpacked columns into the original DataFrame
                   lineaUnica = pd.concat([lineaUnica, impuesto_df], axis=1)
                
                   # Drop the original 'Descuento' column (optional)
                   lineaUnica.drop(columns=['Impuesto'], inplace=True)
               else:
                   # Add a 'Descuento' column with null values (using pd.NA or None)
                   lineaUnica['CodImpuesto']=0
                   lineaUnica['ImpuestoCodTarifa']=0
                   lineaUnica['ImpuestoPorcentaje']=0
                   lineaUnica['MontoImpuesto'] = 0  
             
               
             # Create the full path for the Excel file
               excel_path = os.path.join(script_dir, 'linea_detalle_data.xlsx')
        
            # Save DataFrame as xlsx file in the same directory as the script
               excel_path.to_excel(excel_path, index=False)
                
             # Optionally, save DataFrame to a CSV file
               #lineaUnica.to_csv('linea_detalle_data.csv', index=False)
               #lineaUnica.to_excel('linea_detalle_data.xlsx', index=False)
               
               #delete_all_from_table(conn, "XLS_Detalle")
               #append_to_access(conn, lineaUnica, "XLS_Detalle") 
                   
            elif isinstance(element, dict):
                # print(f"Element at index {idx} is a dictionary: {element}")           
                # Create a list to store extracted data
                detalle = []
                # Extract values for each item
                for item in line_items:
                    code = item.get('Codigo')
                    # detalle = item.get('Detalle', '')
    
                    record = {
    
                        'NumeroLinea': item.get('NumeroLinea', ''),
                        'FacturaID': fc['FacturaID'],
                        'Codigo': item.get('Codigo', ''),
                        'Detalle':item.get('Detalle', ''),
                        'ImpuestoPorcentaje': item.get('Impuesto', {}).get('Tarifa', ''),
                        'NaturalezaDescuento': item.get('Descuento', {}).get('NaturalezaDescuento', ''),
                        'CodImpuesto': item.get('Impuesto', {}).get('Codigo', ''),
                        'ImpuestoCodTarifa': item.get('Impuesto', {}).get('CodigoTarifa', ''),
                        'Cantidad': item.get('Cantidad', ''),
                        'PrecioUnitario': item.get('PrecioUnitario', ''),
                        'MontoDescuento': item.get('Descuento', {}).get('MontoDescuento', 0),
                        'MontoImpuesto': item.get('Impuesto', {}).get('Monto', 0),
                        'SubTotal': item.get('SubTotal', ''),
                        'MontoTotalLinea': item.get('MontoTotalLinea', ''),
                        'id_producto': code[:5] + str(len(item.get('Detalle', ''))),
                        'UnidadMedida': item.get('UnidadMedida',''),
                        'UnidadMedidaComercial': item.get('UnidadMedidaComercial',''),
                        'MontoTotal': item.get('MontoTotal','')
                         }
                    detalle.append(record)
                # Convert list of dictionaries to DataFrame
                detalleLinea = pd.DataFrame(detalle)

                # Create the full path for the Excel file
                excel_path = os.path.join(script_dir, 'linea_detalle_data.xlsx')
        
                # Save DataFrame as xlsx file in the same directory as the script
                detalleLinea.to_excel(excel_path, index=False)
                                        
                # Optionally, save DataFrame to a CSV file
                #detalleLinea.to_excel('linea_detalle_data.xlsx', index=False)
                
            else:
                print(f"Element at index {idx} is of type {type(element)}: {element}")
    
        Mbox('Importar XML', f'Archivo xml importado con éxito.\n\nFactura: {fc["FacturaID"]}\nProveedor: {emisor.get("Nombre", "No encontrado")}', 64)

        
    except xmltodict.expat.ExpatError as e:
        print(f"XML parsing error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        Mbox('Importar XML', f"No se pudo subir la factura electronica:      {e}", 16)
    except Exception as e:
        print(f"Error processing XML: {str(e)}")

#%% funcion para el mensaje box

def Mbox(title, text, style):
    
    ## https://stackoverflow.com/questions/2963263/how-can-i-create-a-simple-message-box-in-python 
    ##  Styles:
    ##  0 : OK
    ##  1 : OK | Cancel
    ##  2 : Abort | Retry | Ignore
    ##  3 : Yes | No | Cancel
    ##  4 : Yes | No
    ##  5 : Retry | Cancel 
    ##  6 : Cancel | Try Again | Continue


    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

        
#%%
if __name__ == "__main__":
    
    
    #Get the directory of the current script
    script_dir = os.path.dirname(os.path.abspath(__file__))
        
    # Open file dialog to select an XML file
    file_path = open_file_dialog()

    # If a file was selected, parse it
    if file_path:
        parse_xml(file_path)
        #Mbox('Importar XML', 'Archivo xml importado con éxito', 64)
    else:
        print("No file selected.")
        Mbox('Importar XML', 'Hubo un error al importar el archivo xml', 16)