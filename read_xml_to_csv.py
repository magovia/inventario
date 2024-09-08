# -*- coding: utf-8 -*-
"""
Created on Fri Sep  6 16:06:42 2024

@author: MAGOVIA
"""

import xmltodict
import pandas as pd
import tkinter as tk
from tkinter import filedialog

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
    
    return file_path

#%%

def parse_xml(file_path):
    try:
    
        with open(file_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()
            
        # Parsear el XML a un diccionario
        data_dict = xmltodict.parse(xml_content)
        
        
#%% Extraer ResumenFactura
        # Extract the list of line items
        resumenfc = data_dict['FacturaElectronica']['ResumenFactura']
        resumenfc = pd.DataFrame([resumenfc])
        #Columna no es necesaria
        resumenfc.drop(columns=['CodigoTipoMoneda'], inplace=True)
        
        # save DataFrame to a CSV file
        resumenfc.to_csv('tblresumenfactura.csv', index=False) 
    
    #%%  
        emisor = data_dict['FacturaElectronica']['Emisor']
        dataEmisor = {
        'Nombre': emisor['Nombre'],
        'personeriaJuridica': emisor['Identificacion']['Numero'],
        'NombreComercial': emisor['NombreComercial'],
        'Telefono': emisor['Telefono']['NumTelefono'],
        'CorreoElectronico': emisor['CorreoElectronico']
        }
        
        # Create DataFrame
        tblProveedor = pd.DataFrame([dataEmisor])
        # Optionally, save DataFrame to a CSV file
        tblProveedor.to_csv('proveedor.csv', index=False)  
     #%%  
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
    
        # Optionally, save DataFrame to a CSV file
        tblfactura.to_csv('factura.csv', index=False)
        
    #%%  
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
               lineaUnica.drop(columns=['CodigoComercial'], inplace=True)
               lineaUnica.drop(columns=['ImpuestoNeto'], inplace=True)
               lineaUnica.drop(columns=['MontoTotalLinea'], inplace=True)
               
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
                   lineaUnica['Descuento'] = 0  # You can also use None if preferred
    
               if 'Impuesto' in lineaUnica.columns:
                   # Unpack the 'Descuento' column into two new columns
                   impuesto_df = lineaUnica['Impuesto'].apply(pd.Series)
              
                   # Rename columns if needed (optional, to match keys)
                   impuesto_df.columns = ['CodImpuesto', 'ImpuestoCodTarifa','PorcentajeImpuesto','MontoImpuesto']
                
                   # Merge the unpacked columns into the original DataFrame
                   lineaUnica = pd.concat([lineaUnica, impuesto_df], axis=1)
                
                   # Drop the original 'Descuento' column (optional)
                   lineaUnica.drop(columns=['Impuesto'], inplace=True)
               else:
                   # Add a 'Descuento' column with null values (using pd.NA or None)
                   lineaUnica['CodImpuesto', 'ImpuestoCodTarifa','PorcentajeImpuesto','MontoImpuesto'] = pd.NA  # You can also use None if preferred
             
                # Optionally, save DataFrame to a CSV file
               lineaUnica.to_csv('linea_detalle_data.csv', index=False)
                   
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
                        'PorcentajeImpuesto': item.get('Impuesto', {}).get('Tarifa', ''),
                        'NaturalezaDescuento': item.get('Descuento', {}).get('NaturalezaDescuento', ''),
                        'CodImpuesto': item.get('Impuesto', {}).get('Codigo', ''),
                        'ImpuestoCodTarifa': item.get('Impuesto', {}).get('CodigoTarifa', ''),
                        'Cantidad': item.get('Cantidad', ''),
                        'PrecioUnitario': item.get('PrecioUnitario', ''),
                        'MontoDescuento': item.get('Descuento', {}).get('MontoDescuento', ''),
                        'MontoImpuesto': item.get('Impuesto', {}).get('Monto', ''),
                        'SubTotal': item.get('SubTotal', ''),
                        'MontoTotalLinea': item.get('MontoTotalLinea', ''),
                        'id_producto': code[:5] + str(len(item.get('Detalle', ''))),
                        'UnidadMedida': item.get('UnidadMedida',''),
                        'UnidadMedidaComercial': item.get('UnidadMedidaComercial',''),
                         }
                    detalle.append(record)
                # Convert list of dictionaries to DataFrame
                detalleLinea = pd.DataFrame(detalle)
                                        
                # Optionally, save DataFrame to a CSV file
                detalleLinea.to_csv('linea_detalle_data.csv', index=False)
            else:
                print(f"Element at index {idx} is of type {type(element)}: {element}")
         
        
    except xmltodict.expat.ExpatError as e:
        print(f"XML parsing error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    except Exception as e:
        print(f"Error processing XML: {str(e)}")
#%%
if __name__ == "__main__":
    # Open file dialog to select an XML file
    file_path = open_file_dialog()

    # If a file was selected, parse it
    if file_path:
        print(f"File selected: {file_path}")
        parse_xml(file_path)
    else:
        print("No file selected.")