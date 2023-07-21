import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

pathWbDatos = './SKU-Competencia - 2.xlsx'

# Lee el archio excel y lo guarda un arreglo de objetos
def Get_Data_From_Excel(path):
    data = pd.read_excel(path)
    return data

# Muestra los datos del arreglo de objetos
def Show_Data(data):
    for i in range(len(data)):
        print(data['SKU'][i], data['COMPETIDOR'][i], data['RUTA'][i])

# Recorre los datos para obtener la información de la página web
def Get_Data_From_Page(driver, data, dfSalida):
    for i in range(len(data)):
        # mostrar datos de entrada
        print(data['SKU'][i], data['COMPETIDOR'][i], data['RUTA'][i])
        # recuperar datos de la pagina
        descripcion, precio_venta, precio_anterior, oferta_porcentaje = Get_Data_From_Web(driver, data['RUTA'][i], data['COMPETIDOR'][i])
        fila = {'SKU': data['SKU'][i], 'COMPETIDOR': data['COMPETIDOR'][i], 'DESCRIPCION': descripcion, 'PRECIO_VENTA': precio_venta, 'PRECIO_ANTERIOR': precio_anterior, 'OFERTA_PORCENTAJE': oferta_porcentaje}
        dfSalida.append(fila)

        # if (data['COMPETIDOR'][i] == 'WALMART'):
        #     # recuperar datos de la pagina
        #     descripcion, precio_venta, precio_anterior, oferta_porcentaje = Get_Data_From_Walmart(driver, data['RUTA'][i])
        #     # escribir datos en el dataframe
        #     fila = {'SKU': data['SKU'][i], 'COMPETIDOR': data['COMPETIDOR'][i], 'DESCRIPCION': descripcion, 'PRECIO_VENTA': precio_venta, 'PRECIO_ANTERIOR': precio_anterior, 'OFERTA_PORCENTAJE': oferta_porcentaje}
        #     dfSalida.append(fila)

        
# Inicia selenium y abre la pagina
def Init_Selenium():
    driver = webdriver.Chrome()
    return driver

# Cierra selenium
def Close_Selenium(driver):
    driver.close()

# Extraer datos de la página de Walmart
def Get_Data_From_Walmart(driver, path):
    
    # Navegar a la pagina
    driver.get(path)
    #driver.get('https://www.walmart.com.gt/dispensador-oster-de-mesa-para-agua-2grf-mod-os-wd522b/p')

    # Obtener el contenido de la pagina
    descripcion = driver.find_element(By.CLASS_NAME, 'vtex-store-components-3-x-productNameContainer')
    precio_venta = driver.find_element(By.CLASS_NAME, 'vtex-store-components-3-x-sellingPrice')

    try:
        precio_anterior = driver.find_element(By.CLASS_NAME, 'vtex-store-components-3-x-listPrice')
    except:
        precio_anterior = '0'

    try:
        oferta_porcentaje = driver.find_element(By.CLASS_NAME, 'vtex-product-price-1-x-savings')
    except:
        oferta_porcentaje = '0'

    retorno = ()
    try:
        retorno = (descripcion.text, precio_venta.text, precio_anterior.text, oferta_porcentaje.text)
    except:
        retorno = (descripcion.text, precio_venta.text, precio_anterior, oferta_porcentaje)
    return retorno

# Extraer datos de la página web según el competidor
def Get_Data_From_Web(driver, path, competidor):
    
    # Navegar a la pagina
    driver.get(path)
    
    # Pausar el tiempo de carga para facilitar la extracción de datos
    #time.sleep(30)

    # Obtener el nombre de las clases según el competidor
    clases = Get_Class_Name(competidor)

    # Obtener el contenido de la pagina
    descripcion = driver.find_element(By.CLASS_NAME, clases['descripcion'])
    
    precio_venta = driver.find_element(By.CLASS_NAME, clases['precio_venta'])

    try:
        precio_anterior = driver.find_element(By.CLASS_NAME, clases['precio_anterior'])
    except NoSuchElementException:
        precio_anterior = '0'

    try:
        oferta_porcentaje = driver.find_element(By.CLASS_NAME, clases['oferta_porcentaje'])
    except NoSuchElementException:
        oferta_porcentaje = '0'

    retorno = ()
    try:
        retorno = (descripcion.text, precio_venta.text, precio_anterior.text, oferta_porcentaje.text)
    except:
        retorno = (descripcion.text, precio_venta.text, precio_anterior, oferta_porcentaje)
    return retorno


# Retornar el nombre de las clases según el competidor
def Get_Class_Name(competidor):
    retorno = {'descripcion': '', 'precio_venta': '', 'precio_anterior': '', 'oferta_porcentaje': ''}

    if competidor == 'WALMART':
        retorno = {'descripcion': 'vtex-store-components-3-x-productNameContainer', 'precio_venta': 'vtex-store-components-3-x-sellingPrice', 'precio_anterior': 'vtex-store-components-3-x-listPrice', 'oferta_porcentaje': 'vtex-product-price-1-x-savings'}
    elif competidor == 'NOVEX':
        retorno = {'descripcion': 'productContainer__InfoTitle', 'precio_venta': 'details__Price', 'precio_anterior': 'disText', 'oferta_porcentaje': 'details__Before'}
    elif competidor == 'EPA':
        retorno = {'descripcion': 'page-title', 'precio_venta': 'price-wrapper', 'precio_anterior': 'old-price', 'oferta_porcentaje': 'aguacate'}
    
    return retorno



# Instrucciones generales del programa
def Main():
    # Crear data frame con las columnas que se van a utilizar
    dfScraping = []
    driver = Init_Selenium()
    data = Get_Data_From_Excel(pathWbDatos)
    print(data)
    Get_Data_From_Page(driver, data, dfScraping)
    # Crear el data frame final con los nombres de las columnas 
    dfFinal = pd.DataFrame(dfScraping, columns=['SKU', 'COMPETIDOR', 'DESCRIPCION', 'PRECIO_VENTA', 'PRECIO_ANTERIOR', 'OFERTA_PORCENTAJE'])
    # Crear el archivo excel de salida
    dfFinal.to_excel('./Salida.xlsx', index=False)

# Ejecutar el programa
if __name__ == "__main__":
    Main()