import pandas as pd
import time
import math
import pyautogui
import os, sys
import json
import pyodbc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException 
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../src')))
from src.Fuji.conexion import Conexion

class Main:

    def __init__(self,driver=None):
        self.driver = driver

    def guardar_en_json(self,data):
        ruta_json = r"C:\Users\aprendiz.serviciosti\Desktop\json_data.json"
        # Guardar los datos de número de resultados y cantidad de páginas en el archivo JSON
        with open(ruta_json, 'w') as file:
            json.dump(data, file, indent=4)
        print("Datos guardados correctamente en el archivo json_data.json")

    def leer_archivoJSON(self):
        ruta_json = r"C:\Users\aprendiz.serviciosti\Desktop\json_data.json"
        try:
            with open(ruta_json, 'r') as json_data:
                data_cargada = json.load(json_data)
                print("Datos cargados desde el archivo JSON:")
                print(data_cargada)
                return data_cargada
        except FileNotFoundError:
            print(f"Archivo JSON no encontrado en {ruta_json}, creando uno nuevo.")
            return {}

    def Iniciar_navegacion(self):
        self.driver = webdriver.Firefox()
        self.driver.maximize_window()
        self.driver.get("https://amazon.com")  
        return self.driver

    def Cerrar_navegacion(self):
        if self.driver:
            self.driver.quit()

    def Buscar_objeto_barra_de_busqueda(self):
        buscar = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="twotabsearchtextbox"]')))
        buscar.click()

        buscar.send_keys("Iphone 15")
        Etiqueta_buscar = '//*[@id="nav-search-submit-button"]'
        buscar = self.driver.find_element(By.XPATH, Etiqueta_buscar)
        buscar.click()

    def Obtener_numero_de_paginas_y_resultados(self):
            h2_element_numero_resultados = self.driver.find_element(By.XPATH, "//h2[contains(@class, 'a-size-base') and contains(@class, 'a-spacing-small')]").text

            a_element_cantidad_paginas = self.driver.find_element(By.XPATH,"//span[@class='s-pagination-item s-pagination-disabled']").text

            numero_resultados = h2_element_numero_resultados.split("of")[1].split("results")[0].split()
            numero_resultados_int = int(numero_resultados[0])

            cantidad_paginas_int = int(a_element_cantidad_paginas)
            print("Numero de resultados: " , numero_resultados_int)
            print("Cantidad de paginas: " , cantidad_paginas_int)
            
            return numero_resultados_int,cantidad_paginas_int

    def Obtener_nombre_iframe(self):
        Nombre_elemento = "N/A"
        try:
            xpath_nombre_iframe = ".//span[@class='display-inline-box dark:text-white line-clamp-1 inherit font-bold text-md normal'] | .//div[@class='a-row']/span[@id='sp_short_strip_title'] | .//span[@class='display-inline-box dark:text-white text-black-700 line-clamp-1 inherit text-md font-bold normal'] | .//div[@id='dynamic-bb']//div//div//div[@style='color: rgb(15, 17, 17); overflow: hidden; text-align: var(--_1g2mm5f0); font-size: 14px; line-height: 20px; --_1g2mm5f1: 1; --_1g2mm5f0: center; display: -webkit-box; -webkit-line-clamp: var(--_1g2mm5f1); -moz-box-orient: vertical; word-break: break-word;']"
            Nombre_elemento = self.driver.find_element(By.XPATH,xpath_nombre_iframe).text
            if Nombre_elemento:
                print("Nombre del producto:",Nombre_elemento)
            else:
                print("Nombre del producto:",Nombre_elemento)
            return Nombre_elemento
        except Exception as e:
            print("No se pudo extraer la informacion de nombre del producto en la pagina" , str(e))

    def Obtener_calificacion_iframe(self):
        Calificacion = "N/A"
        try:
            xpath_calificacion_iframe = ".//div[@class='flex flex-col']//div[@class='flex flex-row'] | //div[@class='a-row sp_short_strip_extra']//a//i"
            Calificacion_encontrar = self.driver.find_element(By.XPATH,xpath_calificacion_iframe)
            Calificacion = Calificacion_encontrar.get_attribute("aria-label")
            if Calificacion:
                Calificacion = Calificacion.split(",")[0]
                print("Calificacion:",Calificacion)
            else:
                clase_estrella = Calificacion_encontrar.get_attribute("class")
            if 'a-star-' in clase_estrella:
                Calificacion = clase_estrella.split('a-star-')[-1].split()[0]
                Calificacion = Calificacion.replace('-', '.')
                Calificacion = Calificacion + ' out of 5 stars'
                print("Calificacion:",Calificacion)
            return Calificacion
        except:
            print("Error al obtener la calificacion")

    def Obtener_precio_iframe(self):
        Precio_anterior = "N/A"
        Precio_completo = "N/A"
        Descuento = "N/A"
        try:
            xpath_precio_anterior_iframe =".//span[@class='display-inline-box dark:text-white text-[0.85em] text-gray-650 whitespace-nowrap inherit direction-ltr'] | .//div[@class='flex flex-row flex-wrap items-baseline gap-0.5 gap-y-0 mb-0.25']//span[@class='display-inline-box dark:text-white text-gray-650 line-through inherit text-[14px]']"

            
            xpath_precio_completo_iframe =".//div[@class='ad']//div[@style='display: flex; flex-direction: row;'] | .//div[@class='a-row sp_short_strip_extra']//span[@class='a-color-base sp_short_strip_price'] | .//div[@class='display-inline-box dark:text-white text-md text-gray-100 whitespace-nowrap inherit direction-ltr leading-[1.2]'] | .//span[@class='display-inline-box dark:text-white text-md text-gray-100 whitespace-nowrap inherit direction-ltr leading-[1.2]']//span[@class='display-inline-box dark:text-white text-black-700 inherit text-[24px]'] | .//span[@class='display-inline-box dark:text-white text-black-700 inherit text-[18px] leading-[18px]'] | //span[@class='display-inline-box dark:text-white text-black-700 inherit text-[24px]'] | .//div[@id='dynamic-bb']//div//div//div//div[@style='display: flex; justify-content: var(--_3jg6vb0); align-items: center; flex-wrap: wrap; column-gap: 4px; --_3jg6vb0: center;']"
            

            try:
                Precio_completo = self.driver.find_element(By.XPATH,xpath_precio_completo_iframe).text.strip().replace('$', '').replace(',', '').replace('\n','.')
                if Precio_completo:
                    print("Precio actual:",Precio_completo)
                else:
                    print("Precio actual:",Precio_completo)
            except:
                print("Precio actual",Precio_completo)
            try:
                Precio_anterior = self.driver.find_element(By.XPATH,xpath_precio_anterior_iframe).text.strip().replace('$', '').replace(',', '')
                if Precio_anterior:
                    print("Precio anterior:",Precio_anterior)
                else:
                    print("Precio anterior:",Precio_anterior)
            except:
                print("Precio anterior:",Precio_anterior)     
            if Precio_completo and Precio_anterior and Precio_anterior != "N/A" and Precio_completo != "N/A":
                try:
                    Precio_completo = float(Precio_completo)
                    Precio_anterior = float(Precio_anterior)
                    Descuento = Precio_anterior - Precio_completo
                    print("Descuento:", Descuento)
                except ValueError as e:
                    print("Error al convertir los precios a números:", e)
            else:  
                if Precio_completo is None:
                    print("Precio completo",Precio_completo)
                else:
                    print("Descuento: N/A")
            return Precio_completo,Precio_anterior,Descuento
        except:
            print("Error al obtener la precio anterior o precio completo")

    def Obtener_nombre_noiframe(self,producto):
        Nombre_elemento = "N/A"
        try:
            xpath_nombre = ".//div[@class='a-section a-spacing-small a-spacing-top-small']//a//h2 | .//h2[@class = 'a-size-base-plus a-spacing-none a-color-base a-text-normal'] | .//div[@class='a-section a-spacing-small']//a[@class='a-link-normal s-line-clamp-3 s-link-style a-text-normal']//h2" 
            Nombre_elemento = producto.find_element(By.XPATH, xpath_nombre).text
            if Nombre_elemento:
                print("Nombre del producto:",Nombre_elemento)
            else:
                print("Nombre del producto:",Nombre_elemento)
            return Nombre_elemento
        except Exception as e:
            print("No se pudo extraer la informacion de nombre del producto en la pagina" , str(e))

    def Obtener_calificacion_noiframe(self,producto):
        Calificacion = "N/A"
        try:
            xpath_calificacion = ".//div[@class='a-row a-size-small']//span[@class='a-declarative']//a | .//a[@class='a-popover-trigger a-declarative']"
            #//span[@class='a-icon-alt'
            CalificacionEnTexto = producto.find_element(By.XPATH,xpath_calificacion).get_attribute("aria-label")
            Calificacion = CalificacionEnTexto.split(",")[0]
            if Calificacion:
                print("Calificacion:",Calificacion)
            else:
                print("Calificacion:",Calificacion)
            return Calificacion
        except:
            print("No se pudo extraer la informacion de calificacion del producto en la pagina")

    def Obtener_precio_noiframe(self,producto):
        Precio_anterior = "N/A"
        Precio_completo = "N/A"
        Descuento = "N/A"
        try:
            try:
                xpath_precio_completo = ".//a[@class='a-link-normal s-no-hover s-underline-text s-underline-link-text s-link-style a-text-normal']//span[@class='a-price']//span[@class='a-offscreen']"
                Precio_completo = producto.find_element(By.XPATH, xpath_precio_completo).get_attribute('textContent').strip().replace('$', '').replace(',', '').strip()
                Precio_completo = float(Precio_completo)
                print("Precio actual:",Precio_completo)
            except NoSuchElementException:
                print("Precio actual:",Precio_completo)

            if isinstance(Precio_completo, (int, float)):
                try:            
                    xpath_precio_anterior = ".//div[@class='a-section aok-inline-block']//span[@class='a-price a-text-price']//span[@class='a-offscreen']"
                    Precio_anterior = producto.find_element(By.XPATH, xpath_precio_anterior).get_attribute("textContent").strip().replace('$', '').replace(',', '').strip()
                    if Precio_anterior:
                        Precio_anterior = float(Precio_anterior)
                        if Precio_completo and Precio_anterior:
                            Descuento = float(Precio_anterior - Precio_completo)
                            print("Descuento:",Descuento)
                        else:
                            Descuento = 0
                            print("No se puede calcular el precio al descuento debido a un precio inválido")
                    else:
                        print("Descuento:",Descuento)
                except NoSuchElementException:
                    Precio_anterior = "N/A"
                    Descuento = "N/A"
                    print("No se encontró precio anterior (sin descuento).")
            else:
                print("Precio completo",Precio_completo)
            if  Precio_completo == "N/A" and Precio_anterior == "N/A":
                print("Producto sin ningun precio")
            return Precio_completo,Precio_anterior,Descuento
        except Exception as e:
            print(f"Error al procesar el precio actual o descuento: {e}")

    def Obtener_patrocinado_noiframe(self,producto):
        Patrocinado = "N/A"
        try:
            # 1. Buscar etiqueta individual de "Sponsored" (para productos individuales fuera del carrusel)
            xpath_patrocinado = ".//span[@class='a-color-secondary' and text()='Sponsored']"
            patrocinado_etiqueta_g = producto.find_elements(By.XPATH, xpath_patrocinado)

            if patrocinado_etiqueta_g:
                Patrocinado = "Patrocinado"
                print("Patrocinado:", Patrocinado)
                return Patrocinado

            # 2. Verificar si el producto está dentro de un contenedor carrusel con etiqueta "Sponsored"
            xpath_sponsored_con_ancestor = (
                ".//ancestor::div[@class='s-include-content-margin s-border-bottom "
                "s-border-top-overlap s-widget-padding-bottom']"
                "//a[@class='a-link-normal aok-inline-block s-widget-sponsored-label-text' and text()='Sponsored']"
            )
            patrocinado_carrusel = producto.find_elements(By.XPATH, xpath_sponsored_con_ancestor)

            if patrocinado_carrusel:
                Patrocinado = "Patrocinado"
                print("Patrocinado:", Patrocinado)
                return Patrocinado

            # 3. Producto general
            Patrocinado = "Producto General"
            print("Patrocinado:", Patrocinado)
            return Patrocinado

        except Exception as e:
            print(f"Error al extraer información de Sponsored: {e}")
            return Patrocinado

    def Almacenar_datos_JSON(self,numero_resultados_int,cantidad_paginas_int):
        json_data = {
            "Numero de resultados": numero_resultados_int,
            "Cantidad de paginas": cantidad_paginas_int
        }
        
        numero_resultados_int, cantidad_paginas_int = self.Obtener_numero_de_paginas_y_resultados()
        self.guardar_en_json(json_data)
        self.leer_archivoJSON()
        
    def Recorrer_Productos(self,cantidad_paginas_int):
        
        Lista_productos1 = []
        Lista_productos2 = []
        pyautogui.hotkey('ctrl', '-')
        for pagina in range(1, cantidad_paginas_int+1):
            print("========================================"*4)
            print(f"Extrayendo datos de la {pagina} de {cantidad_paginas_int}")

            if pagina > 1:
                siguiente_pagina_xpath = ".//a[@class='s-pagination-item s-pagination-next s-pagination-button s-pagination-button-accessibility s-pagination-separator']"
                siguiente_pagina = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, siguiente_pagina_xpath)))
                siguiente_pagina.click()

                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='s-main-slot s-result-list s-search-results sg-row']//div[@class='sg-col-inner']")))

            productos = self.driver.find_elements(By.XPATH,".//div[@class='s-main-slot s-result-list s-search-results sg-row']//div[contains(@class,'puis-card-container')] | .//div[@class='puis-card-container s-card-container s-overflow-hidden aok-relative puis-expand-height puis-include-content-margin puis puis-v24mxf0k74s6dj2mwx0jexn7ud4 puis-card-border'] |.//iframe[@title='Sponsored ad'] | .//div[@class='sg-col-0-of-12 sg-col-4-of-16 sbv-product-container sg-col sg-col-12-of-24 sg-col-8-of-20']")
            time.sleep(2)

            for producto in productos:
                print("========================================"*3)
                print("Pagina numero:", pagina)
                producto_numero = productos.index(producto) + 1
                print("Producto numero:", producto_numero)
                #time.sleep(0.2)

                if producto.tag_name.lower() == "iframe":
                    try:
                        
                        self.driver.switch_to.frame(producto)
                        Patrocinado = "Patrocinado"
                        print(Patrocinado)
                        Nombre_elemento = self.Obtener_nombre_iframe()
                        Calificacion = self.Obtener_calificacion_iframe()
                        Precio_completo,Precio_anterior,Descuento = self.Obtener_precio_iframe()

                        Diccionario_productos2 = {"Nombre del producto" : Nombre_elemento , "Calificacion" : Calificacion , "Precio_actual" : Precio_completo ,"Precio_anterior" : Precio_anterior , "Descuento": Descuento , "Patrocinado": Patrocinado}

                        Lista_productos2.append(Diccionario_productos2)

                        self.driver.switch_to.default_content()
                        continue
                    except Exception as e:
                        print(f"Error al procesar iframe: {e}")
                        self.driver.switch_to.default_content()
                        continue
                        
                try:
                    Nombre_elemento = self.Obtener_nombre_noiframe(producto)
                    Calificacion = self.Obtener_calificacion_noiframe(producto)
                    Precio_completo,Precio_anterior,Descuento = self.Obtener_precio_noiframe(producto)
                    Patrocinado = self.Obtener_patrocinado_noiframe(producto)

                    Diccionario_productos1 = {"Nombre del producto" : Nombre_elemento , "Calificacion" : Calificacion , "Precio_actual" : Precio_completo ,"Precio_anterior" : Precio_anterior , "Descuento": Descuento , "Patrocinado": Patrocinado}

                    Lista_productos1.append(Diccionario_productos1)
                except Exception as e:
                    print(f"Error al procesar producto: {e}")
        Lista_productos = Lista_productos1 + Lista_productos2
        return Lista_productos

    def Procesar_datos_DataFrame(self,Lista_productos,cantidad_paginas_int):
        df = pd.DataFrame(Lista_productos)
        pd.set_option('display.width', 1000)
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        print(df)
        total_productos = len(df)
        print(total_productos)

        total_productos = len(df)
        print(total_productos)
    
        productos_por_paginas = math.ceil(total_productos / cantidad_paginas_int)

        df['Precio_completo'] = pd.to_numeric(df['Precio_actual'], errors='coerce').round()
        df = df.sort_values(by='Precio_completo', ascending=False)

        df['Precio_anterior'] = pd.to_numeric(df['Precio_anterior'], errors='coerce').round()

        df['Precio_completo'] = df['Precio_completo'].fillna("N/A")
        df['Precio_anterior'] = df['Precio_anterior'].fillna("N/A")
        df['Descuento'] = df['Descuento'].fillna("N/A")
        df['Calificacion'] = df['Calificacion'].fillna("N/A")
        df['Patrocinado'] = df['Patrocinado'].fillna("N/A")
        df['Nombre del producto'] = df['Nombre del producto'].fillna("N/A")



        return df, total_productos, productos_por_paginas
    
    def Filtrar_datos_DataFrame(self,df,palabra_clave):

        df_filtrado = df[df["Nombre del producto"].str.contains(palabra_clave, case=False, na=False)]
        print(df_filtrado)
        return df_filtrado

    def Exportar_divisiones_a_excel(self,df,total_productos,productos_por_paginas):

        Divisiones = []
        for i in range(0,total_productos,productos_por_paginas):
            division = df.iloc[i:i+productos_por_paginas]
            Divisiones.append(division)

        with pd.ExcelWriter('Productos_por_paginas_divididos.xlsx',engine="openpyxl") as writer:
            for idx , division in enumerate(Divisiones):
                division.to_excel(writer, sheet_name=f'Hoja_{idx + 1}', index=False)
        
        os.startfile('Productos_por_paginas_divididos.xlsx')

    def Conectar_BD(self):
        try:
            conn = pyodbc.connect('Driver={SQL Server};'
                                'Server=localhost;'
                                'Database=DB_Automatizacion_Reportes;'
                                'UID=sa;'
                                'PWD=1234;')
            print("Conexión exitosa a la base de datos")
            return conn
        except pyodbc.Error as e:
            print(f"Error al conectar a la base de datos: {e}")
            return None
        
    def Insertar_datos_SQLServer(self,df):
        conn = self.Conectar_BD()
        if conn:
            try:
                cursor = conn.cursor()
                for index, row in df.iterrows():
                    try:
                        cursor.execute("""
                            INSERT INTO Productos (Nombre_Completo, Calificacion, Precio_actual, Precio_anterior, Descuento, Patrocinado)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        row['Nombre_del_producto'],
                        row['Calificacion'],
                        row['Precio_actual'],
                        row['Precio_anterior'],
                        row['Descuento'],
                        row['Patrocinado'],
                        row['Precio_completo_redondeado'])
                    except Exception as e:
                        print(f"Error al insertar la fila {index}: {e}")
                conn.commit()
                print("✅ Datos insertados correctamente en la base de datos")
            except Exception as e:
                print(f"Error durante la inserción de datos: {e}")
            finally:
                cursor.close()
                conn.close()
        else:
            print("No se pudo establecer la conexión con la base de datos")

    def Correr_Todo_El_Programa(self):

        self.driver = self.Iniciar_navegacion()

        self.Buscar_objeto_barra_de_busqueda()

        numero_resultados_int,cantidad_paginas_int = self.Obtener_numero_de_paginas_y_resultados()

        self.Almacenar_datos_JSON(numero_resultados_int,cantidad_paginas_int)

        Lista_productos = self.Recorrer_Productos(cantidad_paginas_int)

        df, total_productos, productos_por_paginas = self.Procesar_datos_DataFrame(Lista_productos,cantidad_paginas_int)

        df = self.Filtrar_datos_DataFrame(df,"Iphone 15")

        self.Exportar_divisiones_a_excel(df,total_productos,productos_por_paginas)

        self.Cerrar_navegacion()

        Conexion.conexion(self)

        #self.Insertar_datos_SQLServer(df)

if __name__ == "__main__":
    main = Main()
    main.Correr_Todo_El_Programa()