 Documentación TécnicaDocumentación Técnica
=====================

Descripción
-----------

Este código implementa un proceso de Web Scraping y RPA utilizando Python y la biblioteca Selenium. El objetivo es extraer precios de productos de un sitio web y almacenarlos en un archivo Excel. También se incluye la funcionalidad de programar el proceso para que se ejecute en un horario específico.

Documentación
-------------

### Clase ProcesoWebScrapingRPA

Esta clase es responsable de realizar el proceso de Web Scraping y RPA.

#### Métodos

- `cargar_archivo_xlsx()`: Carga un archivo Excel para extraer los datos.
- `obtener_total_productos()`: Obtiene el número total de productos en el archivo.
- `iniciar_proceso()`: Inicia el proceso de Web Scraping y RPA.
- `actualizar_estado_bot()`: Actualiza el estado de un producto en el archivo.
- `actualizar_progreso()`: Actualiza el progreso del proceso en la interfaz.
- `obtener_valor_span()`: Obtiene el valor de un producto desde el sitio web.
- `actualizar_valor_xpath()`: Actualiza el valor XPath de un producto en el archivo.
- `guardar_archivo_xlsx()`: Guarda los cambios en el archivo Excel.
- `cerrar_webdriver()`: Cierra el navegador web.
- `mostrar_producto_precio()`: Muestra el nombre y el precio de un producto en la interfaz.
 
### Funciones auxiliares

- `cargar_archivo()`: Abre un cuadro de diálogo para seleccionar un archivo Excel.
- `iniciar_proceso_scraping()`: Inicia el proceso de Web Scraping y RPA.
- `programar_proceso_scraping()`: Programa el proceso para que se ejecute en un horario específico.
- `crear_grafica_precios()`: Crea una gráfica de barras con los precios obtenidos.
 
Dependencias
------------

- [Python](https://www.python.org/)
- [Biblioteca openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [Biblioteca Selenium](https://selenium-python.readthedocs.io/)
- [Biblioteca Schedule](https://schedule.readthedocs.io/)
- [Biblioteca Matplotlib](https://matplotlib.org/)
 
Recursos
--------

- [Documentación de la API de Selenium para Python](https://selenium-python.readthedocs.io/api.html)
- [Documentación de la biblioteca openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [Documentación de la biblioteca Tkinter de Python](https://docs.python.org/3/library/tkinter.html)
- [Documentación de la biblioteca Matplotlib](https://matplotlib.org/stable/contents.html)
 
Ejecución del Código
--------------------

1. Instalación de dependencias:

- Asegúrate de tener instalado Python en tu sistema. Puedes descargarlo desde el sitio web oficial de Python: 
- Instala las dependencias necesarias ejecutando el siguiente comando en la línea de comandos:
 
```
pip install -r requirements.txt
```

19. Descarga de archivos adicionales:
- Descarga el archivo `chromedriver.exe` para tu versión de Chrome
- Coloca el archivo `chromedriver.exe` en el mismo directorio que el archivo Python.
 
21. Ejecución del código:
- Abre una terminal o línea de comandos y navega hasta el directorio donde se encuentra el archivo Python.
- Ejecuta el siguiente comando:
 
```
python app.py
```