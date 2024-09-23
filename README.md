# Instrucciones para el uso
- El archivo base de datos es el archivo que contiene toda la informacion sobre los AGA.
    - **AGA - LIM_POB_PARR_BARR 07-2024.xlsx**
    - La aplicacion lo detecta automaticamente cuando lo guarda en el directorio Descargas de su sistema.
- El archivo que desea procesar debe tener las siguientes cabezeras.
    - 'Ingresa el Barrio en que vives'
    - 'Ingresa tu dirección y una referencia'
    - 'Ingresa la Parroquia a la que pertenece tu Sector'
- Debe subir estos dos archivos para habilitar el boton de procesamiento.



# Instruccions for developers

## Notes
- python = ">=3.9,<3.14"


## Commands
```
poetry add pyinstaller -G dev
#
# create a spec file
#pyinstaller --name ExcelApp --onefile --windowed aga_search/gui.py

# create the executable in /dist
poetry run pyinstaller --name SearchAgaApp --onefile --windowed main.py

```