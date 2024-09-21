# Instrucciones para el uso

- El archivo que desea procesar debe tener las siguientes cabezeras.
    - 'Ingresa el Barrio en que vives'
    - 'Ingresa tu dirección y una referencia'
    - 'Ingresa la Parroquia a la que pertenece tu Sector'
- El archivo que se usara como base de datos `AGA - LIM_POB_PARR_BARR 07-2024.xlsx` puede guardarlo en.
    - El directorio Desacargas o Downloads y el programa lo detectara automaticamente. (Recomendado)
    -  Cualquier directorio.
- Para habilitar las opciones el programa debe detectar el archivo `AGA - LIM_POB_PARR_BARR 07-2024.xlsx`.

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
poetry run pyinstaller --name ExcelApp --onefile --windowed main.py

```