import os
import time

# downloads path for any operating system
DOWNLOADS_PATH = os.path.expanduser("~/Downloads")

# patterns
DB_FILE_NAME = 'AGA - LIM_POB_PARR_BARR 07-2024.xlsx'
DB_FILE_PATH = os.path.join(DOWNLOADS_PATH, DB_FILE_NAME)

# start and end sheet names
START_SHEET_NAME = 'A-01 - BARRIOS'
END_SHEET_NAME = 'A-18 - BARRIOS'
COLUMNS = 'B, C'
# remove the first row
ROW_START = 1
BASE_PERCENTAGE_SIMILARITY = 0.6
BOOST = 0.2
LEVEL = 0
HEADER_ROW = 1

# FILE_NAME = 'Copy of ASISTENCIA A TALLERES (respuestas).xlsx'
FILE_NAME = 'test1.xlsx'
FILE_PATH = os.path.join(DOWNLOADS_PATH, FILE_NAME)
# first sheet is index 0
SHEET_NAME = 'Addressess and AGA'
SHEET_INDEX = 0
COLUMN1 = 'Ingresa el Barrio en que vives'
COLUMN2 = 'Ingresa tu dirección y una referencia'
COLUMN3 = 'Ingresa la Parroquia a la que pertenece tu Sector'
FILE_COLUMNS = [COLUMN1, COLUMN2, COLUMN3]
COLUMN_INDEX = COLUMN2
NEW_COLUMNS_NAME = ['AGA', 'address']

NEW_FILE_NAME = f'{FILE_NAME}-{time.strftime("%Y-%m-%d-%H-%M-%S")}.xlsx'
NEW_FILE_NAME_PATH = os.path.join(DOWNLOADS_PATH, NEW_FILE_NAME)

PATTERN_FILE_NAME = 'AGA - LIM_POB_PARR_BARR 07-2024.xlsx'
PATTERN_FILE_PATH = os.path.join(DOWNLOADS_PATH, PATTERN_FILE_NAME)
