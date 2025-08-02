import pandas as pd
import pyodbc 

# Configuración de conexión
conn = pyodbc.connect(
    'DRIVER={SQL Server};'
    'SERVER=test;'
    'DATABASE=DB_1;'
    'Trusted_Connection=yes;'
)

# Columnas a leer del Excel
columnas = [0, 1, 2, 3, 4, 5]
errores = []

# Leer el archivo Excel
dataframe = pd.read_excel('../Libro.xlsx', usecols=columnas)
print(dataframe) # Se imprime las tablas para verificar que sean las que se necesitan

for index, row in dataframe.iterrows():
    try:
        cursor = conn.cursor()
        
        # PRIMERO: Verificar si el artículo existe
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM table 
                WHERE articulo = ? OR Codigo_proveedor = ?
            )
            BEGIN
                RAISERROR('El artículo no existe en la base de datos', 16, 1)
            END
        """, row["Articulo"], row["Noarticulo"])
        
        # SEGUNDO: Si pasa la validación, ejecutar el procedimiento
        cursor.execute("""
            EXEC DB_1.insertar_archivo_excel 
                @Articulo = ?, 
                @Codigo_proveedor = ?, 
                @Marca = ?, 
                @proveedor = 188, 
                @Cantidad = ?, 
                @costo = ?, 
                @numero_orden = ?, 
                @documento = ?, 
                @creado_por = 'jose'
        """, 
        row["Articulo"], row["Noarticulo"], row["marca"], 
        row["CantidadConfirmada"], row["PC"], row["Noorden"], row["Noorden"])
        
        conn.commit()  
        
    except Exception as e:
        errores.append({
            "Articulo": row["Articulo"],
            "Noarticulo": row["Noarticulo"],
            "marca": row["marca"],
            "CantidadConfirmada": row["CantidadConfirmada"],
            "PC": row["PC"],
            "Noorden": row["Noorden"],
            "Error": str(e)
        })
        conn.rollback()
        print(f"Error procesando fila {index}: {e}")

# Guardar errores en CSV
if errores:
    pd.DataFrame(errores).to_csv("errores.csv", index=False)
    print(f"Se encontraron {len(errores)} errores. Ver 'errores.csv'")
else:
    print("Proceso completado sin errores")

# Cerrar conexión
conn.close()