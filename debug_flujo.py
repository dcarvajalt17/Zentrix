print("🔧 Ejecutando paso 1: cálculo de parámetros...")
from Calculo_de_niveles_de_consumo import ejecutar_parametros
ejecutar_parametros('data-consumo1.xlsx', 'Referencia V2.xlsx')
print("✅ Paso 1 completado.")

print("📦 Ejecutando paso 2: planificación y exportación...")
from Buffer_Mejorado import exportar_resumen
exportar_resumen('Referencia V2.xlsx', 'data-consumo1.xlsx', 'Resumen_Buffer_NoBuffer_Semanal.xlsx')
print("✅ Paso 2 completado.")

print("🌐 Ejecutando paso 3: lanzando visualizador...")
from Visualizador_DDMRP import lanzar_visualizador
lanzar_visualizador()
print("✅ Paso 3 ejecutado. Visualizador abierto.")

