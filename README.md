# 💰 Panel de Gastos 2026

Aplicación web local que lee y escribe el Excel **CONTROL DE GASTOS 2026 CLOUD.xlsx**.

## Qué hace

- **Resumen mensual y anual**: gastos, disponible, ahorro y totales del año.
- **Comparativa entre meses**: gráfico de gastos de cada mes.
- **Listado detallado**: búsqueda por concepto y filtro por categoría/mes.
- **Añadir gasto**: escribe el gasto nuevo en la hoja del mes correspondiente,
  recolocando el `=SUM` y haciendo **backup** (`CONTROL DE GASTOS 2026 CLOUD.bak.xlsx`)
  antes de guardar.

## Cómo ejecutarlo

1. Doble clic en **`iniciar.bat`**.
2. Se abre el navegador en `http://127.0.0.1:8720`.

También puedes arrancarlo a mano:

```
py -3 servidor.py
```

El servidor usa el puerto **8720** por defecto. Para otro puerto:

```
py -3 servidor.py 9000
```

## Requisitos

- Python 3.9+ con el paquete **openpyxl**:

```
py -m pip install openpyxl
```

## Cómo encuentra el Excel

1. Si existe `config.json` con una clave `excel_path` válida, usa esa ruta.
2. Si no, busca en Escritorio/Documentos un `.xlsx` que contenga `gasto` y `2026`
   en el nombre (preferirá `control de gastos 2026 cloud.xlsx`).

Puedes cambiar la ruta editando `config.json`.

## Notas importantes

- Si al guardar sale **"El Excel está abierto en otro programa"**, ciérralo en
  Excel/OneDrive y reintenta.
- Las tablas de resumen de la **parte derecha** de cada hoja (resumen por categoría,
  CTA Carlos/Karla) se mantienen a mano en Excel. Al insertar un gasto, la fila se
  inserta igual que lo hace tu bot de Telegram: revisa que esas tablas laterales
  sigan cuadradas.
- Cada guardado crea una copia de seguridad `.bak.xlsx` al lado del Excel.
