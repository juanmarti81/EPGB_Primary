# Conexión con Primary API usando Python y planilla EPGB

## **Instalación:**

1. Instalar dependencias utilizando "pip install requirements.txt"

2. Crear archivo config.py

   Escribir:

   ```
   COMITENTE = 'numero de comitente'

   PASSWORD = 'password'
   ```

## **Configuraciones:**

La hoja "**Tickers**" contiene todos los tickers que se van a suscribir para obtener datos. La primer columna es el ticker utilizado por Primary y la segunda es el propio simbolo que vemos en las plataformas.

La actualización de datos se realiza cada 1 segundo. Se puede cambiar en la siguiente línea:

```
time.sleep(1)
```

Planilla creada por Guillermo Cutella @gcutte

pyRofex creado por Primary https://github.com/matbarofex/pyRofex
