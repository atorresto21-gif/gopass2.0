# Bot Telegram - Consolidado de Transacciones

Bot simple. Tu mandas Excel, bot devuelve Excel ordenado.

## Que hace el bot

1. Recibe tu archivo Excel (.xlsx).
2. Quita todas las filas de **Contratos**.
3. Separa por **Placa** → una hoja por cada placa.
4. Ordena por **Fecha de la transacción**.
5. Agrupa por **mes** y pone **AUTOSUMA (=SUM)** por cada mes.
6. Te devuelve el Excel ordenado.

El bot te pregunta el orden:
- Mes más reciente primero
- Mes más antiguo primero

El formato tabular original NO se modifica. Solo se agrupa y se agrega la fila de total por mes (con fórmula real, no número pegado).

## Archivos

- `bot.py` — el bot de Telegram.
- `processor.py` — la lógica que procesa el Excel.
- `requirements.txt` — librerías.
- `Procfile` — dice a Railway que corra `python bot.py` como worker.
- `runtime.txt` — versión de Python.

## Pasos para subir a Railway

### 1. Crear el bot en Telegram
1. En Telegram abre **@BotFather**.
2. Escribe `/newbot`.
3. Pon nombre y usuario.
4. BotFather te da un **TOKEN**. Cópialo (algo como `123456:ABC-DEF...`).

### 2. Subir el código
Opción A — GitHub (recomendado):
1. Sube esta carpeta a un repositorio de GitHub.
2. En Railway → **New Project** → **Deploy from GitHub repo**.
3. Selecciona el repo.

Opción B — Railway CLI:
```bash
npm i -g @railway/cli
railway login
railway init
railway up
```

### 3. Poner el TOKEN en Railway
1. En tu proyecto de Railway → pestaña **Variables**.
2. Agrega:
   - **Nombre:** `TELEGRAM_BOT_TOKEN`
   - **Valor:** el token de BotFather
3. Guarda.

### 4. Verificar el servicio
- Railway debe correr el proceso **worker** (lo dice el `Procfile`).
- Si Railway crea un servicio "web" por error, no pasa nada; este bot usa *polling*, no necesita puerto.
- Revisa la pestaña **Deployments / Logs**. Debe decir: `Bot iniciado.`

### 5. Usar
1. Abre tu bot en Telegram.
2. `/start`.
3. Manda el archivo Excel.
4. Elige el orden.
5. Recibes el Excel procesado.

## Probar local (opcional)
```bash
pip install -r requirements.txt
export TELEGRAM_BOT_TOKEN="tu_token"
python bot.py
```

## Notas
- Si tu Excel tiene columnas con otros nombres, el bot busca de forma tolerante: `Placa`, `Servicio`, `Valor` y `Fecha de la transacción` (ignora acentos y espacios).
- Si una fila de Contratos no tiene placa, igual queda excluida.
- Soporta varias placas: una hoja por placa automáticamente.
