# IDEP — Instrumento Diagnóstico de Ecosistemas Productivos
**Backend Python / FastAPI + Frontend Web**  
Todas las respuestas se consolidan en tiempo real en un único archivo Excel.

---

## 📁 Estructura del proyecto

```
idep_server/
├── server.py           ← Servidor Python (FastAPI)
├── requirements.txt    ← Dependencias
├── static/
│   └── index.html      ← Sitio web del formulario
└── data/
    └── IDEP_Respuestas_Consolidadas.xlsx  ← Se crea automáticamente
```

---

## ▶️ Cómo ejecutar

### 1. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 2. Iniciar el servidor
```bash
python server.py
```

O con uvicorn directamente:
```bash
uvicorn server:app --host 0.0.0.0 --port 8000 --reload
```

### 3. Abrir en el navegador
```
http://localhost:8000
```

---

## 🌐 API Endpoints

| Método | Ruta           | Descripción                              |
|--------|----------------|------------------------------------------|
| GET    | `/`            | Sirve el formulario web                  |
| POST   | `/api/submit`  | Recibe y guarda una respuesta            |
| GET    | `/api/download`| Descarga el Excel consolidado            |
| GET    | `/api/count`   | Retorna el número de respuestas          |
| GET    | `/api/status`  | Estado del servidor                      |

---

## 📊 Archivo Excel consolidado

`data/IDEP_Respuestas_Consolidadas.xlsx` contiene:
- **Hoja "Consolidado"**: todas las respuestas en filas, lista para análisis cuantitativo
- **Hoja individual por participante**: respuesta detallada con formato visual

---

## 🚀 Despliegue en servidor remoto (para acceso desde internet)

### Opción A — Railway.app (gratis)
1. Subir el proyecto a GitHub
2. Conectar en railway.app → "New Project from GitHub"
3. Railway detecta automáticamente FastAPI
4. La URL pública queda disponible en minutos

### Opción B — VPS (Digital Ocean, Linode, etc.)
```bash
# Instalar
pip install -r requirements.txt
# Ejecutar en background
nohup python server.py &
```

---

## 📚 Referencias bibliográficas (Q1/Q2)
- Carayannis, E. G., & Campbell, D. F. J. (2010). *Int. J. Social Ecology and Sustainable Development, 1*(1).
- Jacobides, M. G. et al. (2018). *Strategic Management Journal, 39*(8), 2255–2276.
- Granstrand, O., & Holgersson, M. (2020). *Technovation, 90–91*, 102098.
- Gereffi, G., & Fernandez-Stark, K. (2016). Global Value Chain Analysis. CGGC.
