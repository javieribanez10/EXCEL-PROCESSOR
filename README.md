# Excel Processor - Análisis Inteligente con IA

## 🚀 Deploy en Streamlit Cloud

### Paso 1: Preparar el repositorio
1. Asegúrate de que todos los archivos estén en tu repositorio
2. Verifica que `requirements.txt` contenga todas las dependencias

### Paso 2: Conectar con Streamlit Cloud
1. Ve a [share.streamlit.io](https://share.streamlit.io)
2. Conecta tu cuenta de GitHub
3. Despliega desde tu repositorio

### Paso 3: Configurar la API Key (IMPORTANTE)
Una vez desplegada tu app:

1. **Ve a tu app desplegada**
2. **Haz clic en "Settings" (⚙️) en la esquina superior derecha**
3. **Selecciona "Secrets" en el menú lateral**
4. **Pega el siguiente código reemplazando con tu API key real:**

```toml
OPENAI_API_KEY = "sk-proj-tu-api-key-real-aqui"
```

5. **Haz clic en "Save"**
6. **La app se reiniciará automáticamente**

### Paso 4: Verificar funcionamiento
- La app debería mostrar "✅ API Key configurada" en el sidebar
- Si no funciona, revisa que la API key esté correcta en Secrets

## 🏠 Uso Local

### Opción 1: Archivo .env (Recomendado)
Crea un archivo `.env` en la raíz del proyecto:
```
OPENAI_API_KEY=tu-api-key-aqui
```

### Opción 2: Variable de entorno
```bash
export OPENAI_API_KEY="tu-api-key-aqui"
streamlit run processor_final.py
```

### Opción 3: Interfaz de usuario
Si no tienes configurada la API key, la app te permitirá ingresarla directamente en el sidebar.

## 🔧 Instalación Local

```bash
# Clonar repositorio
git clone tu-repositorio
cd Excel-processor

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar la aplicación
streamlit run processor_final.py
```

## 📋 Características

- ✅ Análisis automático de archivos Excel/CSV
- ✅ Generación de resúmenes ejecutivos con IA
- ✅ Chat interactivo con tus datos
- ✅ Exportación de análisis y contextos
- ✅ Interfaz web intuitiva
- ✅ Soporte para múltiples hojas de Excel

## 🔐 Seguridad

- Las API keys se manejan de forma segura
- No se almacenan en el código fuente
- Solo se muestran caracteres parciales por seguridad
- Los datos se procesan localmente en la sesión