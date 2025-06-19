import streamlit as st
import pandas as pd
import numpy as np
import os
import json
import asyncio
from typing import Dict, List, Any
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from dotenv import load_dotenv

# Cargar variables de entorno (solo funciona en local)
load_dotenv()

# Configuración de la página
st.set_page_config(
    page_title="Excel Processor - Análisis Inteligente",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def get_api_key():
    """Obtiene la API key desde múltiples fuentes"""
    # 1. Primero desde secrets de Streamlit Cloud
    try:
        if hasattr(st, 'secrets') and 'OPENAI_API_KEY' in st.secrets:
            return st.secrets['OPENAI_API_KEY']
    except:
        pass
    
    # 2. Desde variable de entorno (local)
    api_key = os.getenv("OPENAI_API_KEY")
    if api_key and len(api_key) > 10:
        return api_key
    
    # 3. Desde session state (configurado por el usuario)
    if 'api_key' in st.session_state and st.session_state.api_key:
        return st.session_state.api_key
    
    return None

def configure_api_key():
    """Interfaz para configurar la API key si no está disponible"""
    st.sidebar.markdown("### 🔑 Configuración OpenAI")
    
    api_key = get_api_key()
    
    if api_key and len(api_key) > 10:
        st.sidebar.success("✅ API Key configurada")
        # Mostrar solo los primeros y últimos caracteres por seguridad
        masked_key = f"{api_key[:8]}...{api_key[-8:]}" if len(api_key) > 16 else "***"
        st.sidebar.text(f"Key: {masked_key}")
        return True
    else:
        st.sidebar.error("❌ API Key requerida")
        
        # Campo para ingresar la API key
        user_api_key = st.sidebar.text_input(
            "Ingresa tu OpenAI API Key:",
            type="password",
            help="Tu API key de OpenAI (sk-...)",
            placeholder="sk-proj-..."
        )
        
        if user_api_key:
            if user_api_key.startswith('sk-') and len(user_api_key) > 20:
                st.session_state.api_key = user_api_key
                st.sidebar.success("✅ API Key guardada para esta sesión")
                st.experimental_rerun()
            else:
                st.sidebar.error("❌ API Key inválida. Debe empezar con 'sk-'")
        
        st.sidebar.markdown("""
        **Para Streamlit Cloud:**
        1. Ve a tu app en Streamlit Cloud
        2. Settings → Secrets
        3. Agrega: `OPENAI_API_KEY = "tu-api-key"`
        
        **Para uso local:**
        Crea un archivo `.env` con:
        ```
        OPENAI_API_KEY=tu-api-key
        ```
        """)
        
        return False

class ExcelProcessor:
    def __init__(self):
        self.client = None
    
    def get_openai_client(self):
        """Obtiene el cliente de OpenAI de forma lazy"""
        if self.client is None:
            api_key = get_api_key()
            if api_key and len(api_key) > 10:
                try:
                    from openai import OpenAI
                    self.client = OpenAI(api_key=api_key)
                    return True
                except Exception as e:
                    st.error(f"Error configurando OpenAI: {e}")
                    return False
            else:
                return False
        return True
    
    def load_file(self, uploaded_file) -> Dict[str, pd.DataFrame]:
        """Carga archivo Excel/CSV"""
        try:
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            if file_extension == 'csv':
                try:
                    df = pd.read_csv(uploaded_file, encoding='utf-8')
                except UnicodeDecodeError:
                    try:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, encoding='latin-1')
                    except UnicodeDecodeError:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, encoding='cp1252')
                return {"Sheet1": df}
            
            elif file_extension in ['xlsx', 'xls']:
                excel_file = pd.ExcelFile(uploaded_file)
                sheets = {}
                for sheet_name in excel_file.sheet_names:
                    try:
                        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                        sheets[sheet_name] = df
                    except Exception as e:
                        st.warning(f"⚠️ Error leyendo la hoja '{sheet_name}': {e}")
                        continue
                return sheets
            
            else:
                st.error(f"Formato no soportado: {file_extension}")
                return {}
                
        except Exception as e:
            st.error(f"Error cargando archivo: {e}")
            return {}
    
    def analyze_dataframe_basic(self, df: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
        """Análisis básico de un DataFrame"""
        try:
            analysis = {
                "sheet_name": sheet_name,
                "shape": df.shape,
                "columns": list(df.columns),
                "dtypes": {str(col): str(dtype) for col, dtype in df.dtypes.items()},
                "null_counts": df.isnull().sum().to_dict(),
                "memory_usage": df.memory_usage(deep=True).sum(),
                "numeric_columns": df.select_dtypes(include=[np.number]).columns.tolist(),
                "categorical_columns": df.select_dtypes(include=['object']).columns.tolist(),
            }
            
            # Estadísticas básicas para numéricas
            if analysis["numeric_columns"]:
                numeric_stats = {}
                for col in analysis["numeric_columns"]:
                    try:
                        stats = df[col].describe()
                        numeric_stats[col] = {
                            'count': float(stats['count']),
                            'mean': float(stats['mean']),
                            'std': float(stats['std']),
                            'min': float(stats['min']),
                            'max': float(stats['max'])
                        }
                    except:
                        numeric_stats[col] = {"error": "Error calculando estadísticas"}
                analysis["numeric_stats"] = numeric_stats
            
            # Info básica categórica
            categorical_info = {}
            for col in analysis["categorical_columns"]:
                try:
                    value_counts = df[col].value_counts().head(10)
                    categorical_info[col] = {
                        "unique_count": df[col].nunique(),
                        "top_values": value_counts.to_dict()
                    }
                except:
                    categorical_info[col] = {"error": "Error procesando columna"}
            
            analysis["categorical_info"] = categorical_info
            return analysis
            
        except Exception as e:
            st.error(f"Error en análisis: {e}")
            return {"error": str(e)}
    
    def generate_llm_context(self, sheets: Dict[str, pd.DataFrame], filename: str) -> str:
        """Genera contexto completo para LLM"""
        context = f"""
# CONTEXTO COMPLETO PARA LLM - ARCHIVO: {filename}

## METADATOS DEL ARCHIVO
- Nombre: {filename}
- Extensión: {filename.split('.')[-1].upper() if '.' in filename else 'Desconocida'}
- Hojas totales: {len(sheets)}
- Nombres de hojas: {', '.join(sheets.keys())}
- Fecha de análisis: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## ANÁLISIS POR HOJA
"""
        
        for sheet_name, df in sheets.items():
            analysis = self.analyze_dataframe_basic(df, sheet_name)
            
            context += f"""
### HOJA: "{sheet_name}"
- Dimensiones: {analysis['shape'][0]:,} filas × {analysis['shape'][1]} columnas
- Memoria: {analysis['memory_usage'] / 1024:.1f} KB

#### COLUMNAS ({len(analysis['columns'])})
"""
            
            for col in analysis['columns']:
                dtype = analysis['dtypes'].get(col, 'unknown')
                null_count = analysis['null_counts'].get(col, 0)
                null_percentage = (null_count / analysis['shape'][0]) * 100 if analysis['shape'][0] > 0 else 0
                
                context += f"""
**{col}** ({dtype})
- Valores nulos: {null_count:,} ({null_percentage:.1f}%)
"""
                
                # Agregar info específica si es numérica o categórica
                if col in analysis.get('numeric_columns', []) and 'numeric_stats' in analysis:
                    stats = analysis['numeric_stats'].get(col, {})
                    if 'error' not in stats:
                        context += f"- Rango: {stats.get('min', 0):.2f} a {stats.get('max', 0):.2f} (media: {stats.get('mean', 0):.2f})\n"
                
                elif col in analysis.get('categorical_columns', []):
                    cat_info = analysis.get('categorical_info', {}).get(col, {})
                    if 'error' not in cat_info:
                        context += f"- Valores únicos: {cat_info.get('unique_count', 0)}\n"
                        top_values = list(cat_info.get('top_values', {}).keys())[:3]
                        if top_values:
                            # Convertir todos los valores a string para evitar errores de join
                            top_values_str = [str(val) for val in top_values]
                            context += f"- Top valores: {', '.join(top_values_str)}\n"
                
                context += "\n"
        
        context += f"""
## RESUMEN EJECUTIVO
Este archivo contiene {len(sheets)} hoja(s) con un total de {sum(df.shape[0] for df in sheets.values()):,} filas de datos.

## INSTRUCCIONES PARA IA
- Este contexto contiene análisis completo de todas las hojas
- Usar esta información para responder preguntas específicas
- Considerar la calidad de datos (nulos, tipos, etc.)
"""
        
        return context

    async def analyze_with_ai(self, context: str, user_query: str = None) -> str:
        """Análisis con IA"""
        if not self.get_openai_client():
            return "❌ Cliente OpenAI no disponible"
        
        try:
            prompt = f"""
Analiza los siguientes datos de Excel/CSV:

{context}

{f"Pregunta específica: {user_query}" if user_query else ""}

Proporciona un análisis detallado incluyendo:
1. Resumen general de los datos
2. Análisis de columnas importantes
3. Patrones identificados
4. Recomendaciones
"""
            
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "Eres un analista de datos experto."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"❌ Error en análisis IA: {str(e)}"

    async def generate_executive_summary(self, sheets: Dict[str, pd.DataFrame], filename: str) -> str:
        """Genera resumen ejecutivo con IA"""
        if not self.get_openai_client():
            return "❌ No se puede generar resumen ejecutivo: Cliente OpenAI no disponible"
        
        try:
            llm_context = self.generate_llm_context(sheets, filename)
            
            prompt = f"""
Eres un analista de datos SENIOR. Genera un RESUMEN EJECUTIVO PROFESIONAL basado en:

{llm_context}

INCLUYE:

## 📋 1. RESUMEN EJECUTIVO GENERAL
- Tipo de documento empresarial
- Propósito principal y contexto
- Calidad y completitud de los datos

## 🔍 2. ANÁLISIS DE DATOS
- Descripción de cada hoja
- Análisis de columnas importantes
- Patrones y tendencias identificados

## 📊 3. INSIGHTS Y HALLAZGOS
- Problemas de calidad detectados
- Métricas clave identificadas
- Oportunidades de mejora

## 🎯 4. RECOMENDACIONES
- Análisis adicionales recomendados
- Mejoras en calidad de datos
- Próximos pasos sugeridos

Sé detallado y profesional.
"""
            
            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {"role": "system", "content": "Eres un analista de datos senior especializado en crear resúmenes ejecutivos."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=3000
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"❌ Error generando resumen ejecutivo: {str(e)}"
    
    def clean_dataframe_for_display(self, df: pd.DataFrame) -> pd.DataFrame:
        """Limpia el DataFrame para evitar errores de Arrow en Streamlit"""
        try:
            # Crear una copia para no modificar el original
            df_clean = df.copy()
            
            # Convertir todas las columnas problemáticas a string
            for col in df_clean.columns:
                # Si la columna tiene tipos mixtos o es problemática
                if df_clean[col].dtype == 'object':
                    try:
                        # Intentar convertir a numeric primero
                        pd.to_numeric(df_clean[col], errors='raise')
                    except (ValueError, TypeError):
                        # Si falla, convertir todo a string
                        df_clean[col] = df_clean[col].astype(str)
                        df_clean[col] = df_clean[col].replace('nan', '')
                        df_clean[col] = df_clean[col].replace('None', '')
                
                # Manejar columnas con nombres problemáticos
                if str(col).startswith('Unnamed:'):
                    # Renombrar columnas sin nombre
                    new_name = f"Col_{col.split(':')[1].strip()}"
                    df_clean.rename(columns={col: new_name}, inplace=True)
            
            return df_clean
            
        except Exception as e:
            st.warning(f"Error limpiando DataFrame: {e}")
            # Si falla la limpieza, convertir todo a string como último recurso
            df_string = df.astype(str)
            return df_string

def main():
    # Título
    st.title("📊 Excel Processor - Análisis Inteligente")
    st.markdown("### Procesamiento y análisis automático de archivos Excel/CSV con IA")
    
    # Configurar API Key (esto detiene la app si no hay API key)
    api_key_configured = configure_api_key()
    
    # Crear procesador
    processor = ExcelProcessor()
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        st.markdown("---")
        st.markdown("### 📋 Instrucciones")
        st.markdown("""
        1. Configura tu API key de OpenAI
        2. Sube tu archivo Excel/CSV
        3. Explora el análisis automático
        4. Usa la IA para insights
        """)
    
    # Solo continuar si la API key está configurada
    if not api_key_configured:
        st.warning("⚠️ Configura tu API key de OpenAI para continuar")
        return
        
    # File uploader
    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel o CSV",
        type=['xlsx', 'xls', 'csv'],
        help="Formatos: Excel (.xlsx, .xls) y CSV (.csv)"
    )
    
    if uploaded_file is not None:
        st.info(f"📁 Archivo: **{uploaded_file.name}** ({uploaded_file.size:,} bytes)")
        
        # Cargar archivo
        with st.spinner("Cargando archivo..."):
            sheets = processor.load_file(uploaded_file)
        
        if sheets:
            st.success(f"✅ Archivo cargado - {len(sheets)} hoja(s)")
            
            # Tabs con funcionalidades básicas + IA
            tabs = st.tabs([
                "📊 Vista General", 
                "🔍 Análisis Detallado", 
                "🤖 Análisis con IA",
                "📄 Exportación"
            ])
            
            with tabs[0]:  # Vista General
                st.subheader("📋 Resumen del Archivo")
                
                # Métricas generales
                total_rows = sum(df.shape[0] for df in sheets.values())
                total_cols = sum(df.shape[1] for df in sheets.values())
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Hojas", len(sheets))
                with col2:
                    st.metric("Total Filas", f"{total_rows:,}")
                with col3:
                    st.metric("Total Columnas", f"{total_cols:,}")
                with col4:
                    st.metric("Total Celdas", f"{total_rows * total_cols:,}")
                
                # Mostrar cada hoja
                for sheet_name, df in sheets.items():
                    with st.expander(f"📄 Hoja: {sheet_name}", expanded=True):
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.metric("Filas", f"{df.shape[0]:,}")
                        with col2:
                            st.metric("Columnas", f"{df.shape[1]:,}")
                        with col3:
                            st.metric("Celdas", f"{df.shape[0] * df.shape[1]:,}")
                        
                        st.subheader("Vista previa")
                        # Limpiar DataFrame antes de mostrarlo
                        df_clean = processor.clean_dataframe_for_display(df)
                        st.dataframe(df_clean.head(10), use_container_width=True)
                        
                        st.subheader("Info de columnas")
                        col_info = pd.DataFrame({
                            'Columna': df.columns,
                            'Tipo': df.dtypes.astype(str),
                            'Nulos': df.isnull().sum(),
                            '% Nulos': (df.isnull().sum() / len(df) * 100).round(2)
                        })
                        # Limpiar también el DataFrame de info de columnas
                        col_info_clean = processor.clean_dataframe_for_display(col_info)
                        st.dataframe(col_info_clean, use_container_width=True)
            
            with tabs[1]:  # Análisis Detallado
                st.subheader("🔍 Análisis Detallado por Hoja")
                
                selected_sheet = st.selectbox("Seleccionar hoja:", list(sheets.keys()))
                
                if selected_sheet:
                    df = sheets[selected_sheet]
                    analysis = processor.analyze_dataframe_basic(df, selected_sheet)
                    
                    # Estadísticas generales
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("Filas", f"{analysis['shape'][0]:,}")
                        st.metric("Columnas", f"{analysis['shape'][1]:,}")
                    
                    with col2:
                        st.metric("Memoria", f"{analysis['memory_usage'] / 1024:.1f} KB")
                        st.metric("Cols. Numéricas", len(analysis['numeric_columns']))
                    
                    with col3:
                        st.metric("Cols. Categóricas", len(analysis['categorical_columns']))
                        null_total = sum(analysis['null_counts'].values())
                        st.metric("Valores Nulos", f"{null_total:,}")
                    
                    # Valores nulos
                    if any(analysis['null_counts'].values()):
                        st.subheader("🚫 Valores Nulos")
                        null_df = pd.DataFrame([
                            {"Columna": k, "Valores Nulos": v, "% Nulos": f"{v/analysis['shape'][0]*100:.1f}%"}
                            for k, v in analysis['null_counts'].items() if v > 0
                        ])
                        st.dataframe(null_df, use_container_width=True)
                    else:
                        st.success("✅ No hay valores nulos")
                    
                    # Estadísticas numéricas
                    if analysis.get('numeric_stats'):
                        st.subheader("📈 Estadísticas Numéricas")
                        stats_df = pd.DataFrame(analysis['numeric_stats']).T
                        st.dataframe(stats_df, use_container_width=True)
                    
                    # Info categórica
                    if analysis.get('categorical_info'):
                        st.subheader("📝 Columnas Categóricas")
                        for col, info in analysis['categorical_info'].items():
                            if 'error' not in info:
                                with st.expander(f"Columna: {col}"):
                                    st.write(f"**Valores únicos:** {info['unique_count']}")
                                    if info['top_values']:
                                        st.write("**Top 5 valores:**")
                                        top_df = pd.DataFrame([
                                            {"Valor": k, "Frecuencia": v}
                                            for k, v in list(info['top_values'].items())[:5]
                                        ])
                                        st.dataframe(top_df, use_container_width=True)
            
            with tabs[2]:  # Análisis con IA
                st.subheader("🤖 Análisis Inteligente con IA")
                
                if not processor.get_openai_client():
                    st.warning("⚠️ Configura tu API key para usar IA")
                else:
                    # Generar resumen ejecutivo
                    if st.button("🧠 Generar Resumen Ejecutivo", type="primary"):
                        with st.spinner("🤖 Generando resumen ejecutivo..."):
                            try:
                                loop = asyncio.new_event_loop()
                                asyncio.set_event_loop(loop)
                                summary = loop.run_until_complete(
                                    processor.generate_executive_summary(sheets, uploaded_file.name)
                                )
                                
                                st.subheader("📋 Resumen Ejecutivo")
                                st.markdown(summary)
                                
                                # Guardar en session state
                                st.session_state['executive_summary'] = summary
                                
                            except Exception as e:
                                st.error(f"Error generando resumen: {e}")
                    
                    st.markdown("---")
                    
                    # Chat interactivo
                    st.subheader("💬 Chat con tus Datos")
                    user_question = st.text_area(
                        "¿Qué quieres saber sobre este archivo?",
                        placeholder="Ej: ¿Cuáles son las tendencias principales? ¿Hay anomalías?",
                        height=100
                    )
                    
                    if st.button("🔍 Analizar") and user_question:
                        with st.spinner("🤖 Procesando pregunta..."):
                            llm_context = processor.generate_llm_context(sheets, uploaded_file.name)
                            
                            try:
                                loop = asyncio.new_event_loop()
                                asyncio.set_event_loop(loop)
                                response = loop.run_until_complete(
                                    processor.analyze_with_ai(llm_context, user_question)
                                )
                                
                                st.subheader("🤖 Respuesta")
                                st.markdown(response)
                            except Exception as e:
                                st.error(f"Error: {e}")
            
            with tabs[3]:  # Exportación
                st.subheader("📄 Exportación")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("### 📄 Contexto LLM")
                    if st.button("📄 Generar Contexto"):
                        llm_context = processor.generate_llm_context(sheets, uploaded_file.name)
                        st.download_button(
                            label="💾 Descargar Contexto TXT",
                            data=llm_context,
                            file_name=f"contexto_{uploaded_file.name.split('.')[0]}.txt",
                            mime="text/plain"
                        )
                    
                    st.write("### 📊 Análisis JSON")
                    if st.button("📊 Generar JSON"):
                        complete_analysis = {}
                        for sheet_name, df in sheets.items():
                            complete_analysis[sheet_name] = processor.analyze_dataframe_basic(df, sheet_name)
                        
                        json_data = json.dumps(complete_analysis, indent=2, default=str)
                        st.download_button(
                            label="💾 Descargar JSON",
                            data=json_data,
                            file_name=f"analisis_{uploaded_file.name.split('.')[0]}.json",
                            mime="application/json"
                        )
                
                with col2:
                    st.write("### 📋 Resumen Ejecutivo")
                    if 'executive_summary' in st.session_state:
                        st.success("✅ Resumen generado")
                        st.download_button(
                            label="💾 Descargar Resumen",
                            data=st.session_state['executive_summary'],
                            file_name=f"resumen_{uploaded_file.name.split('.')[0]}.txt",
                            mime="text/plain"
                        )
                    else:
                        st.info("Genera el resumen en la pestaña IA")
        
        else:
            st.error("❌ No se pudo cargar el archivo")
    
    else:
        # Ayuda
        st.markdown("""
        ### 🚀 Cómo usar Excel Processor
        
        1. **Sube tu archivo**: Arrastra o selecciona tu Excel/CSV
        2. **Explora**: Revisa las pestañas de análisis
        3. **Usa IA**: Genera resúmenes y haz preguntas
        
        ### 📋 Formatos soportados
        - Excel: `.xlsx`, `.xls`
        - CSV: `.csv`
        """)

if __name__ == "__main__":
    main()