import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import json
from typing import Dict, List, Any, Optional
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from openai import OpenAI
import base64
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
import markdown
from io import BytesIO

# ===== CONFIGURACIÓN DE API KEY =====
# Configura tu API key de OpenAI como variable de entorno
# En tu terminal ejecuta: export OPENAI_API_KEY="tu_api_key_aqui"
# O crea un archivo .env con: OPENAI_API_KEY=tu_api_key_aqui
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")  # Obtener de variable de entorno
# ====================================

# Configuración de la página
st.set_page_config(
    page_title="Excel Processor - Análisis Inteligente",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class ExcelProcessor:
    """
    Procesador de archivos Excel/CSV con análisis de IA
    """
    
    def __init__(self):
        self.client = None
        self.setup_openai()
    
    def setup_openai(self):
        """Configura el cliente de OpenAI"""
        # Prioridad: 1. Hardcodeada, 2. Secrets, 3. Environment
        api_key = OPENAI_API_KEY
        if not api_key or api_key == "sk-proj-xxxxxxxxxxxxxxxxxx":
            api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
        
        if api_key and api_key != "sk-proj-xxxxxxxxxxxxxxxxxx":
            try:
                self.client = OpenAI(api_key=api_key)
                # Test the connection
                self.client.models.list()
                return True
            except Exception as e:
                st.error(f"❌ Error configurando AskNOA Processor: {str(e)}")
                return False
        else:
            st.warning("⚠️ Configura tu API key en la línea 16 del código para habilitar AskNOA Processor.")
            return False
    
    def clean_dataframe_for_display(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Limpia el DataFrame para evitar errores de serialización en Streamlit
        """
        df_clean = df.copy()
        
        # Convertir todas las columnas a string para evitar problemas de tipo
        for col in df_clean.columns:
            # Si la columna tiene tipos mixtos, convertir a string
            if df_clean[col].dtype == 'object':
                df_clean[col] = df_clean[col].astype(str)
            # Reemplazar valores problemáticos
            df_clean[col] = df_clean[col].replace([np.inf, -np.inf], 'inf')
            df_clean[col] = df_clean[col].fillna('N/A')
        
        return df_clean
    
    def load_file(self, uploaded_file) -> Dict[str, pd.DataFrame]:
        """
        Carga un archivo Excel/CSV y retorna un diccionario con todas las hojas
        """
        try:
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            if file_extension == 'csv':
                # Para CSV, intentar diferentes encodings
                try:
                    df = pd.read_csv(uploaded_file, encoding='utf-8')
                except UnicodeDecodeError:
                    try:
                        uploaded_file.seek(0)  # Reset file pointer
                        df = pd.read_csv(uploaded_file, encoding='latin-1')
                    except UnicodeDecodeError:
                        uploaded_file.seek(0)  # Reset file pointer
                        df = pd.read_csv(uploaded_file, encoding='cp1252')
                
                return {"Sheet1": df}
            
            elif file_extension in ['xlsx', 'xls']:
                # Para Excel, leer todas las hojas con manejo de errores mejorado
                try:
                    excel_file = pd.ExcelFile(uploaded_file)
                    sheets = {}
                    for sheet_name in excel_file.sheet_names:
                        try:
                            # Leer la hoja con parámetros adicionales para evitar problemas
                            df = pd.read_excel(
                                uploaded_file, 
                                sheet_name=sheet_name,
                                na_values=['', 'NA', 'N/A', 'null', 'NULL'],
                                keep_default_na=True
                            )
                            sheets[sheet_name] = df
                        except Exception as sheet_error:
                            st.warning(f"⚠️ Error leyendo la hoja '{sheet_name}': {sheet_error}")
                            continue
                    return sheets
                except Exception as e:
                    st.error(f"Error leyendo el archivo Excel: {str(e)}")
                    return {}
            
            else:
                st.error(f"Formato de archivo no soportado: {file_extension}")
                return {}
                
        except Exception as e:
            st.error(f"Error al cargar el archivo: {str(e)}")
            return {}
    
    def analyze_dataframe(self, df: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
        """
        Analiza un DataFrame y extrae información relevante
        """
        try:
            # Crear una copia para análisis
            df_analysis = df.copy()
            
            analysis = {
                "sheet_name": sheet_name,
                "shape": df_analysis.shape,
                "columns": list(df_analysis.columns),
                "dtypes": {str(col): str(dtype) for col, dtype in df_analysis.dtypes.items()},
                "null_counts": df_analysis.isnull().sum().to_dict(),
                "memory_usage": df_analysis.memory_usage(deep=True).sum(),
                "numeric_columns": df_analysis.select_dtypes(include=[np.number]).columns.tolist(),
                "categorical_columns": df_analysis.select_dtypes(include=['object']).columns.tolist(),
                "datetime_columns": df_analysis.select_dtypes(include=['datetime64']).columns.tolist(),
            }
            
            # Muestra de datos segura
            try:
                sample_df = df_analysis.head(5)
                # Convertir a strings para evitar problemas de serialización
                sample_dict = []
                for idx, row in sample_df.iterrows():
                    row_dict = {}
                    for col in sample_df.columns:
                        try:
                            val = row[col]
                            if pd.isna(val):
                                row_dict[str(col)] = "N/A"
                            elif isinstance(val, (int, float)) and (np.isinf(val) or np.isnan(val)):
                                row_dict[str(col)] = "inf" if np.isinf(val) else "N/A"
                            else:
                                row_dict[str(col)] = str(val)
                        except:
                            row_dict[str(col)] = "Error"
                    sample_dict.append(row_dict)
                analysis["sample_data"] = sample_dict
            except Exception as e:
                analysis["sample_data"] = [{"error": f"Error generando muestra: {str(e)}"}]
            
            # Estadísticas descriptivas para columnas numéricas
            if analysis["numeric_columns"]:
                try:
                    numeric_stats = {}
                    for col in analysis["numeric_columns"]:
                        try:
                            stats = df_analysis[col].describe()
                            numeric_stats[col] = {
                                'count': float(stats['count']) if not pd.isna(stats['count']) else 0,
                                'mean': float(stats['mean']) if not pd.isna(stats['mean']) else 0,
                                'std': float(stats['std']) if not pd.isna(stats['std']) else 0,
                                'min': float(stats['min']) if not pd.isna(stats['min']) else 0,
                                '25%': float(stats['25%']) if not pd.isna(stats['25%']) else 0,
                                '50%': float(stats['50%']) if not pd.isna(stats['50%']) else 0,
                                '75%': float(stats['75%']) if not pd.isna(stats['75%']) else 0,
                                'max': float(stats['max']) if not pd.isna(stats['max']) else 0,
                            }
                        except Exception as col_error:
                            numeric_stats[col] = {"error": f"Error calculando estadísticas: {str(col_error)}"}
                    analysis["numeric_stats"] = numeric_stats
                except Exception as e:
                    analysis["numeric_stats"] = {"error": f"Error en estadísticas numéricas: {str(e)}"}
            
            # Valores únicos para columnas categóricas (limitado a 20)
            categorical_info = {}
            for col in analysis["categorical_columns"]:
                try:
                    # Filtrar valores nulos antes de obtener únicos
                    non_null_values = df_analysis[col].dropna()
                    unique_values = non_null_values.unique()
                    
                    categorical_info[col] = {
                        "unique_count": len(unique_values),
                        "unique_values": [str(val) for val in unique_values[:20]],
                        "top_values": {str(k): int(v) for k, v in non_null_values.value_counts().head(10).items()}
                    }
                except Exception as col_error:
                    categorical_info[col] = {
                        "unique_count": 0,
                        "unique_values": [],
                        "top_values": {},
                        "error": f"Error procesando columna: {str(col_error)}"
                    }
            
            analysis["categorical_info"] = categorical_info
            
            return analysis
            
        except Exception as e:
            st.error(f"Error en el análisis del DataFrame: {str(e)}")
            return {
                "sheet_name": sheet_name,
                "shape": (0, 0),
                "columns": [],
                "error": str(e)
            }

    def generate_context_description(self, analysis: Dict[str, Any]) -> str:
        """
        Genera una descripción contextual del análisis para la IA
        """
        if "error" in analysis:
            return f"Error en el análisis: {analysis['error']}"
            
        context = f"""
# Análisis del archivo Excel/CSV

## Información general:
- **Hoja**: {analysis['sheet_name']}
- **Dimensiones**: {analysis['shape'][0]} filas × {analysis['shape'][1]} columnas
- **Uso de memoria**: {analysis.get('memory_usage', 0):,} bytes

## Estructura de columnas:
### Columnas numéricas ({len(analysis.get('numeric_columns', []))}):
{', '.join(analysis.get('numeric_columns', [])) if analysis.get('numeric_columns') else 'Ninguna'}

### Columnas categóricas ({len(analysis.get('categorical_columns', []))}):
{', '.join(analysis.get('categorical_columns', [])) if analysis.get('categorical_columns') else 'Ninguna'}

### Columnas de fecha ({len(analysis.get('datetime_columns', []))}):
{', '.join(analysis.get('datetime_columns', [])) if analysis.get('datetime_columns') else 'Ninguna'}

## Descripción detallada de columnas:
"""
        
        # Agregar información detallada de cada columna
        for col in analysis.get('columns', []):
            dtype = analysis.get('dtypes', {}).get(col, 'unknown')
            null_count = analysis.get('null_counts', {}).get(col, 0)
            null_percentage = (null_count / analysis['shape'][0]) * 100 if analysis['shape'][0] > 0 else 0
            
            context += f"\n### {col}:\n"
            context += f"- **Tipo de dato**: {dtype}\n"
            context += f"- **Valores nulos**: {null_count} ({null_percentage:.1f}%)\n"
            
            # Información específica según el tipo de columna
            if col in analysis.get('numeric_columns', []) and 'numeric_stats' in analysis:
                stats = analysis['numeric_stats'].get(col, {})
                if 'error' not in stats:
                    context += f"- **Estadísticas**: Min: {stats.get('min', 0)}, Max: {stats.get('max', 0)}, Media: {stats.get('mean', 0):.2f}, Mediana: {stats.get('50%', 0):.2f}\n"
            
            elif col in analysis.get('categorical_columns', []):
                cat_info = analysis.get('categorical_info', {}).get(col, {})
                if 'error' not in cat_info:
                    context += f"- **Valores únicos**: {cat_info.get('unique_count', 0)}\n"
                    top_values = list(cat_info.get('top_values', {}).keys())[:5]
                    context += f"- **Valores más frecuentes**: {top_values}\n"
        
        # Agregar muestra de datos
        context += f"\n## Muestra de datos (primeras 5 filas):\n"
        sample_data = analysis.get('sample_data', [])
        for i, row in enumerate(sample_data[:5]):
            if 'error' not in row:
                context += f"\n**Fila {i+1}**: {row}\n"
        
        return context
    
    async def analyze_with_ai(self, context: str, user_query: str = None) -> str:
        """
        Usa Code Interpreter de OpenAI para analizar el contexto del archivo
        """
        if not self.client:
            return "❌ No se puede realizar análisis con IA: API key no configurada"
        
        try:
            system_prompt = """
Eres un analista de datos experto especializado en Excel y CSV. Tu tarea es analizar archivos de datos y proporcionar insights completos y detallados.

Para cada archivo que analices, debes:

1. **Resumen General**: Describe qué tipo de datos contiene el archivo y su propósito aparente
2. **Análisis de Columnas**: Explica cada columna, su significado y relevancia
3. **Calidad de Datos**: Identifica problemas de calidad (valores nulos, inconsistencias, etc.)
4. **Patrones y Tendencias**: Identifica patrones interesantes en los datos
5. **Recomendaciones**: Sugiere análisis adicionales o mejoras en los datos
6. **Casos de Uso**: Propón posibles aplicaciones o análisis que se pueden hacer con estos datos

Sé detallado, preciso y proporciona insights valiosos que permitan a una IA posterior trabajar efectivamente con estos datos.
"""
            
            user_message = f"""
Analiza el siguiente archivo de datos y proporciona un análisis completo:

{context}

{f"Pregunta específica del usuario: {user_query}" if user_query else ""}

Por favor, proporciona un análisis detallado siguiendo la estructura solicitada.
"""
            
            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_message}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"❌ Error en el análisis con IA: {str(e)}"
    
    def create_visualizations(self, df: pd.DataFrame, analysis: Dict[str, Any]) -> List[go.Figure]:
        """
        Crea visualizaciones automáticas basadas en el análisis
        """
        figures = []
        
        try:
            # Gráfico de valores nulos
            null_counts = analysis.get('null_counts', {})
            if any(null_counts.values()):
                null_data = {k: v for k, v in null_counts.items() if v > 0}
                if null_data:
                    fig = px.bar(
                        x=list(null_data.keys()),
                        y=list(null_data.values()),
                        title="Valores Nulos por Columna",
                        labels={'x': 'Columnas', 'y': 'Cantidad de Valores Nulos'}
                    )
                    figures.append(fig)
            
            # Histogramas para columnas numéricas
            numeric_columns = analysis.get('numeric_columns', [])
            for col in numeric_columns[:3]:  # Máximo 3 gráficos
                try:
                    # Filtrar valores no finitos
                    col_data = df[col].replace([np.inf, -np.inf], np.nan).dropna()
                    if len(col_data) > 0:
                        fig = px.histogram(
                            x=col_data,
                            title=f"Distribución de {col}",
                            nbins=30
                        )
                        figures.append(fig)
                except Exception as e:
                    continue
            
            # Gráficos de barras para columnas categóricas
            categorical_columns = analysis.get('categorical_columns', [])
            for col in categorical_columns[:2]:  # Máximo 2 gráficos
                try:
                    value_counts = df[col].value_counts().head(10)
                    if len(value_counts) > 0:
                        fig = px.bar(
                            x=value_counts.index,
                            y=value_counts.values,
                            title=f"Top 10 valores en {col}",
                            labels={'x': col, 'y': 'Frecuencia'}
                        )
                        figures.append(fig)
                except Exception as e:
                    continue
        
        except Exception as e:
            st.warning(f"⚠️ Error generando algunas visualizaciones: {str(e)}")
        
        return figures

    async def generate_executive_summary(self, sheets: Dict[str, pd.DataFrame], filename: str) -> str:
        """
        Genera un resumen ejecutivo completo del archivo usando IA
        """
        if not self.client:
            return "❌ No se puede generar resumen ejecutivo: API key no configurada"
        
        try:
            # Construir contexto completo de todas las hojas
            full_context = f"""
# ANÁLISIS COMPLETO DEL ARCHIVO: {filename}

## ESTRUCTURA GENERAL DEL ARCHIVO:
- **Nombre del archivo**: {filename}
- **Número total de hojas**: {len(sheets)}
- **Hojas disponibles**: {', '.join(sheets.keys())}
- **Fecha de análisis**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

"""
            
            # Analizar cada hoja individualmente
            for sheet_name, df in sheets.items():
                analysis = self.analyze_dataframe(df, sheet_name)
                sheet_context = self.generate_context_description(analysis)
                full_context += f"\n{'='*80}\n"
                full_context += sheet_context
                full_context += f"\n{'='*80}\n"
            
            # Prompt especializado para resumen ejecutivo
            executive_prompt = """
Eres un analista de datos senior especializado en crear resúmenes ejecutivos detallados. Tu tarea es analizar archivos Excel/CSV completos y generar un resumen ejecutivo que permita a cualquier IA posterior entender completamente el contenido y propósito del archivo.

DEBES GENERAR UN RESUMEN EJECUTIVO QUE INCLUYA:

## 1. 📋 RESUMEN GENERAL
- ¿Qué tipo de documento es este?
- ¿Cuál es su propósito principal?
- ¿Qué industria o área de negocio representa?
- ¿Qué período de tiempo cubre?

## 2. 📊 ESTRUCTURA DEL ARCHIVO
- Descripción de cada hoja y su función
- Relación entre las hojas (si existe)
- Jerarquía o flujo de información

## 3. 🔍 ANÁLISIS DETALLADO DE COLUMNAS POR HOJA
Para cada hoja, explica:
- Qué representa cada columna
- Tipo de datos y formato
- Significado de negocio de cada campo
- Relaciones entre columnas
- Columnas clave o identificadores

## 4. 📈 PATRONES Y INSIGHTS IDENTIFICADOS
- Tendencias importantes en los datos
- Anomalías o problemas de calidad detectados
- Métricas clave del negocio
- Rangos de valores esperados

## 5. 🎯 CASOS DE USO Y APLICACIONES
- ¿Para qué se puede usar este archivo?
- ¿Qué análisis se pueden realizar?
- ¿Qué preguntas de negocio puede responder?
- ¿Qué reportes se pueden generar?

## 6. 🤖 CONTEXTO PARA IA
Proporciona un contexto claro y completo que permita a una IA:
- Entender perfectamente el contenido
- Responder preguntas específicas sobre los datos
- Realizar análisis y cálculos apropiados
- Generar insights relevantes

IMPORTANTE: Sé específico, detallado y profesional. Este resumen será usado por otras IAs para entender y trabajar con estos datos.
"""
            
            user_message = f"""
Por favor, analiza el siguiente archivo completo y genera un resumen ejecutivo detallado:

{full_context}

Genera un resumen ejecutivo profesional y completo siguiendo la estructura solicitada.
"""
            
            # Generar resumen con modelo más potente si está disponible
            try:
                response = self.client.chat.completions.create(
                    model="gpt-4o",  # Usar modelo más potente para mejor análisis
                    messages=[
                        {"role": "system", "content": executive_prompt},
                        {"role": "user", "content": user_message}
                    ],
                    temperature=0.2,
                    max_tokens=4000   # Más tokens para resumen completo
                )
            except:
                # Fallback a modelo más básico si gpt-4o no está disponible
                response = self.client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": executive_prompt},
                        {"role": "user", "content": user_message}
                    ],
                    temperature=0.2,
                    max_tokens=3000
                )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"❌ Error generando resumen ejecutivo: {str(e)}"

    async def generate_comprehensive_executive_summary(self, sheets: Dict[str, pd.DataFrame], filename: str) -> str:
        """
        Genera un resumen ejecutivo extensivo y detallado usando gpt-4.1-mini
        """
        if not self.client:
            return "❌ No se puede generar resumen ejecutivo: AskNOA Processor no configurado"
        
        try:
            # Construir contexto ultra-detallado de todas las hojas
            full_context = f"""
# ANÁLISIS EXHAUSTIVO DEL ARCHIVO: {filename}

## METADATOS DEL ARCHIVO:
- **Nombre completo**: {filename}
- **Extensión**: {filename.split('.')[-1].upper() if '.' in filename else 'Desconocida'}
- **Total de hojas/pestañas**: {len(sheets)}
- **Nombres de hojas**: {', '.join(sheets.keys())}
- **Fecha y hora de análisis**: {datetime.now().strftime('%d/%m/%Y a las %H:%M:%S')}
- **Tamaño total estimado**: {sum(df.memory_usage(deep=True).sum() for df in sheets.values()) / 1024:.1f} KB

## ANÁLISIS DETALLADO POR HOJA:
"""
            
            # Análisis exhaustivo hoja por hoja
            for sheet_name, df in sheets.items():
                analysis = self.analyze_dataframe(df, sheet_name)
                
                full_context += f"""

{'='*100}
### HOJA: "{sheet_name}"
{'='*100}

#### ESTRUCTURA GENERAL:
- **Dimensiones**: {analysis['shape'][0]:,} filas × {analysis['shape'][1]} columnas
- **Densidad de datos**: {((analysis['shape'][0] * analysis['shape'][1] - sum(analysis.get('null_counts', {}).values())) / (analysis['shape'][0] * analysis['shape'][1]) * 100):.1f}% (datos no nulos)
- **Memoria utilizada**: {analysis.get('memory_usage', 0) / 1024:.1f} KB

#### CLASIFICACIÓN DE COLUMNAS:
- **Numéricas**: {len(analysis.get('numeric_columns', []))} columnas → {analysis.get('numeric_columns', [])}
- **Categóricas**: {len(analysis.get('categorical_columns', []))} columnas → {analysis.get('categorical_columns', [])}
- **Fechas**: {len(analysis.get('datetime_columns', []))} columnas → {analysis.get('datetime_columns', [])}

#### ANÁLISIS DETALLADO DE CADA COLUMNA:
"""
                
                # Análisis columna por columna MUY detallado
                for i, col in enumerate(analysis.get('columns', []), 1):
                    dtype = analysis.get('dtypes', {}).get(col, 'unknown')
                    null_count = analysis.get('null_counts', {}).get(col, 0)
                    null_percentage = (null_count / analysis['shape'][0]) * 100 if analysis['shape'][0] > 0 else 0
                    
                    full_context += f"""
**COLUMNA #{i}: "{col}"**
├─ Tipo de dato: {dtype}
├─ Posición: Columna {i} de {len(analysis.get('columns', []))}
├─ Valores nulos: {null_count:,} ({null_percentage:.2f}%)
├─ Valores no nulos: {analysis['shape'][0] - null_count:,} ({100-null_percentage:.2f}%)
"""
                    
                    # Análisis específico para columnas numéricas
                    if col in analysis.get('numeric_columns', []) and 'numeric_stats' in analysis:
                        stats = analysis['numeric_stats'].get(col, {})
                        if 'error' not in stats:
                            full_context += f"""├─ ESTADÍSTICAS NUMÉRICAS:
│  ├─ Valor mínimo: {stats.get('min', 0):,.2f}
│  ├─ Valor máximo: {stats.get('max', 0):,.2f}
│  ├─ Promedio: {stats.get('mean', 0):,.2f}
│  ├─ Mediana (Q2): {stats.get('50%', 0):,.2f}
│  ├─ Cuartil 1 (Q1): {stats.get('25%', 0):,.2f}
│  ├─ Cuartil 3 (Q3): {stats.get('75%', 0):,.2f}
│  ├─ Desviación estándar: {stats.get('std', 0):,.2f}
│  ├─ Rango: {stats.get('max', 0) - stats.get('min', 0):,.2f}
│  └─ Registros válidos: {int(stats.get('count', 0)):,}
"""
                    
                    # Análisis específico para columnas categóricas
                    elif col in analysis.get('categorical_columns', []):
                        cat_info = analysis.get('categorical_info', {}).get(col, {})
                        if 'error' not in cat_info:
                            full_context += f"""├─ ANÁLISIS CATEGÓRICO:
│  ├─ Valores únicos: {cat_info.get('unique_count', 0)} 
│  ├─ Diversidad: {(cat_info.get('unique_count', 0)/analysis['shape'][0]*100):.1f}% (únicos/total)
│  └─ TOP 5 valores más frecuentes:
"""
                            for rank, (value, count) in enumerate(list(cat_info.get('top_values', {}).items())[:5], 1):
                                percentage = (count / analysis['shape'][0]) * 100
                                full_context += f"     {rank}. '{value}': {count:,} ocurrencias ({percentage:.1f}%)\n"
                
                # Muestra representativa de datos
                full_context += f"""
#### MUESTRA REPRESENTATIVA DE DATOS (Primeras 5 filas):
"""
                sample_data = analysis.get('sample_data', [])[:5]
                for i, row in enumerate(sample_data):
                    if 'error' not in row:
                        full_context += f"\n**Registro #{i+1}:** {row}\n"
                
                full_context += f"\n{'='*100}\n"
            
            # Prompt ultra-especializado para análisis exhaustivo
            comprehensive_prompt = """
Eres un analista de datos SENIOR especializado en análisis exhaustivos de archivos empresariales. Tu tarea es crear un RESUMEN EJECUTIVO EXTENSIVO Y PROFESIONAL que permita a cualquier stakeholder o IA posterior entender completamente el archivo y trabajar con él de manera efectiva.

DEBES GENERAR UN ANÁLISIS COMPLETO QUE INCLUYA:

## 📋 1. RESUMEN EJECUTIVO GENERAL
- Identifica QUÉ TIPO de documento empresarial es (ventas, inventario, RRHH, financiero, etc.)
- Determina el PROPÓSITO PRINCIPAL y contexto empresarial
- Identifica la INDUSTRIA o SECTOR específico
- Establece el PERÍODO TEMPORAL que cubre los datos
- Evalúa la CALIDAD y COMPLETITUD general de los datos

## 🏗️ 2. ARQUITECTURA DE INFORMACIÓN
Para cada hoja/pestaña:
- **Función específica** de cada hoja en el contexto empresarial
- **Relaciones jerárquicas** entre hojas (si existen)
- **Flujo de información** y dependencias
- **Importancia relativa** de cada hoja para el negocio

## 🔍 3. ANÁLISIS GRANULAR DE CAMPOS
Para CADA COLUMNA de CADA HOJA:
- **Significado empresarial** específico del campo
- **Tipo de información** que contiene (ID, métrica, dimensión, etc.)
- **Criticidad** para el negocio (crítico, importante, auxiliar)
- **Calidad del dato** (completitud, consistencia, anomalías)
- **Relaciones** con otros campos del mismo archivo
- **Restricciones** o reglas de negocio evidentes

## 📊 4. INSIGHTS Y PATRONES IDENTIFICADOS
- **Tendencias** detectadas en los datos numéricos
- **Anomalías** o valores atípicos significativos
- **Problemas de calidad** críticos que requieren atención
- **Métricas clave** del negocio identificadas
- **Oportunidades de mejora** en los procesos de datos

## 🎯 5. APLICACIONES EMPRESARIALES
- **Análisis** específicos que se pueden realizar
- **Reportes** ejecutivos que se pueden generar
- **KPIs** que se pueden calcular
- **Decisiones empresariales** que estos datos pueden informar
- **Procesos** que estos datos pueden optimizar

## 🤖 6. CONTEXTO TÉCNICO PARA IA
Proporciona instrucciones ESPECÍFICAS para que una IA pueda:
- **Interpretar correctamente** cada campo
- **Realizar cálculos** apropiados y relevantes
- **Generar análisis** de valor empresarial
- **Crear visualizaciones** significativas
- **Responder preguntas** específicas del negocio
- **Identificar** oportunidades y riesgos

## 📈 7. RECOMENDACIONES ESTRATÉGICAS
- **Mejoras** en la estructura de datos
- **Procesos** de limpieza y validación recomendados
- **Análisis adicionales** de alto valor
- **Automatizaciones** posibles
- **Integraciones** con otras fuentes de datos

IMPORTANTE: 
- Sé EXTREMADAMENTE DETALLADO y específico
- Usa terminología empresarial apropiada
- Proporciona ejemplos concretos cuando sea relevante
- Enfócate en el VALOR EMPRESARIAL de cada elemento
- Este análisis será usado por ejecutivos y sistemas de IA para tomar decisiones
"""
            
            user_message = f"""
Analiza exhaustivamente el siguiente archivo empresarial y genera un resumen ejecutivo extensivo y profesional:

{full_context}

GENERA UN ANÁLISIS COMPLETO Y DETALLADO siguiendo la estructura solicitada. Este documento será utilizado por la alta dirección y sistemas de IA para tomar decisiones estratégicas.
"""
            
            # Usar gpt-4.1-mini específicamente para análisis extensivo
            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {"role": "system", "content": comprehensive_prompt},
                    {"role": "user", "content": user_message}
                ],
                temperature=0.1,  # Muy baja para máxima precisión y consistencia
                max_tokens=8000   # Máximo permitido para análisis extensivo
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"❌ Error generando resumen ejecutivo extensivo: {str(e)}"

    def generate_professional_pdf(self, content: str, filename: str) -> bytes:
        """
        Genera un PDF profesional a partir del contenido markdown
        """
        try:
            buffer = BytesIO()
            
            # Crear documento PDF
            doc = SimpleDocTemplate(
                buffer,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=72
            )
            
            # Definir estilos profesionales
            styles = getSampleStyleSheet()
            
            # Estilo para título principal
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Title'],
                fontSize=20,
                spaceAfter=30,
                alignment=TA_CENTER,
                textColor=colors.HexColor('#1f4e79')
            )
            
            # Estilo para subtítulos
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading1'],
                fontSize=14,
                spaceAfter=12,
                spaceBefore=20,
                textColor=colors.HexColor('#2f5f8f')
            )
            
            # Estilo para texto normal
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=10,
                spaceAfter=6,
                alignment=TA_JUSTIFY
            )
            
            # Construir contenido del PDF
            story = []
            
            # Título principal
            story.append(Paragraph(f"RESUMEN EJECUTIVO EXTENSIVO", title_style))
            story.append(Paragraph(f"Archivo: {filename}", normal_style))
            story.append(Paragraph(f"Fecha: {datetime.now().strftime('%d de %B de %Y')}", normal_style))
            story.append(Spacer(1, 30))
            
            # Procesar contenido markdown
            lines = content.split('\n')
            current_paragraph = ""
            
            for line in lines:
                line = line.strip()
                
                if line.startswith('# '):
                    # Título principal
                    if current_paragraph:
                        story.append(Paragraph(current_paragraph, normal_style))
                        current_paragraph = ""
                    story.append(Spacer(1, 20))
                    story.append(Paragraph(line[2:], title_style))
                    
                elif line.startswith('## '):
                    # Subtítulo
                    if current_paragraph:
                        story.append(Paragraph(current_paragraph, normal_style))
                        current_paragraph = ""
                    story.append(Spacer(1, 15))
                    story.append(Paragraph(line[3:], heading_style))
                    
                elif line.startswith('### '):
                    # Subtítulo menor
                    if current_paragraph:
                        story.append(Paragraph(current_paragraph, normal_style))
                        current_paragraph = ""
                    story.append(Spacer(1, 10))
                    sub_heading_style = ParagraphStyle(
                        'SubHeading',
                        parent=styles['Heading2'],
                        fontSize=12,
                        spaceAfter=6,
                        spaceBefore=10,
                        textColor=colors.HexColor('#4f7faf')
                    )
                    story.append(Paragraph(line[4:], sub_heading_style))
                    
                elif line.startswith('- ') or line.startswith('* '):
                    # Lista
                    if current_paragraph:
                        story.append(Paragraph(current_paragraph, normal_style))
                        current_paragraph = ""
                    bullet_style = ParagraphStyle(
                        'Bullet',
                        parent=styles['Normal'],
                        fontSize=10,
                        leftIndent=20,
                        bulletIndent=10,
                        spaceAfter=3
                    )
                    story.append(Paragraph(f"• {line[2:]}", bullet_style))
                    
                elif line.startswith('**') and line.endswith('**'):
                    # Texto en negrita
                    if current_paragraph:
                        story.append(Paragraph(current_paragraph, normal_style))
                        current_paragraph = ""
                    bold_style = ParagraphStyle(
                        'Bold',
                        parent=styles['Normal'],
                        fontSize=10,
                        spaceAfter=6
                    )
                    story.append(Paragraph(f"<b>{line[2:-2]}</b>", bold_style))
                    
                elif line:
                    # Texto normal
                    if current_paragraph:
                        current_paragraph += " " + line
                    else:
                        current_paragraph = line
                else:
                    # Línea vacía
                    if current_paragraph:
                        story.append(Paragraph(current_paragraph, normal_style))
                        current_paragraph = ""
                    story.append(Spacer(1, 6))
            
            # Agregar último párrafo si existe
            if current_paragraph:
                story.append(Paragraph(current_paragraph, normal_style))
            
            # Pie de página
            story.append(Spacer(1, 30))
            footer_style = ParagraphStyle(
                'Footer',
                parent=styles['Normal'],
                fontSize=8,
                alignment=TA_CENTER,
                textColor=colors.grey
            )
            story.append(Paragraph("Generado por AskNOA Processor - Excel Analysis System", footer_style))
            
            # Construir PDF
            doc.build(story)
            
            buffer.seek(0)
            return buffer.getvalue()
            
        except Exception as e:
            st.error(f"Error generando PDF: {str(e)}")
            return b""

    def clean_and_analyze_file(self, uploaded_file) -> Dict[str, Any]:
        """
        Limpia y analiza un archivo Excel/CSV cargado, y genera un resumen ejecutivo
        """
        try:
            # Cargar archivo
            sheets = self.load_file(uploaded_file)
            
            if not sheets:
                return {"error": "No se pudieron cargar hojas del archivo."}
            
            # Análisis de cada hoja
            analysis_results = {}
            for sheet_name, df in sheets.items():
                analysis = self.analyze_dataframe(df, sheet_name)
                analysis_results[sheet_name] = analysis
            
            # Generar resumen ejecutivo
            summary = self.generate_context_description(analysis_results[list(analysis_results.keys())[0]])
            
            return {
                "sheets": analysis_results,
                "executive_summary": summary
            }
        
        except Exception as e:
            return {"error": str(e)}

    async def process_and_analyze_file(self, uploaded_file) -> Dict[str, Any]:
        """
        Procesa y analiza un archivo Excel/CSV cargado, y genera un resumen ejecutivo de forma asíncrona
        """
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(None, self.clean_and_analyze_file, uploaded_file)
        return result

    async def handle_file_upload(self, uploaded_file) -> None:
        """
        Maneja la carga y procesamiento de un archivo, y actualiza la interfaz de usuario
        """
        if not uploaded_file:
            return
        
        with st.spinner("🔄 Procesando archivo..."):
            result = await self.process_and_analyze_file(uploaded_file)
        
        if "error" in result:
            st.error(f"❌ Error: {result['error']}")
            return
        
        # Desplegar resultados
        sheets = result.get("sheets", {})
        executive_summary = result.get("executive_summary", "")
        
        st.success(f"✅ Archivo procesado. Se encontraron {len(sheets)} hoja(s)")
        
        # Mostrar resumen ejecutivo
        st.subheader("📋 Resumen Ejecutivo")
        st.markdown(executive_summary)
        
        # Mostrar análisis de cada hoja
        for sheet_name, analysis in sheets.items():
            st.subheader(f"📊 Análisis de la hoja: {sheet_name}")
            st.write(analysis)
        
        # Opción de descarga del resumen ejecutivo en PDF
        if st.button("📥 Descargar Resumen Ejecutivo en PDF"):
            try:
                pdf_data = self.generate_professional_pdf(executive_summary, uploaded_file.name)
                b64 = base64.b64encode(pdf_data).decode('utf-8')
                href = f"data:application/pdf;base64,{b64}"
                st.markdown(f"**Resumen Ejecutivo PDF**: [Descargar aquí]({href})", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error generando PDF: {str(e)}")

        # Opción de descarga del análisis en JSON
        if st.button("📥 Descargar Análisis en JSON"):
            try:
                json_data = json.dumps(sheets, indent=2, default=str)
                b64 = base64.b64encode(json_data.encode()).decode()
                href = f"data:file/json;base64,{b64}"
                st.markdown(f"**Análisis JSON**: [Descargar aquí]({href})", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error generando JSON: {str(e)}")

        # Opción de descarga de datos procesados en CSV
        if st.button("📥 Descargar Datos Procesados en CSV"):
            try:
                csv_data = ""
                for sheet_name, df in sheets.items():
                    csv_data += f"\n\n# {sheet_name}\n"
                    csv_data += df.to_csv(index=False)
                
                b64 = base64.b64encode(csv_data.encode()).decode()
                href = f"data:file/csv;base64,{b64}"
                st.markdown(f"**Datos Procesados CSV**: [Descargar aquí]({href})", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error generando CSV: {str(e)}")

    # Footer
    st.markdown("---")
    st.markdown("*Excel Processor - Análisis Inteligente con IA*")

if __name__ == "__main__":
    main()