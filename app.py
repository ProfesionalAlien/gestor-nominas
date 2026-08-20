import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from fpdf import FPDF
import io
import zipfile

# Configuración de la página Web
st.set_page_config(page_title="Gestor de Nóminas Wolmer", page_icon="📄", layout="wide")

st.title("📄 Gestor de Nóminas desde Partes Diarios")
st.markdown("Sube tu archivo **Nominas.xlsx** para calcular las nóminas del mes de forma automática.")

# Subir archivo Excel
uploaded_file = st.file_uploader("Arrastra aquí tu archivo Nominas.xlsx", type=["xlsx"])

def calcular_quinquenios(fecha_alta, fecha_nomina):
    diferencia = relativedelta(fecha_nomina, fecha_alta)
    return diferencia.years // 5

def generar_pdf_nomina(emp, inc, devengos, deducciones, liquido, total_devengado, total_deducciones):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    
    # Encabezado
    pdf.cell(0, 10, "NÓMINA DE TRABAJADOR", ln=True, align="C")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 5, f"Periodo: {inc['Periodo']}", ln=True, align="C")
    pdf.ln(5)
    
    # Datos Empleado
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, f" Trabajador: {emp['Nombre']} | NIF: {emp['NIF']} | SS: {emp['Num:SS']}", ln=True, fill=True)
    pdf.cell(0, 7, f" Categoría: {emp['Categoria']} | Fecha Alta: {str(emp['Fecha_Alta'])[:10]}", ln=True, fill=True)
    pdf.ln(5)
    
    # Tabla Devengos
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(140, 7, "Concepto Devengado", border=1, fill=True)
    pdf.cell(50, 7, "Importe (€)", border=1, ln=True, align="R", fill=True)
    
    pdf.set_font("Helvetica", "", 10)
    for concepto, importe in devengos.items():
        if importe > 0:
            pdf.cell(140, 6, concepto, border=1)
            pdf.cell(50, 6, f"{importe:,.2f} €", border=1, ln=True, align="R")
            
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(140, 7, "TOTAL DEVENGADO", border=1)
    pdf.cell(50, 7, f"{total_devengado:,.2f} €", border=1, ln=True, align="R")
    pdf.ln(5)
    
    # Tabla Deducciones
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(140, 7, "Deducción / Retención", border=1, fill=True)
    pdf.cell(50, 7, "Importe (€)", border=1, ln=True, align="R", fill=True)
    
    pdf.set_font("Helvetica", "", 10)
    for concepto, importe in deducciones.items():
        pdf.cell(140, 6, concepto, border=1)
        pdf.cell(50, 6, f"{importe:,.2f} €", border=1, ln=True, align="R")
        
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(140, 7, "TOTAL DEDUCCIONES", border=1)
    pdf.cell(50, 7, f"{total_deducciones:,.2f} €", border=1, ln=True, align="R")
    pdf.ln(10)
    
    # Líquido
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, f"LÍQUIDO A PERCIBIR: {liquido:,.2f} €", ln=True, align="R")
    
    return bytes(pdf.output())

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_empresa = pd.read_excel(xls, 'Empresa')
    df_empleados = pd.read_excel(xls, 'EmpleadosID_Empleado')
    df_incidencias = pd.read_excel(xls, 'Incidencias_Mes')
    
    st.success("✅ Archivo cargado correctamente.")
    
    st.subheader("📝 Revisión e Incidencias del Mes")
    df_edited = st.data_editor(df_incidencias, num_rows="dynamic")
    
    if st.button("🚀 Calcular Nóminas y Generar PDFs"):
        resultados = []
        archivos_pdf = {}
        
        fecha_nomina = datetime(2026, 7, 31) # Fecha por defecto o extensible
        
        for idx, inc in df_edited.iterrows():
            emp_match = df_empleados[df_empleados['ID_Empleado'] == inc['ID_Empleado']]
            if emp_match.empty:
                continue
            emp = emp_match.iloc[0]
            
            dias_liq = inc.get('Dias_Liquidados', 30)
            
            # 1. Antigüedad (7.5% por quinquenio sobre Salario Base)
            fecha_alta = pd.to_datetime(emp['Fecha_Alta'])
            quinquenios = calcular_quinquenios(fecha_alta, fecha_nomina)
            salario_base = round(emp['Precio_Salario_Base'] * dias_liq, 2)
            antiguedad = round(quinquenios * 0.075 * salario_base, 2)
            
            # 2. Pluses y Partes Diarios
            domingos = round(inc.get('Cant_Trabajo_Domingo', 0) * emp.get('Trabajo_Domingos', 0), 2)
            plus_toxico = round(inc.get('Cant_Plus_Toxico', 0) * emp.get('Precio_Plus_Toxico', 0), 2)
            plus_nocturno = round(inc.get('Cant_Plus_Nocturnidad', 0) * emp.get('Precio_Plus_Nocturnidad', 0), 2)
            comp_toxicidad = round(inc.get('Cant_Comp_Toxicidad', 0) * emp.get('Precio_Comp_Toxicidad', 0), 2)
            plus_transporte = round(inc.get('Cant_Plus_Transporte', 0) * emp.get('Precio_Plus_Transporte', 0), 2)
            comp_descanso = round(inc.get('Cant_Compensacion_Descanso', 0) * emp.get('Compensacion_Descanso_Laborables', 0), 2)
            comp_festivos = round(inc.get('Cant_Compensacion_Descanso_Festivas', 0) * emp.get('Compensacion_Descanso_Festivos', 0), 2)
            gratificacion = round(emp.get('Precio_Gratificacion_Voluntaria', 0) * dias_liq, 2)
            
            devengos = {
                "Salario Base": salario_base,
                "Antigüedad (Quinquenios)": antiguedad,
                "Trabajo en Domingo": domingos,
                "Plus Tóxico": plus_toxico,
                "Plus Nocturnidad": plus_nocturno,
                "Comp. Complemen. Toxicidad": comp_toxicidad,
                "Plus Transporte": plus_transporte,
                "Compensación Descanso Laborables": comp_descanso,
                "Compensación Descanso Festivos": comp_festivos,
                "Gratificación Voluntaria": gratificacion
            }
            
            total_devengado = round(sum(devengos.values()), 2)
            
            # 3. Deducciones
            irpf_pct = emp.get('IRPF_Porcentaje', 0.1719)
            irpf_val = round(total_devengado * irpf_pct, 2)
            
            prorrata = emp.get('Prorrata_Pagas_Extra_Mes', 400.98)
            base_cotizacion = total_devengado + prorrata
            
            cc_val = round(base_cotizacion * 0.0470, 2)
            desempleo_val = round(base_cotizacion * 0.0155, 2)
            fp_val = round(base_cotizacion * 0.0010, 2)
            mei_val = round(base_cotizacion * 0.0015, 2)
            
            deducciones = {
                f"IRPF ({irpf_pct*100:.2f}%)": irpf_val,
                "Contingencias Comunes (4.70%)": cc_val,
                "Desempleo (1.55%)": desempleo_val,
                "Formación Profesional (0.10%)": fp_val,
                "MEI (0.15%)": mei_val
            }
            
            total_deducciones = round(sum(deducciones.values()), 2)
            liquido = round(total_devengado - total_deducciones, 2)
            
            resultados.append({
                "ID": emp['ID_Empleado'],
                "Nombre": emp['Nombre'],
                "Categoría": emp['Categoria'],
                "Quinquenios": quinquenios,
                "Total Devengado": f"{total_devengado:,.2f} €",
                "Total Deducciones": f"{total_deducciones:,.2f} €",
                "Líquido a Percibir": f"{liquido:,.2f} €"
            })
            
            # Generar el PDF
            pdf_bytes = generar_pdf_nomina(emp, inc, devengos, deducciones, liquido, total_devengado, total_deducciones)
            archivos_pdf[f"Nomina_{emp['ID_Empleado']}_{emp['Nombre']}.pdf"] = pdf_bytes
            
        st.subheader("📊 Resumen de Resultados")
        st.dataframe(pd.DataFrame(resultados))
        
        # Generar un archivo ZIP con todos los PDFs
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for nombre_archivo, contenido_pdf in archivos_pdf.items():
                zip_file.writestr(nombre_archivo, contenido_pdf)
                
        st.download_button(
            label="📦 Descargar Todas las Nóminas en ZIP",
            data=zip_buffer.getvalue(),
            file_name="Nominas_Del_Mes.zip",
            mime="application/zip"
        )
