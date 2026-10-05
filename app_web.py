from flask import Flask, render_template_string, request, redirect, url_for, flash
import pandas as pd
import os
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler

app = Flask(__name__)
app.secret_key = "clave_secreta_super_segura"

EXCEL_DATA = "CONSOLIDADO EPC.xlsx"
ASISTENCIA_FILE = "registro_asistencias.csv"

def cargar_datos_base():
    try:
        if not os.path.exists(EXCEL_DATA):
            return pd.DataFrame()
        df = pd.read_excel(EXCEL_DATA, sheet_name="EQUIPO POLITICO", header=1)
        df.columns = [str(c).strip() for c in df.columns]
        df["CEDULA"] = df["CEDULA"].astype(str).str.split(".").str[0].str.strip()
        return df
    except Exception as e:
        print(f"Error cargando Excel base: {e}")
        return pd.DataFrame()

# Tarea automática a las 11:59 PM
def tarea_medianoche():
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    if not os.path.exists(ASISTENCIA_FILE):
        return
    df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
    df_filtrado = df_asist[df_asist["FECHA"] == fecha_hoy]
    if not df_filtrado.empty:
        nombre_archivo = f"Cierre_Asistencia_{fecha_hoy.replace('-', '_')}.xlsx"
        try:
            df_filtrado.to_excel(nombre_archivo, index=False)
            print(f"[AUTOMÁTICO] Asistencia guardada en {nombre_archivo}")
        except Exception as e:
            print(f"[ERROR AUTOMÁTICO] {e}")

scheduler = BackgroundScheduler()
scheduler.add_job(func=tarea_medianoche, trigger="cron", hour=23, minute=59)
scheduler.start()

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es" data-bs-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Control Diario de Asistencia</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
</head>
<body class="bg-body-tertiary">
    <nav class="navbar navbar-expand-lg border-bottom shadow-sm bg-body">
        <div class="container-fluid px-4">
            <a class="navbar-brand fw-bold" href="/"><i class="fa-solid fa-clipboard-user text-success me-2"></i>Control de Asistencia EPC</a>
            <div class="d-flex align-items-center">
                <button class="btn btn-outline-secondary btn-sm me-3" id="btn-theme-toggle" onclick="toggleTheme()" title="Cambiar modo claro/oscuro">
                    <i class="fa-solid fa-moon" id="theme-icon"></i>
                </button>
                <span class="navbar-text small text-muted"><i class="fa-regular fa-calendar-days me-1"></i>{{ fecha_hoy }}</span>
            </div>
        </div>
    </nav>

    <div class="container py-4">
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ 'success' if category == 'success' else ('danger' if category == 'danger' else 'warning') }} alert-dismissible fade show shadow-sm" role="alert">
                        <i class="fa-solid {{ 'fa-circle-check' if category == 'success' else ('fa-triangle-exclamation' if category == 'danger' else 'fa-circle-info') }} me-2"></i>
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <div class="row g-4">
            <!-- Formulario de Registro -->
            <div class="col-lg-4">
                <div class="card shadow-sm border-0 mb-4">
                    <div class="card-header bg-success text-white py-3">
                        <h5 class="card-title mb-0 fs-6 fw-bold"><i class="fa-solid fa-user-check me-2"></i>Registrar por Cédula</h5>
                    </div>
                    <div class="card-body p-4">
                        <form action="/registrar" method="POST">
                            <div class="mb-3">
                                <label for="cedula" class="form-label fw-semibold">Número de Cédula:</label>
                                <div class="input-group">
                                    <span class="input-group-text bg-light text-secondary"><i class="fa-solid fa-id-card"></i></span>
                                    <input type="text" class="form-control form-control-lg" id="cedula" name="cedula" placeholder="Ej. 12345678" required autofocus>
                                </div>
                            </div>
                            <button type="submit" class="btn btn-success w-100 py-2 fw-bold shadow-sm">
                                <i class="fa-solid fa-paper-plane me-2"></i>Registrar Asistencia
                            </button>
                        </form>
                    </div>
                </div>
            </div>

            <!-- Tabla y Filtros -->
            <div class="col-lg-8">
                <div class="card shadow-sm border-0">
                    <div class="card-header bg-body py-3 d-flex flex-wrap justify-content-between align-items-center gap-2">
                        <h5 class="card-title mb-0 fs-6 fw-bold"><i class="fa-solid fa-list-check me-2 text-primary"></i>Historial de Asistencias</h5>
                        <form method="GET" action="/" class="d-flex gap-2 align-items-center">
                            <input type="date" class="form-control form-control-sm" name="fecha" value="{{ fecha_filtro }}">
                            <button type="submit" class="btn btn-primary btn-sm"><i class="fa-solid fa-filter"></i> Filtrar</button>
                            <a href="/" class="btn btn-outline-secondary btn-sm" title="Ver todo"><i class="fa-solid fa-rotate-right"></i></a>
                        </form>
                    </div>
                    <div class="card-body p-0">
                        <div class="table-responsive" style="max-height: 480px;">
                            <table class="table table-hover table-striped align-middle mb-0 text-now0wrap">
                                <thead class="table-dark sticky-top">
                                    <tr>
                                        <th>Cédula</th>
                                        <th>Nombre</th>
                                        <th>Municipio</th>
                                        <th>Comisión</th>
                                        <th>Fecha</th>
                                        <th>Hora</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {% if registros %}
                                        {% for row in registros %}
                                        <tr>
                                            <td class="fw-semibold">{{ row.CEDULA }}</td>
                                            <td>{{ row.NOMBRE }}</td>
                                            <td>{{ row.MUNICIPIO }}</td>
                                            <td><span class="badge bg-secondary-subtle text-secondary-emphasis border">{{ row.COMISION }}</span></td>
                                            <td>{{ row.FECHA }}</td>
                                            <td class="text-muted small"><i class="fa-regular fa-clock me-1"></i>{{ row.HORA }}</td>
                                        </tr>
                                        {% endfor %}
                                    {% else %}
                                        <tr>
                                            <td colspan="6" class="text-center py-4 text-muted">
                                                <i class="fa-solid fa-folder-open fs-3 mb-2 d-block"></i>
                                                No hay registros para la fecha seleccionada.
                                            </td>
                                        </tr>
                                    {% endif %}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function toggleTheme() {
            const htmlEl = document.documentElement;
            const themeIcon = document.getElementById('theme-icon');
            if (htmlEl.getAttribute('data-bs-theme') === 'light') {
                htmlEl.setAttribute('data-bs-theme', 'dark');
                themeIcon.className = 'fa-solid fa-sun';
                localStorage.setItem('theme', 'dark');
            } else {
                htmlEl.setAttribute('data-bs-theme', 'light');
                themeIcon.className = 'fa-solid fa-moon';
                localStorage.setItem('theme', 'light');
            }
        }
        // Cargar preferencia guardada
        window.addEventListener('DOMContentLoaded', () => {
            const savedTheme = localStorage.getItem('theme');
            if (savedTheme) {
                document.documentElement.setAttribute('data-bs-theme', savedTheme);
                if (savedTheme === 'dark') {
                    document.getElementById('theme-icon').className = 'fa-solid fa-sun';
                }
            }
        });
    </script>
</body>
</html>
"""

@app.route("/")
def index():
    df_base = cargar_datos_base()
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    fecha_filtro = request.args.get("fecha", fecha_hoy)
    
    registros = []
    if os.path.exists(ASISTENCIA_FILE):
        df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
        if not df_asist.empty:
            if fecha_filtro:
                df_asist = df_asist[df_asist["FECHA"] == fecha_filtro]
            registros = df_asist.iloc[::-1].to_dict(orient="records")
            
    return render_template_string(HTML_TEMPLATE, fecha_hoy=fecha_hoy, fecha_filtro=fecha_filtro, registros=registros)

@app.route("/registrar", methods=["POST"])
def registrar():
    cedula = request.form.get("cedula", "").strip()
    if not cedula:
        flash("Ingrese un número de cédula válido.", "warning")
        return redirect(url_for("index"))
        
    df_base = cargar_datos_base()
    if df_base.empty:
        flash("No se encontró el archivo de base de datos o está vacío.", "danger")
        return redirect(url_for("index"))
        
    persona = df_base[df_base["CEDULA"] == cedula]
    if persona.empty:
        flash(f"La cédula '{cedula}' no está registrada en el archivo base.", "danger")
        return redirect(url_for("index"))
        
    row = persona.iloc[0]
    nombre = row.get("NOMBRE Y APELLIDO", "N/D")
    municipio = row.get("MUNICIPIO", "N/D")
    parroquia = row.get("PARroquia", row.get("PARROQUIA", "N/D"))
    comision = row.get("COMISION", "N/D")
    
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    hora_actual = datetime.now().strftime("%H:%M:%S")
    
    # Validar si ya registró hoy
    ya_asistio = False
    if os.path.exists(ASISTENCIA_FILE):
        df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
        if not df_asist.empty:
            duplicado = df_asist[(df_asist["CEDULA"] == cedula) & (df_asist["FECHA"] == fecha_hoy)]
            if not duplicado.empty:
                ya_asistio = True
                
    if ya_asistio:
        flash(f"AVISO: {nombre} (Cédula: {cedula}) ya tenía registrada su asistencia hoy ({fecha_hoy}).", "warning")
    else:
        nuevo_reg = {
            "CEDULA": cedula,
            "NOMBRE": nombre,
            "MUNICIPIO": municipio,
            "PARROQUIA": parroquia,
            "COMISION": comision,
            "FECHA": fecha_hoy,
            "HORA": hora_actual
        }
        df_nuevo = pd.DataFrame([nuevo_reg])
        if not os.path.exists(ASISTENCIA_FILE):
            df_nuevo.to_csv(ASISTENCIA_FILE, index=False)
        else:
            df_nuevo.to_csv(ASISTENCIA_FILE, mode="a", header=False, index=False)
        flash(f"¡Asistencia Registrada! -> {nombre} ({comision}) a las {hora_actual}", "success")
        
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)