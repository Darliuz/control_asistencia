from datetime import datetime
import os
from apscheduler.schedulers.background import BackgroundScheduler
from flask import Flask, flash, redirect, render_template_string, request, url_for
import pandas as pd

# Archivos de trabajo
EXCEL_DATA = "CONSOLIDADO EPC.xlsx"
ASISTENCIA_FILE = "registro_asistencias.csv"

app = Flask(__name__)
app.secret_key = "clave_secreta_equipo_politico"

# --- HTML TEMPLATE EN UNA SOLA CADENA ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Control de Asistencia - EPC</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
    <div class="container py-4">
        <h2 class="text-center text-success mb-4">📋 Control Diario de Asistencia - Equipo Político Comunal</h2>
        
        <!-- Mensajes Flash -->
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ category }} alert-dismissible fade show" role="alert">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <!-- Formulario de Registro -->
        <div class="card shadow-sm mb-4">
            <div class="card-header bg-success text-white">Registrar Asistencia por Cédula</div>
            <div class="card-body">
                <form method="POST" action="/registrar" class="row g-3">
                    <div class="col-md-8">
                        <label for="cedula" class="form-label">Número de Cédula:</label>
                        <input type="text" class="form-control form-control-lg" id="cedula" name="cedula" placeholder="Ingrese la cédula sin puntos" required autofocus>
                    </div>
                    <div class="col-md-4 d-flex align-items-end">
                        <button type="submit" class="btn btn-success btn-lg w-100">Registrar Asistencia</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- Filtros y Visualización -->
        <div class="card shadow-sm">
            <div class="card-header bg-dark text-white d-flex justify-content-between align-items-center">
                <span>Historial de Asistencias</span>
                <form method="GET" action="/" class="d-flex gap-2">
                    <input type="date" name="fecha" class="form-control form-control-sm" value="{{ fecha_actual }}">
                    <button type="submit" class="btn btn-primary btn-sm">Filtrar</button>
                    <a href="/" class="btn btn-secondary btn-sm">Ver Todo</a>
                </form>
            </div>
            <div class="card-body">
                <div class="table-responsive">
                    <table class="table table-striped table-hover text-center align-middle">
                        <thead class="table-secondary">
                            <tr>
                                <th>Cédula</th>
                                <th>Nombre</th>
                                <th>Municipio</th>
                                <th>Parroquia</th>
                                <th>Comisión</th>
                                <th>Fecha</th>
                                <th>Hora</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% if registros %}
                                {% for row in registros %}
                                <tr>
                                    <td>{{ row.CEDULA }}</td>
                                    <td>{{ row.NOMBRE }}</td>
                                    <td>{{ row.MUNICIPIO }}</td>
                                    <td>{{ row.PARROQUIA }}</td>
                                    <td>{{ row.COMISION }}</td>
                                    <td>{{ row.FECHA }}</td>
                                    <td>{{ row.HORA }}</td>
                                </tr>
                                {% endfor %}
                            {% else %}
                                <tr>
                                    <td colspan="7" class="text-muted">No hay registros para mostrar.</td>
                                </tr>
                            {% endif %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"""


def cargar_datos_base():
  """Carga el archivo Excel principal"""
  if not os.path.exists(EXCEL_DATA):
    return pd.DataFrame()
  try:
    df = pd.read_excel(EXCEL_DATA, sheet_name="EQUIPO POLITICO", header=1)
    df.columns = [str(c).strip() for c in df.columns]
    df["CEDULA"] = df["CEDULA"].astype(str).str.split(".").str[0].str.strip()
    return df
  except Exception as e:
    print(f"Error cargando Excel base: {e}")
    return pd.DataFrame()


def tarea_automatica_medianoche():
  """Guarda y exporta automáticamente la asistencia a las 11:59 PM"""
  fecha_hoy = datetime.now().strftime("%Y-%m-%d")
  if not os.path.exists(ASISTENCIA_FILE):
    return

  df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
  if df_asist.empty:
    return

  df_filtrado = df_asist[df_asist["FECHA"] == fecha_hoy]
  if not df_filtrado.empty:
    nombre_archivo = f"Cierre_Asistencia_{fecha_hoy.replace('-', '_')}.xlsx"
    try:
      df_filtrado.to_excel(nombre_archivo, index=False)
      print(f"[AUTOMÁTICO 11:59 PM] Archivo guardado: {nombre_archivo}")
    except Exception as e:
      print(f"[ERROR AUTOMÁTICO] No se pudo exportar: {e}")


@app.route("/", methods=["GET"])
def index():
  fecha_filtro = request.args.get("fecha", datetime.now().strftime("%Y-%m-%d"))

  if not os.path.exists(ASISTENCIA_FILE):
    registros = []
  else:
    df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
    if fecha_filtro and not df_asist.empty:
      df_asist = df_asist[df_asist["FECHA"] == fecha_filtro]
    registros = df_asist.iloc[::-1].to_dict(orient="records") if not df_asist.empty else []

  return render_template_string(
      HTML_TEMPLATE, registros=registros, fecha_actual=fecha_filtro
  )


@app.route("/registrar", methods=["POST"])
def registrar():
  cedula = request.form.get("cedula", "").strip()
  if not cedula:
    flash("Ingrese un número de cédula válido.", "warning")
    return redirect(url_for("index"))

  df_base = cargar_datos_base()
  if df_base.empty:
    flash("Error: No se encontró o está vacío el archivo base 'CONSOLIDADO EPC.xlsx'.", "danger")
    return redirect(url_for("index"))

  persona = df_base[df_base["CEDULA"] == cedula]
  if persona.empty:
    flash(f"La cédula '{cedula}' no está registrada en el archivo base.", "danger")
    return redirect(url_for("index"))

  row = persona.iloc[0]
  nombre = row.get("NOMBRE Y APELLIDO", "N/D")
  municipio = row.get("MUNICIPIO", "N/D")
  parroquia = row.get("PARROQUIA", "N/D")
  comision = row.get("COMISION", "N/D")

  fecha_hoy = datetime.now().strftime("%Y-%m-%d")
  hora_actual = datetime.now().strftime("%H:%M:%S")

  # Validar duplicados en el día
  ya_asistio = False
  if os.path.exists(ASISTENCIA_FILE):
    df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
    if not df_asist.empty:
      duplicado = df_asist[(df_asist["CEDULA"] == cedula) & (df_asist["FECHA"] == fecha_hoy)]
      if not duplicado.empty:
        ya_asistio = True

  if ya_asistio:
    flash(f"⚠️ AVISO: {nombre} ya tenía registrada su asistencia hoy.", "warning")
  else:
    nuevo_reg = {
        "CEDULA": cedula,
        "NOMBRE": nombre,
        "MUNICIPIO": municipio,
        "PARROQUIA": parroquia,
        "COMISION": comision,
        "FECHA": fecha_hoy,
        "HORA": hora_actual,
    }
    df_nuevo = pd.DataFrame([nuevo_reg])
    if not os.path.exists(ASISTENCIA_FILE):
      df_nuevo.to_csv(ASISTENCIA_FILE, index=False)
    else:
      df_nuevo.to_csv(ASISTENCIA_FILE, mode="a", header=False, index=False)

    flash(f"✅ ¡Asistencia Registrada exitosamente para {nombre}!", "success")

  return redirect(url_for("index"))


if __name__ == "__main__":
  # Configurar el programador automático para las 11:59 PM (23:59)
  scheduler = BackgroundScheduler()
  scheduler.add_job(func=tarea_automatica_medianoche, trigger="cron", hour=23, minute=59)
  scheduler.start()

  print("=" * 60)
  print(" Servidor web de asistencia iniciado correctamente.")
  print(" URL Local: http://127.0.0.1:5000")
  print(" Tarea automática programada: Todos los días a las 11:59 PM.")
  print("=" * 60)

  # Ejecutar la aplicación web
  app.run(host="0.0.0.0", port=5000, debug=False)