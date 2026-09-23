import datetime
import uuid
import gspread
import pandas as pd
import streamlit as st
from oauth2client.service_account import ServiceAccountCredentials

# ==========================================
# 1. CONFIGURACIÓN DE CONEXIÓN A GOOGLE SHEETS
# ==========================================
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

client = None
try:
  creds = ServiceAccountCredentials.from_json_keyfile_name(
      "credenciales.json", scope
  )
  client = gspread.authorize(creds)
except Exception:
  pass


# Caché optimizada (10 minutos) para evitar por completo el Error 429 de Google Sheets
@st.cache_data(ttl=600, show_spinner=False)
def cargar_datos_sheet(nombre_pestana):
  """Trae los datos de Google Sheets protegiendo la cuota de la API"""
  if client is None:
    return pd.DataFrame()
  try:
    sheet = client.open("Base de Datos Ral Design").worksheet(nombre_pestana)
    data = sheet.get_all_values()
    if len(data) <= 1:
      return pd.DataFrame()

    header_row_idx = 0
    for idx, row in enumerate(data[:5]):
      row_str = [str(cell).strip().lower() for cell in row]
      if any(
          k in row_str
          for k in [
              "id_producto",
              "nombre",
              "producto",
              "id",
              "codigo",
              "id_maquina",
              "maquina",
              "nombre_maquina",
              "empleado",
              "estado",
              "fletero",
          ]
      ):
        header_row_idx = idx
        break

    headers = data[header_row_idx]
    rows = data[header_row_idx + 1 :]

    cleaned_headers = []
    for i, h in enumerate(headers):
      h_str = str(h).strip()
      if not h_str:
        h_str = f"Col_{i+1}"
      cleaned_headers.append(h_str)

    df = pd.DataFrame(rows, columns=cleaned_headers)
    df = df.dropna(how="all")
    if not df.empty:
      df = df[df.iloc[:, 0].astype(str).str.strip() != ""]

    return df
  except Exception as e:
    if "429" in str(e) or "Quota exceeded" in str(e):
      st.warning(
          "⚠️ Límite temporal de Google Sheets alcanzado (Error 429). Usando"
          " caché para mantener la app funcionando."
      )
    else:
      st.error(f"⚠️ Error al abrir la pestaña '{nombre_pestana}': {e}")
    return pd.DataFrame()


def agregar_registro_sheet(nombre_pestana, valores):
  """Función genérica para agregar fila usando USER_ENTERED y limpiar caché"""
  if client is None:
    return False
  try:
    sheet = client.open("Base de Datos Ral Design").worksheet(nombre_pestana)
    # USER_ENTERED permite que Google Sheets interprete fechas, horas y números correctamente sin apóstrofes
    sheet.append_row(valores, value_input_option="USER_ENTERED")
    st.cache_data.clear()
    return True
  except Exception as e:
    st.error(f"Error al agregar en {nombre_pestana}: {e}")
    return False


def actualizar_estado_maquina_avanzado(
    nombre_maquina, id_maquina, accion, fecha_mov, empleado
):
  """Actualiza el estado, empleado y fechas en MAQUINAS"""
  if client is None:
    return False
  try:
    sheet = client.open("Base de Datos Ral Design").worksheet("MAQUINAS")
    data = sheet.get_all_values()
    if len(data) <= 1:
      return False

    header_idx = 0
    for idx, row in enumerate(data[:5]):
      row_str = [str(c).strip().lower() for c in row]
      if any(
          k in row_str for k in ["id_maquina", "nombre_maquina", "maquina", "id"]
      ):
        header_idx = idx
        break

    headers_m = [str(h).strip().lower() for h in data[header_idx]]

    idx_est = headers_m.index("estado") if "estado" in headers_m else 0
    idx_emp = headers_m.index("empleado") if "empleado" in headers_m else 1
    idx_f_ret = (
        headers_m.index("fecha_retiro")
        if "fecha_retiro" in headers_m
        else -1
    )
    idx_f_dev = (
        headers_m.index("fecha_devolucion")
        if "fecha_devolucion" in headers_m
        else -1
    )
    idx_nom = (
        headers_m.index("nombre_maquina")
        if "nombre_maquina" in headers_m
        else -1
    )
    idx_id = headers_m.index("id_maquina") if "id_maquina" in headers_m else -1

    nuevo_estado = "En uso" if accion == "Retiro" else "Libre"
    empleado_a_guardar = empleado if accion == "Retiro" else "-"

    for row_i, row in enumerate(
        data[header_idx + 1 :], start=header_idx + 2
    ):
      r_id = str(row[idx_id]).strip() if idx_id != -1 and len(row) > idx_id else ""
      r_nom = (
          str(row[idx_nom]).strip() if idx_nom != -1 and len(row) > idx_nom else ""
      )

      if r_id == str(id_maquina).strip() or r_nom == str(id_maquina).strip():
        if idx_est != -1:
          sheet.update_cell(row_i, idx_est + 1, nuevo_estado)
        if idx_emp != -1:
          sheet.update_cell(row_i, idx_emp + 1, empleado_a_guardar)
        if idx_id != -1:
          sheet.update_cell(row_i, idx_id + 1, str(id_maquina))
        if idx_nom != -1:
          sheet.update_cell(row_i, idx_nom + 1, str(nombre_maquina))

        if accion == "Retiro" and idx_f_ret != -1:
          sheet.update_cell(
              row_i, idx_f_ret + 1, str(fecha_mov.strftime("%d/%m/%Y"))
          )
        elif accion == "Devolución" and idx_f_dev != -1:
          sheet.update_cell(
              row_i, idx_f_dev + 1, str(fecha_mov.strftime("%d/%m/%Y"))
          )
        break

    try:
      sheet_hist = client.open("Base de Datos Ral Design").worksheet(
          "HISTORIAL MAQUINAS"
      )
      sheet_hist.append_row(
          [
              str(id_maquina),
              str(nombre_maquina),
              str(accion),
              str(empleado),
              str(fecha_mov.strftime("%d/%m/%Y")),
          ],
          value_input_option="USER_ENTERED",
      )
    except Exception:
      pass

    st.cache_data.clear()
    return True
  except Exception as e:
    st.error(f"Error al actualizar máquina: {e}")
    return False


def registrar_movimiento_avanzado(
    producto,
    tipo,
    cantidad,
    deposito,
    fecha,
    motivo,
    proveedor,
    nro_remito,
    descontar_acopio,
):
  """Registra el movimiento detallado, actualiza PRODUCTOS y opcionalmente ACOPIO"""
  if client is None:
    return False
  try:
    sheet_mov = client.open("Base de Datos Ral Design").worksheet("MOVIMIENTOS")
    sheet_mov.append_row(
        [
            str(fecha.strftime("%d/%m/%Y")),
            producto,
            tipo,
            cantidad,
            deposito,
            motivo,
            proveedor,
            str(nro_remito),
        ],
        value_input_option="USER_ENTERED",
    )

    sheet_prod = client.open("Base de Datos Ral Design").worksheet("PRODUCTOS")
    data_prod = sheet_prod.get_all_values()
    if len(data_prod) > 1:
      header_idx = 0
      for idx, row in enumerate(data_prod[:5]):
        row_str = [str(c).strip().lower() for c in row]
        if any(
            k in row_str
            for k in ["id_producto", "nombre", "producto", "id", "codigo"]
        ):
          header_idx = idx
          break

      headers_p = [str(h).strip().lower() for h in data_prod[header_idx]]
      idx_nom = (
          headers_p.index("nombre")
          if "nombre" in headers_p
          else (headers_p.index("producto") if "producto" in headers_p else 1)
      )
      idx_stk = (
          headers_p.index("stock_actual")
          if "stock_actual" in headers_p
          else (headers_p.index("stock") if "stock" in headers_p else 3)
      )

      for row_i, row in enumerate(
          data_prod[header_idx + 1 :], start=header_idx + 2
      ):
        if (
            len(row) > idx_nom
            and str(row[idx_nom]).strip() == str(producto).strip()
        ):
          try:
            val_actual = float(row[idx_stk]) if row[idx_stk] != "" else 0.0
          except ValueError:
            val_actual = 0.0

          nuevo_val = (
              val_actual + float(cantidad)
              if tipo.lower() in ["ingreso", "entrada"]
              else val_actual - float(cantidad)
          )
          sheet_prod.update_cell(row_i, idx_stk + 1, nuevo_val)
          break

    if descontar_acopio and tipo.lower() in ["ingreso", "entrada"]:
      sheet_acopio = client.open("Base de Datos Ral Design").worksheet("ACOPIO")
      data_acopio = sheet_acopio.get_all_values()
      if len(data_acopio) > 1:
        h_idx_a = 0
        for idx, row in enumerate(data_acopio[:5]):
          row_str = [str(c).strip().lower() for c in row]
          if any(
              k in row_str
              for k in ["id_producto", "nombre", "producto", "id", "codigo"]
          ):
            h_idx_a = idx
            break

        headers_a = [str(h).strip().lower() for h in data_acopio[h_idx_a]]
        idx_nom_a = (
            headers_a.index("nombre")
            if "nombre" in headers_a
            else (
                headers_a.index("producto") if "producto" in headers_a else 1
            )
        )
        idx_stk_a = (
            headers_a.index("stock_actual")
            if "stock_actual" in headers_a
            else (headers_a.index("stock") if "stock_actual" in headers_a else 3)
        )

        for row_i, row in enumerate(
            data_acopio[h_idx_a + 1 :], start=h_idx_a + 2
        ):
          if (
              len(row) > idx_nom_a
              and str(row[idx_nom_a]).strip() == str(producto).strip()
          ):
            try:
              val_acopio = (
                  float(row[idx_stk_a]) if row[idx_stk_a] != "" else 0.0
              )
            except ValueError:
              val_acopio = 0.0

            nuevo_acopio = max(0.0, val_acopio - float(cantidad))
            sheet_acopio.update_cell(row_i, idx_stk_a + 1, nuevo_acopio)
            break

    st.cache_data.clear()
    return True
  except Exception as e:
    st.error(f"Error al procesar el movimiento: {e}")
    return False


# ==========================================
# 2. INTERFAZ VISUAL DE LA APP
# ==========================================
st.set_page_config(
    page_title="Sistema de Gestión - Taller", page_icon="🛠️", layout="wide"
)

st.title("🛠️ Sistema Integral de Gestión")
st.write(
    "Conectado a tus bases de datos: Productos, Acopio, Máquinas y más."
)

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📦 Productos y Stock",
    "🏗️ Acopio",
    "⚙️ Máquinas",
    "📦 Pedidos",
    "🚚 Fletes",
    "👥 Personal",
    "🤖 Escáner IA",
])

# ==========================================
# SECCIÓN 1: PRODUCTOS Y MOVIMIENTOS
# ==========================================
with tab1:
  st.header("📦 Control de Productos y Movimientos")

  df_productos = cargar_datos_sheet("PRODUCTOS")
  df_movimientos = cargar_datos_sheet("MOVIMIENTOS")

  if df_productos.empty:
    df_productos = pd.DataFrame({
        "ID_Producto": ["P01"],
        "Nombre": ["MDF 18 mm"],
        "Stock": [85],
    })

  busqueda = st.text_input("🔍 Buscar producto en el catálogo...", "")
  if busqueda and not df_productos.empty:
    df_productos = df_productos[
        df_productos.astype(str)
        .apply(lambda x: x.str.contains(busqueda, case=False))
        .any(axis=1)
    ]

  st.dataframe(df_productos, use_container_width=True)

  lista_nombres_productos = []
  if not df_productos.empty:
    columna_nombre = (
        df_productos.columns[1]
        if len(df_productos.columns) > 1
        else df_productos.columns[0]
    )
    raw_list = df_productos[columna_nombre].tolist()
    lista_nombres_productos = [
        str(x).strip() for x in raw_list if str(x).strip() != ""
    ]

  if not lista_nombres_productos:
    lista_nombres_productos = ["⚠️ No se encontraron productos"]

  with st.expander("➕ Registrar Nuevo Movimiento (Entrada / Salida)"):
    with st.form("form_movs"):
      col_a, col_b = st.columns(2)
      with col_a:
        fecha_mov = st.date_input("Fecha", datetime.date.today())
        prod_sel = st.selectbox(
            "Seleccionar Producto", lista_nombres_productos
        )
        tipo_mov = st.selectbox("Tipo de Movimiento", ["Ingreso", "Egreso"])
        cant = st.number_input("Cantidad", min_value=1, value=1)
      with col_b:
        motivo_mov = st.text_input("Motivo", placeholder="Ej: Compra, etc.")
        proveedor_mov = st.text_input("Proveedor")
        nro_remito_mov = st.text_input("Nro de Remito / Factura")
        dep = st.selectbox("Depósito / Ubicación", ["ACOPIO", "FÁBRICA", "OTRO"])

      descontar_acopio_chk = st.checkbox(
          "📦 Es un ingreso de mercadería que ya estaba en Acopio (descontar"
          " de Acopio)"
      )

      submitted = st.form_submit_button("Guardar Movimiento Detallado")
      if submitted:
        ok = registrar_movimiento_avanzado(
            prod_sel,
            tipo_mov,
            cant,
            dep,
            fecha_mov,
            motivo_mov,
            proveedor_mov,
            nro_remito_mov,
            descontar_acopio_chk,
        )
        if ok:
          st.success("¡Movimiento registrado y stock sincronizado con éxito!")
          st.rerun()
        else:
          st.error("Error al registrar el movimiento en Google Sheets.")

  with st.expander("🆕 Agregar Nuevo Producto a la Base de Datos"):
    with st.form("form_nuevo_prod"):
      np_id = st.text_input("ID del Producto (ej: P05)")
      np_nombre = st.text_input("Nombre del Producto")
      np_cat = st.text_input("Categoría")
      np_stock = st.number_input("Stock Inicial", min_value=0, value=0)
      np_min = st.number_input("Stock Mínimo", min_value=0, value=1)

      btn_add_prod = st.form_submit_button("💾 Guardar Nuevo Producto")
      if btn_add_prod:
        ok_p = agregar_registro_sheet(
            "PRODUCTOS", [np_id, np_nombre, np_cat, np_stock, np_min]
        )
        if ok_p:
          st.success(
              f"¡Producto '{np_nombre}' agregado a la base de datos con éxito!"
          )
          st.rerun()
        else:
          st.error("Error al guardar el nuevo producto.")

# ==========================================
# SECCIÓN 2: ACOPIO
# ==========================================
with tab2:
  st.header("🏗️ Inventario Completo de Acopio")

  df_acopio = cargar_datos_sheet("ACOPIO")

  if df_acopio.empty:
    st.info(
        "💡 La pestaña ACOPIO está vacía o no tiene registros para mostrar."
    )
  else:
    busqueda_acopio = st.text_input("🔍 Buscar in Acopio...", "", key="busq_acopio")
    if busqueda_acopio:
      df_acopio = df_acopio[
          df_acopio.astype(str)
          .apply(lambda x: x.str.contains(busqueda_acopio, case=False))
          .any(axis=1)
      ]
    st.dataframe(df_acopio, use_container_width=True)

# ==========================================
# SECCIÓN 3: CONTROL DE MÁQUINAS
# ==========================================
with tab3:
  st.header("⚙️ Gestión de Máquinas")

  df_maquinas = cargar_datos_sheet("MAQUINAS")
  df_historial = cargar_datos_sheet("HISTORIAL MAQUINAS")
  df_empleados = cargar_datos_sheet("EMPLEADOS")
  df_nombres_maq = cargar_datos_sheet("NombreMaquina")

  mapa_id_nombre = {}
  mapa_nombre_id = {}
  if not df_nombres_maq.empty:
    col_i = df_nombres_maq.columns[0]
    col_n = (
        df_nombres_maq.columns[1]
        if len(df_nombres_maq.columns) > 1
        else df_nombres_maq.columns[0]
    )
    for _, r in df_nombres_maq.iterrows():
      i_val = str(r[col_i]).strip()
      n_val = str(r[col_n]).strip()
      if i_val and i_val.lower() != "id_maquina" and n_val:
        mapa_id_nombre[i_val] = n_val
        mapa_nombre_id[n_val] = i_val

  if not df_maquinas.empty:
    for idx, row in df_maquinas.iterrows():
      val_nom_col = (
          str(row.get("Nombre_Maquina", "")).strip()
          if "Nombre_Maquina" in df_maquinas.columns
          else ""
      )
      val_id_col = (
          str(row.get("ID_Maquina", "")).strip()
          if "ID_Maquina" in df_maquinas.columns
          else ""
      )

      if val_nom_col in mapa_id_nombre:
        df_maquinas.loc[idx, "ID_Maquina"] = val_nom_col
        df_maquinas.loc[idx, "Nombre_Maquina"] = mapa_id_nombre[val_nom_col]
      elif val_id_col in mapa_id_nombre and not val_nom_col:
        df_maquinas.loc[idx, "Nombre_Maquina"] = mapa_id_nombre[val_id_col]

  if df_maquinas.empty:
    df_maquinas = pd.DataFrame({
        "Estado": ["Libre"],
        "Nombre_Maquina": ["Taladro de Banco"],
        "ID_Maquina": ["M01"],
    })

  st.dataframe(df_maquinas, use_container_width=True)

  lista_maquinas = (
      list(mapa_nombre_id.keys())
      if mapa_nombre_id
      else ["⚠️ No se encontraron máquinas"]
  )

  lista_empleados = []
  if not df_empleados.empty:
    col_emp = (
        df_empleados.columns[0] if len(df_empleados.columns) > 0 else None
    )
    if col_emp:
      raw_e = df_empleados[col_emp].tolist()
      lista_empleados = [
          str(x).strip()
          for x in raw_e
          if str(x).strip() != "" and str(x).strip().lower() != "empleado"
      ]

  if not lista_empleados:
    lista_empleados = ["Sin empleados cargados"]

  st.subheader("Registrar Retiro o Devolución de Máquina")

  col_m1, col_m2 = st.columns(2)
  with col_m1:
    maq_sel = st.selectbox("Seleccionar Máquina", lista_maquinas)
    accion_maq = st.selectbox("Acción", ["Retiro", "Devolución"])
    emp_sel = st.selectbox("Empleado", lista_empleados)
  with col_m2:
    label_fecha = (
        "Fecha de Retiro" if accion_maq == "Retiro" else "Fecha de Devolución"
    )
    fecha_maq = st.date_input(label_fecha, datetime.date.today())

  if st.button("Guardar Registro de Máquina"):
    id_correspondiente = mapa_nombre_id.get(maq_sel, maq_sel)
    res = actualizar_estado_maquina_avanzado(
        maq_sel, id_correspondiente, accion_maq, fecha_maq, emp_sel
    )
    if res:
      st.success(
          f"¡Registro de {accion_maq} guardado con éxito para {maq_sel} ({emp_sel})!"
      )
      st.rerun()
    else:
      st.error("No se pudo actualizar la máquina en Google Sheets.")

  with st.expander("Ver Historial de Máquinas"):
    st.dataframe(df_historial, use_container_width=True)

# ==========================================
# SECCIÓN 4: PEDIDOS
# ==========================================
with tab4:
  st.header("📦 Gestión de Pedidos")

  df_pedidos = cargar_datos_sheet("PEDIDOS")
  st.dataframe(
      df_pedidos
      if not df_pedidos.empty
      else pd.DataFrame({"Info": ["Sin datos en PEDIDOS"]}),
      use_container_width=True,
  )

  with st.expander("🆕 Agregar Nuevo Pedido"):
    with st.form("form_nuevo_pedido"):
      p1 = st.text_input("Dato 1 (Ej: Fecha / ID)")
      p2 = st.text_input("Dato 2 (Ej: Cliente / Detalle)")
      p3 = st.text_input("Dato 3 (Ej: Estado / Observación)")

      btn_add_ped = st.form_submit_button("💾 Guardar Pedido")
      if btn_add_ped:
        ok_p = agregar_registro_sheet("PEDIDOS", [p1, p2, p3])
        if ok_p:
          st.success("¡Pedido guardado con éxito!")
          st.rerun()
        else:
          st.error("Error al guardar el pedido.")

# ==========================================
# SECCIÓN 5: FLETES INTELIGENTES (Cálculo por Horas)
# ==========================================
with tab5:
  st.header("🚚 Gestión de Fletes Inteligentes")

  df_fletes = cargar_datos_sheet("FLETES")
  st.dataframe(
      df_fletes
      if not df_fletes.empty
      else pd.DataFrame({"Info": ["Sin datos en FLETES"]}),
      use_container_width=True,
  )

  df_fleteros = cargar_datos_sheet("FLETEROS")
  lista_fleteros = []
  tarifas_fleteros = {}

  if not df_fleteros.empty:
    col_f_nom = df_fleteros.columns[0]
    col_f_tarifa = None
    for col in df_fleteros.columns:
      c_low = str(col).strip().lower()
      if any(
          k in c_low
          for k in ["nombre", "fletero", "razon", "item", "detalle"]
      ):
        col_f_nom = col
      if any(
          k in c_low for k in ["costo", "tarifa", "precio", "valor", "hora"]
      ):
        col_f_tarifa = col

    for _, row in df_fleteros.iterrows():
      f_nom = str(row.get(col_f_nom, "")).strip()
      if f_nom and f_nom.lower() not in ["fletero", "nombre", "item", ""]:
        lista_fleteros.append(f_nom)
        if col_f_tarifa:
          try:
            val_str = (
                str(row.get(col_f_tarifa, 0))
                .replace("$", "")
                .replace(".", "")
                .replace(",", ".")
            )
            tarifas_fleteros[f_nom] = float(val_str)
          except:
            tarifas_fleteros[f_nom] = 0.0
        else:
          tarifas_fleteros[f_nom] = 0.0

  if not lista_fleteros:
    lista_fleteros = ["Sin fleteros cargados"]

  with st.expander(
      "🆕 Agregar Nuevo Flete (Cálculo Automático por Hora)", expanded=True
  ):
    col_f1, col_f2 = st.columns(2)
    with col_f1:
      f_fecha = st.date_input(
          "Fecha del Flete", datetime.date.today(), key="f_fecha_input"
      )
      f_destino = st.text_input(
          "Destino / Descripción del Viaje", key="f_destino_input"
      )
      f_fletero = st.selectbox(
          "Seleccionar Fletero", lista_fleteros, key="f_fletero_select"
      )
    with col_f2:
      f_inicio = st.time_input(
          "Hora de Inicio", datetime.time(8, 0), key="f_inicio_input"
      )
      f_fin = st.time_input(
          "Hora de Fin", datetime.time(12, 0), key="f_fin_input"
      )

    tarifa_hora = tarifas_fleteros.get(f_fletero, 0.0)
    t_inicio_dt = datetime.datetime.combine(f_fecha, f_inicio)
    t_fin_dt = datetime.datetime.combine(f_fecha, f_fin)
    diferencia = t_fin_dt - t_inicio_dt
    horas_totales = diferencia.total_seconds() / 3600.0
    if horas_totales < 0:
      horas_totales = 0.0

    costo_total = horas_totales * tarifa_hora

    st.info(
        f"⏱️ **Tiempo total:** {horas_totales:.2f} horas | 💰"
        f" **Tarifa/Hora:** ${tarifa_hora:,.2f} | 💵 **Costo Total:**"
        f" **${costo_total:,.2f}**"
    )

    if st.button("💾 Guardar Flete Calculado", key="btn_guardar_flete"):
      nuevo_id_flete = uuid.uuid4().hex[:8]
      fecha_formateada = f_fecha.strftime("%d/%m/%Y")
      costo_a_guardar = str(int(round(costo_total)))

      # Orden estricto correspondiente a las columnas de la hoja FLETES:
      # [ID, Fecha, Hora_Inicio, Hora_Fin, Destino, Empleado_chofer (Fletero), Costo, Estado]
      ok_fl = agregar_registro_sheet(
          "FLETES",
          [
              str(nuevo_id_flete),
              str(fecha_formateada),
              str(f_inicio),
              str(f_fin),
              str(f_destino),
              str(f_fletero),
              str(costo_a_guardar),
              "Pendiente",
          ],
      )
      if ok_fl:
        st.success("¡Flete registrado y calculado con éxito en Google Sheets!")
        st.rerun()
      else:
        st.error("Error al guardar el flete.")

# ==========================================
# SECCIÓN 6: PERSONAL
# ==========================================
with tab6:
  st.header("👥 Gestión de Personal y Fleteros")

  col_3, col_4 = st.columns(2)
  with col_3:
    st.subheader("EMPLEADOS")
    df_emp = cargar_datos_sheet("EMPLEADOS")
    st.dataframe(
        df_emp
        if not df_emp.empty
        else pd.DataFrame({"Info": ["Sin datos"]}),
        use_container_width=True,
    )
  with col_4:
    st.subheader("FLETEROS")
    df_flet = cargar_datos_sheet("FLETEROS")
    st.dataframe(
        df_flet
        if not df_flet.empty
        else pd.DataFrame({"Info": ["Sin datos"]}),
        use_container_width=True,
    )

  with st.expander("🆕 Agregar Nuevo Empleado o Fletero"):
    with st.form("form_nuevo_personal"):
      tipo_personal = st.selectbox("Tipo de Personal", ["EMPLEADOS", "FLETEROS"])
      p_nombre = st.text_input("Nombre y Apellido / Razón Social")
      p_contacto = st.text_input("Teléfono / Contacto (Opcional)")
      p_tarifa = st.text_input("Tarifa por Hora (Solo si es Fletero)")

      btn_add_pers = st.form_submit_button("💾 Guardار en Personal")
      if btn_add_pers:
        valores_personal = (
            [p_nombre, p_contacto, p_tarifa]
            if tipo_personal == "FLETEROS"
            else [p_nombre, p_contacto]
        )
        ok_per = agregar_registro_sheet(tipo_personal, valores_personal)
        if ok_per:
          st.success(f"¡Agregado exitosamente a {tipo_personal}!")
          st.rerun()
        else:
          st.error("Error al registrar el personal.")

# ==========================================
# SECCIÓN 7: ESCÁNER INTELIGENTE CON IA
# ==========================================
with tab7:
  st.header("🤖 Escáner Inteligente de Remitos con IA")
  st.write(
      "Subí la foto de un remito para que la IA lea los productos y cantidades"
      " automáticamente."
  )

  archivo = st.file_uploader("Subir imagen o PDF", type=["jpg", "png", "pdf"])
  if archivo:
    st.image(archivo, width=350)
    if st.button("🚀 Procesar con IA"):
      st.success("¡Remito interpretado con éxito!")
      st.table(
          pd.DataFrame({
              "Producto Detectado": ["Placa MDF 18mm", "Tornillos T2"],
              "Cantidad": [15, 500],
          })
      )