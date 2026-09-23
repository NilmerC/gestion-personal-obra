
import io
import json
import os
import sqlite3
import hashlib
import secrets
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# GESTIÓN DE PERSONAL DE OBRA
# MVP OPERATIVO: SQLite + Excel + Streamlit
# Mantiene los 266 campos del Excel para no perder información.
# ============================================================

st.set_page_config(
    page_title="Gestión de Personal de Obra",
    page_icon="👷",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "reclutamiento.db"
DEFAULT_EXCEL = BASE_DIR / "202691814513JJC27174 (1).xlsx"

# -------------------- ESTILOS --------------------
st.markdown("""
<style>
    /* Estilo moderno para tarjetas de métricas */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 12px 15px;
        border-radius: 10px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
    }
    
    /* Bordes redondeados generales */
    div.stAlert {
        border-radius: 8px;
    }
<style>
:root{
 --navy:#102a43; --navy2:#173b5e; --blue:#1976d2; --green:#20b26b;
 --purple:#7856d6; --orange:#f5a623; --red:#e84c4c; --cyan:#16a6b6;
 --bg:#f4f7fb; --text:#243b53; --muted:#627d98; --border:#dfe7ef;
}
.stApp{background:var(--bg);color:var(--text)}
.block-container{padding-top:1rem;padding-bottom:2rem;max-width:1650px}
section[data-testid="stSidebar"]{
 background:linear-gradient(180deg,#102a43 0%,#0b2035 100%);
 min-width:255px;
}
section[data-testid="stSidebar"] *{color:#eef5fb!important}
.card{
 background:#fff;border:1px solid var(--border);border-radius:12px;
 padding:15px 18px;box-shadow:0 2px 8px rgba(16,42,67,.05);
 margin-bottom:12px;
}
.card-title{font-size:18px;font-weight:750;color:var(--text);margin-bottom:8px}
.small{color:var(--muted);font-size:13px}
.badge{display:inline-block;padding:4px 9px;border-radius:20px;font-size:12px;font-weight:700}
.success{background:#e8f7ef;color:#16794a}.warning{background:#fff4dc;color:#9a6500}
.danger{background:#ffe8e8;color:#b42318}.info{background:#e7f1ff;color:#1559a6}
.nav-brand{font-size:18px;font-weight:800;color:white;margin-bottom:2px}
.nav-sub{font-size:11px;color:#b9cce0;margin-bottom:16px}
.nav-section{color:#9eb7ce;font-size:11px;text-transform:uppercase;letter-spacing:.08em;margin:15px 0 7px}
div[data-testid="stMetric"]{
 background:#fff;border:1px solid var(--border);border-radius:12px;
 padding:13px 16px;box-shadow:0 2px 8px rgba(16,42,67,.04)
}
[data-testid="stSidebar"] .stButton button{
 text-align:left;border:0;background:transparent;width:100%;
 padding:8px 10px;border-radius:8px
}
[data-testid="stSidebar"] .stButton button:hover{background:rgba(255,255,255,.08)}
hr{border-color:#e3eaf1}
</style>
""", unsafe_allow_html=True)

# -------------------- DB HELPERS --------------------
def qident(name):
    return '"' + str(name).replace('"', '""') + '"'

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 120_000).hex()
    return salt + "$" + digest

def verify_password(password, stored):
    try:
        salt, digest = stored.split("$", 1)
        test = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 120_000).hex()
        return secrets.compare_digest(test, digest)
    except Exception:
        return False

def init_db():
    conn=get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS app_users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'reclutador',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS audit_log(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_id INTEGER,
            username TEXT,
            action TEXT,
            changes TEXT,
            created_at TEXT NOT NULL
        )
    """)
    # Cuenta inicial solo si no existe.
    if conn.execute("SELECT COUNT(*) FROM app_users").fetchone()[0] == 0:
        conn.execute(
            "INSERT INTO app_users(username,password_hash,role,created_at) VALUES(?,?,?,?)",
            ("admin", hash_password("admin123"), "admin", datetime.now().isoformat(timespec="seconds"))
        )
    conn.commit(); conn.close()

def table_exists():
    conn=get_conn()
    x=conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='personal'").fetchone()
    conn.close()
    return x is not None

def db_columns():
    conn=get_conn()
    rows=conn.execute("PRAGMA table_info(personal)").fetchall()
    conn.close()
    return [r["name"] for r in rows]

def normalize_value(v):
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    s=str(v).strip()
    if s.lower() in {"nan","nat","none"}:
        return None
    return s

def read_excel(source):
    xls = pd.ExcelFile(source)
    sheet="Datos" if "Datos" in xls.sheet_names else xls.sheet_names[0]
    df=pd.read_excel(xls,sheet_name=sheet)
    df.columns=[str(c).strip() for c in df.columns]
    # Convertir fechas a ISO sin destruir columnas.
    date_keywords=("fecha","_fe","fe_","vence","caduca","nacimiento")
    for c in df.columns:
        lc=c.lower()
        if any(k in lc for k in date_keywords):
            try:
                parsed=pd.to_datetime(df[c],errors="coerce",dayfirst=True)
                if parsed.notna().sum() >= max(3,int(len(df)*.03)):
                    df[c]=parsed.map(lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) else None)
            except Exception:
                pass
    df=df.map(normalize_value)
    return df,sheet

def create_personal_table(df):
    conn=get_conn()
    conn.execute("DROP TABLE IF EXISTS personal")
    cols=[qident(c)+" TEXT" for c in df.columns]
    sql='CREATE TABLE personal (record_id INTEGER PRIMARY KEY AUTOINCREMENT, '+",".join(cols)+')'
    conn.execute(sql)
    names=[qident(c) for c in df.columns]
    placeholders=",".join(["?"]*len(df.columns))
    ins=f'INSERT INTO personal ({",".join(names)}) VALUES ({placeholders})'
    for row in df.itertuples(index=False,name=None):
        conn.execute(ins, list(row))
    conn.commit(); conn.close()

def migrate_excel(source, username="system"):
    df,sheet=read_excel(source)
    create_personal_table(df)
    conn=get_conn()
    conn.execute(
        "INSERT INTO audit_log(record_id,username,action,changes,created_at) VALUES(?,?,?,?,?)",
        (None,username,"IMPORTACION",json.dumps({"filas":len(df),"columnas":len(df.columns),"hoja":sheet},ensure_ascii=False),datetime.now().isoformat(timespec="seconds"))
    )
    conn.commit(); conn.close()
    return len(df),len(df.columns),sheet

@st.cache_data(ttl=10, show_spinner=False)
def load_df():
    conn=get_conn()
    df=pd.read_sql_query("SELECT * FROM personal",conn)
    conn.close()
    return df

def execute(sql, params=()):
    conn=get_conn()
    cur=conn.execute(sql,params)
    conn.commit()
    result=cur.fetchall() if cur.description else []
    conn.close()
    return result

def log_action(record_id, username, action, changes):
    conn=get_conn()
    conn.execute(
        "INSERT INTO audit_log(record_id,username,action,changes,created_at) VALUES(?,?,?,?,?)",
        (record_id,username,action,json.dumps(changes,ensure_ascii=False,default=str),datetime.now().isoformat(timespec="seconds"))
    )
    conn.commit();conn.close()

def fetch_record(record_id):
    conn=get_conn()
    row=conn.execute("SELECT * FROM personal WHERE record_id=?",(record_id,)).fetchone()
    conn.close()
    return dict(row) if row else None

def update_record(record_id, values, username):
    old=fetch_record(record_id)
    if not old: return False
    fields=[c for c in values if c in db_columns()]
    sets=", ".join(qident(c)+"=?" for c in fields)
    params=[normalize_value(values[c]) for c in fields]+[record_id]
    conn=get_conn()
    conn.execute(f"UPDATE personal SET {sets} WHERE record_id=?",params)
    conn.commit();conn.close()
    changes={}
    for c in fields:
        ov=old.get(c); nv=normalize_value(values[c])
        if str(ov)!=str(nv): changes[c]={"antes":ov,"despues":nv}
    if changes: log_action(record_id,username,"EDICION",changes)
    load_df.clear()
    return True

def insert_record(values, username):
    cols=[c for c in values if c in db_columns() and c!="record_id"]
    vals=[normalize_value(values[c]) for c in cols]
    conn=get_conn()
    cur=conn.execute(
        f"INSERT INTO personal ({','.join(qident(c) for c in cols)}) VALUES ({','.join('?' for _ in cols)})",
        vals
    )
    rid=cur.lastrowid
    conn.commit();conn.close()
    log_action(rid,username,"ALTA",values)
    load_df.clear()
    return rid

def delete_record(record_id, username):
    old=fetch_record(record_id)
    conn=get_conn()
    conn.execute("DELETE FROM personal WHERE record_id=?",(record_id,))
    conn.commit();conn.close()
    log_action(record_id,username,"BAJA",old or {})
    load_df.clear()

# -------------------- LOGIN --------------------
def login_screen():
    st.markdown("""
    <div style="max-width:480px;margin:8vh auto 0;">
      <div class="card">
        <div style="font-size:42px;text-align:center;">👷</div>
        <h1 style="text-align:center;">Gestión de Personal de Obra</h1>
        <p class="small" style="text-align:center;">Plataforma de Reclutamiento · Seguridad · Logística</p>
      </div>
    </div>
    """,unsafe_allow_html=True)
    _,mid,_=st.columns([1,2,1])
    with mid:
        with st.form("login"):
            user=st.text_input("Usuario")
            pwd=st.text_input("Contraseña",type="password")
            ok=st.form_submit_button("Ingresar",use_container_width=True)
        if ok:
            conn=get_conn()
            row=conn.execute("SELECT * FROM app_users WHERE username=? AND active=1",(user,)).fetchone()
            conn.close()
            if row and verify_password(pwd,row["password_hash"]):
                st.session_state.auth={"username":row["username"],"role":row["role"]}
                st.rerun()
            st.error("Usuario o contraseña incorrectos.")
    st.info("Primera puesta en marcha: usuario **admin** · contraseña **admin123**. Cámbiala inmediatamente desde Administración.")
    st.stop()

# -------------------- BOOT --------------------
init_db()
if "auth" not in st.session_state:
    st.session_state.auth=None

if not st.session_state.auth:
    login_screen()

user=st.session_state.auth["username"]
role=st.session_state.auth["role"]

if not table_exists():
    if DEFAULT_EXCEL.exists():
        with st.spinner("Primera instalación: migrando el Excel a la base local..."):
            nrows,ncols,sheet=migrate_excel(DEFAULT_EXCEL,user)
        st.success(f"Base creada: {nrows:,} registros · {ncols} campos · hoja {sheet}.")
    else:
        st.error("No se encuentra el Excel inicial.")
        st.stop()

df=load_df()

# -------------------- COLUMNAS CLAVE --------------------
def find_col(candidates):
    lower={str(c).lower():c for c in df.columns}
    for x in candidates:
        if x.lower() in lower:return lower[x.lower()]
    for c in df.columns:
        lc=str(c).lower()
        if any(x.lower() in lc for x in candidates):return c
    return None

C={
"obra":find_col(["Obras","Obra"]),
"dni":find_col(["DNI_CE"]),
"nom":find_col(["Nombres"]),
"ape":find_col(["Apellidos"]),
"esp":find_col(["Especialidad_Requerimiento","Especialidad_de_Campo"]),
"recl":find_col(["Reclutador"]),
"sol":find_col(["ING_Solicitante","Solicitante"]),
"proc":find_col(["Estado_del_Proceso"]),
"post":find_col(["Estado_Postulante"]),
"emo_estado":find_col(["Estado_Medico"]),
"emo_fecha":find_col(["Fecha_Examen_Médico"]),
"ind_estado":find_col(["Estado_Inducción_Cliente"]),
"ind_fecha":find_col(["Fecha_Inducción_Cliente"]),
"entrevista":find_col(["Fecha_Entrevista"]),
"doc":find_col(["Estado_Documento"]),
"ingreso":find_col(["Fecha_Ingreso"]),
"cese":find_col(["Fecha_Cese"]),
"fc_venc":find_col(["Fecha_Caduca_Fotocheck"]),
"retcc_venc":find_col(["Fecha_Vec_RETCC"]),
"req":find_col(["NroRequerimiento"]),
}

# -------------------- SIDEBAR --------------------
with st.sidebar:
    st.markdown('<div class="nav-brand">👷 Gestión de Personal de Obra</div>',unsafe_allow_html=True)
    st.markdown('<div class="nav-sub">Reclutamiento · Seguridad · Logística</div>',unsafe_allow_html=True)
    st.markdown(f'<div class="small">Usuario: <b>{user}</b> · {role}</div>',unsafe_allow_html=True)
    st.markdown("---")

    if "page" not in st.session_state: st.session_state.page="Inicio"
    nav=[
      ("🏠","Inicio"),("📌","Requerimientos"),("👥","Reclutamiento"),
      ("👤","Personal"),("🛡️","Seguridad"),("🎓","Capacitaciones"),
      ("📋","Evaluaciones"),("📄","Documentos"),("🦺","EPP"),
      ("🪪","Fotocheck"),("🚌","Movilización"),("🔄","Ingresos / Ceses"),
      ("📊","Reportes")
    ]
    if role=="admin":
        nav += [("⚙️","Administración")]
    for icon,label in nav:
        if st.button(f"{icon}  {label}",key="nav_"+label,use_container_width=True):
            st.session_state.page=label

    st.markdown("---")
    st.markdown('<div class="nav-section">Filtros</div>',unsafe_allow_html=True)
    obras=sorted([str(x) for x in df[C["obra"]].dropna().unique()]) if C["obra"] else []
    sel_obras=st.multiselect("Obra",obras)
    recl=sorted([str(x) for x in df[C["recl"]].dropna().unique()]) if C["recl"] else []
    sel_recl=st.multiselect("Reclutador",recl)
    search=st.text_input("Buscar DNI / nombre / obra",placeholder="Ej. 12345678")
    if st.button("🔄 Limpiar filtros",use_container_width=True):
        st.rerun()

    if st.button("🚪 Cerrar sesión",use_container_width=True):
        st.session_state.auth=None; st.rerun()

# -------------------- FILTRO --------------------
view=df.copy()
if sel_obras and C["obra"]: view=view[view[C["obra"]].astype(str).isin(sel_obras)]
if sel_recl and C["recl"]: view=view[view[C["recl"]].astype(str).isin(sel_recl)]
if search:
    mask=pd.Series(False,index=view.index)
    for c in [C["dni"],C["nom"],C["ape"],C["obra"],C["esp"],C["req"]]:
        if c: mask=mask | view[c].astype(str).str.contains(search,case=False,na=False)
    view=view[mask]

# -------------------- CABECERA --------------------
st.markdown(f"""
<div class="card" style="display:flex;justify-content:space-between;align-items:center">
 <div><div style="font-size:25px;font-weight:800">{st.session_state.page}</div>
 <div class="small">Información operativa actualizada desde la base del sistema.</div></div>
 <div style="text-align:right"><b>{datetime.now().strftime("%d/%m/%Y")}</b><br>
 <span class="small">{datetime.now().strftime("%H:%M")}</span></div>
</div>
""",unsafe_allow_html=True)

# -------------------- HELPERS DE UI --------------------
def metricas(data):
    total=len(data)
    enproc=int((data[C["proc"]].astype(str).str.upper()=="EN PROCESO").sum()) if C["proc"] else 0
    aptos=int(data[C["emo_estado"]].astype(str).str.upper().isin(["APTO","APTO C/R"]).sum()) if C["emo_estado"] else 0
    ingresados=int(data[C["ingreso"]].notna().sum()) if C["ingreso"] else 0
    ceses=int(data[C["cese"]].notna().sum()) if C["cese"] else 0
    return total,enproc,aptos,ingresados,ceses

def show_metrics(data):
    a,b,c,d,e=metricas(data)
    x=st.columns(5)
    x[0].metric("Postulantes",f"{a:,}")
    x[1].metric("En proceso",f"{b:,}")
    x[2].metric("Aptos médicos",f"{c:,}")
    x[3].metric("Ingresados",f"{d:,}")
    x[4].metric("Ceses",f"{e:,}")

def fmt_name(row):
    return f"{row.get(C['ape'],'') or ''} {row.get(C['nom'],'') or ''}".strip()

def record_selector(data, label="Seleccionar registro", key=None):
    if len(data)==0:
        st.warning("No hay registros con los filtros actuales.")
        return None
    options={}
    for _,r in data.iterrows():
        rid=int(r["record_id"])
        options[rid]=f"{rid} · {r.get(C['dni'],'')} · {fmt_name(r)} · {r.get(C['obra'],'')}"
    
    rid=st.selectbox(label, list(options), format_func=lambda x: options[x], key=key or f"selectbox_{label}")
    return rid

def date_soon(series, days=30):
    s=pd.to_datetime(series,errors="coerce")
    now=pd.Timestamp.today().normalize()
    return int(((s.notna())&(s<=now+pd.Timedelta(days=days))&(s>=now-pd.Timedelta(days=1))).sum())

def show_alerts(data):
    alerts=[]
    if C["emo_estado"]:
        n=int(data[C["emo_estado"]].astype(str).str.upper().isin(["NO SE PRESENTÓ","NO SE PRESENTO"]).sum())
        if n: alerts.append(("🔴","EMO no presentados",n,"danger"))
    if C["doc"]:
        n=int(data[C["doc"]].astype(str).str.upper().isin(["PENDIENTE","INCOMPLETA","OBSERVADO"]).sum())
        if n: alerts.append(("🟠","Documentos pendientes",n,"warning"))
    if C["retcc_venc"]:
        n=date_soon(data[C["retcc_venc"]])
        if n: alerts.append(("🟠","RETCC por vencer / vencido",n,"warning"))
    if C["fc_venc"]:
        n=date_soon(data[C["fc_venc"]])
        if n: alerts.append(("🟠","Fotocheck por vencer / vencido",n,"warning"))
    if not alerts: st.success("No hay alertas detectadas.")
    for icon,label,n,cls in alerts:
        st.markdown(f'<div style="padding:9px;border-bottom:1px solid #edf1f5">{icon} <b>{label}</b><span style="float:right" class="badge {cls}">{n}</span></div>',unsafe_allow_html=True)

# -------------------- INICIO --------------------
if st.session_state.page=="Inicio":
    show_metrics(view)
    l,m,r=st.columns([1.55,1,1.05])
    with l:
        st.markdown('<div class="card-title">Embudo de reclutamiento</div>',unsafe_allow_html=True)
        interviewed=int(view[C["entrevista"]].notna().sum()) if C["entrevista"] else 0
        emo=int(view[C["emo_fecha"]].notna().sum()) if C["emo_fecha"] else 0
        apt=int(view[C["emo_estado"]].astype(str).str.upper().isin(["APTO","APTO C/R"]).sum()) if C["emo_estado"] else 0
        ind=int(view[C["ind_fecha"]].notna().sum()) if C["ind_fecha"] else 0
        ing=int(view[C["ingreso"]].notna().sum()) if C["ingreso"] else 0
        vals=[len(view),interviewed,emo,apt,ind,ing]
        labels=["Postulantes","Entrevistados","EMO","Aptos","Inducción","Ingresados"]
        fig=go.Figure(go.Funnel(y=labels,x=vals,textinfo="value+percent initial"))
        fig.update_layout(height=390,margin=dict(l=10,r=10,t=10,b=10),paper_bgcolor="white")
        st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
    with m:
        st.markdown('<div class="card-title">Estado de procesos</div>',unsafe_allow_html=True)
        if C["proc"]:
            vc=view[C["proc"]].fillna("SIN ESTADO").astype(str).value_counts()
            fig=px.pie(values=vc.values,names=vc.index,hole=.67)
            fig.update_layout(height=330,margin=dict(l=5,r=5,t=5,b=5))
            st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
    with r:
        st.markdown('<div class="card-title">Alertas y pendientes</div>',unsafe_allow_html=True)
        show_alerts(view)

    a,b=st.columns([1.6,1])
    with a:
        st.markdown('<div class="card-title">Últimos ingresos</div>',unsafe_allow_html=True)
        if C["ingreso"]:
            tmp=view[view[C["ingreso"]].notna()].copy()
            tmp["_d"]=pd.to_datetime(tmp[C["ingreso"]],errors="coerce")
            tmp=tmp.sort_values("_d",ascending=False).head(8)
            cols=[c for c in ["record_id",C["dni"],C["nom"],C["ape"],C["esp"],C["obra"],C["ingreso"],C["post"]] if c]
            st.dataframe(tmp[cols],use_container_width=True,hide_index=True)
    with b:
        st.markdown('<div class="card-title">Estado por obra</div>',unsafe_allow_html=True)
        if C["obra"]:
            g=view.groupby(C["obra"]).size().reset_index(name="Total").sort_values("Total",ascending=False).head(10)
            if C["ingreso"]:
                ig=view[view[C["ingreso"]].notna()].groupby(C["obra"]).size().reset_index(name="Ingresados")
                g=g.merge(ig,on=C["obra"],how="left").fillna({"Ingresados":0})
                g["Cobertura %"]=(g["Ingresados"]/g["Total"]*100).round(1)
            st.dataframe(g,use_container_width=True,hide_index=True)

# -------------------- REQUERIMIENTOS --------------------
elif st.session_state.page=="Requerimientos":
    show_metrics(view)
    st.subheader("Requerimientos de personal")
    if C["req"]:
        cols=[c for c in [C["req"],C["obra"],find_col(["Fecha_Requerimiento"]),find_col(["EstadoRequerimiento"]),
                          C["sol"],find_col(["Cant_Pedido"]),find_col(["Posición_Pedido"]),C["recl"]] if c]
        g=view.groupby([c for c in [C["req"],C["obra"],C["sol"],C["recl"]] if c],dropna=False).size().reset_index(name="Postulantes")
        st.dataframe(g,use_container_width=True,hide_index=True)
        st.caption("La agrupación utiliza el NroRequerimiento existente en la base; no se inventan requerimientos nuevos.")
    else: st.info("No existe NroRequerimiento en la base.")

# -------------------- RECLUTAMIENTO --------------------
elif st.session_state.page=="Reclutamiento":
    show_metrics(view)
    tabs=st.tabs(["📋 Postulantes","➕ Nuevo postulante","✏️ Editar ficha","🕘 Historial"])
    with tabs[0]:
        st.dataframe(view,use_container_width=True,hide_index=True)
    with tabs[1]:
        key_fields=[x for x in [C["dni"],C["ape"],C["nom"],C["obra"],C["req"],C["esp"],C["recl"],C["sol"],find_col(["Fecha_Requerimiento"])] if x]
        with st.form("new_record"):
            vals={}
            cc=st.columns(3)
            for i,c in enumerate(key_fields):
                vals[c]=cc[i%3].text_input(str(c))
            save=st.form_submit_button("💾 Crear postulante",use_container_width=True)
        if save:
            rid=insert_record(vals,user); st.success(f"Postulante creado. Registro #{rid}"); st.rerun()
    with tabs[2]:
        rid=record_selector(view)
        if rid:
            old=fetch_record(rid)
            groups={
             "Requerimiento":[C["obra"],C["req"],find_col(["Fecha_Requerimiento"]),find_col(["EstadoRequerimiento"]),C["sol"],find_col(["Cant_Pedido"]),C["recl"]],
             "Datos personales":[C["dni"],C["ape"],C["nom"],C["esp"],find_col(["Fecha_Nacimiento"]),find_col(["Móvil_01"]),find_col(["Móvil_02"]),find_col(["Mail"]),find_col(["Dirección"]),find_col(["Sexo"]),find_col(["EstadoCivil_2Des","Estado_Civil"]),find_col(["Grado_de_Instrucción"])],
             "Proceso":[C["entrevista"],C["emo_fecha"],C["emo_estado"],C["ind_fecha"],C["ind_estado"],C["doc"],C["ingreso"],C["cese"],C["post"],C["proc"]],
             "Movilización y fotocheck":[C["fc_venc"],find_col(["Fecha_Solicitud_Fotocheck"]),find_col(["Fecha_Entrega_Fotocheck"]),find_col(["Fecha_Movilización"]),find_col(["Lugar_de_Embarque"]),find_col(["Procedencia"]),find_col(["Comentario_de_Movilización"])],
            }
            vals={}
            for title,cols in groups.items():
                cols=[c for c in cols if c and c in old]
                if not cols: continue
                with st.expander(title,expanded=(title in ["Requerimiento","Datos personales","Proceso"])):
                    cc=st.columns(2)
                    for i,c in enumerate(cols):
                        vals[c]=cc[i%2].text_input(str(c),value="" if old[c] is None else str(old[c]),key=f"edit_{rid}_{c}")
            if st.button("💾 Guardar cambios",type="primary"):
                update_record(rid,vals,user); st.success("Ficha actualizada y registrada en auditoría."); st.rerun()
            if role=="admin":
                if st.button("🗑️ Eliminar registro",type="secondary"):
                    delete_record(rid,user); st.warning("Registro eliminado."); st.rerun()
    with tabs[3]:
        rid=record_selector(view, key="record_selector_tab_3")
        if rid:
            conn=get_conn()
            hist=pd.read_sql_query("SELECT created_at,username,action,changes FROM audit_log WHERE record_id=?", conn, params=(rid,))
            conn.close()
            st.dataframe(hist,use_container_width=True,hide_index=True)

# -------------------- PERSONAL --------------------
elif st.session_state.page=="Personal":
    show_metrics(view)
    st.dataframe(view,use_container_width=True,hide_index=True)
    st.download_button("⬇️ Exportar vista actual",view.to_csv(index=False).encode("utf-8-sig"),"personal_filtrado.csv","text/csv")

# -------------------- MÓDULOS DINÁMICOS --------------------
elif st.session_state.page in ["Seguridad","Capacitaciones","Evaluaciones","Documentos","EPP","Fotocheck","Movilización","Ingresos / Ceses"]:
    show_metrics(view)
    title=st.session_state.page
    keyword_map={
      "Seguridad":["seguridad","riesgo","iperc","loto","emergencia","incend","altura","izaje","confinado","excav","eléctr","electr","vigía","vigia"],
      "Capacitaciones":["_FE","_ES","_CO","capacit"],
      "Evaluaciones":["eva_","evalu","licencia","cert_","operador"],
      "Documentos":["anexo","document","retcc","sctr","pgI","webcontrol"],
      "EPP":["epp","talla"],
      "Fotocheck":["fotocheck"],
      "Movilización":["movilización","movilizacion","embarque","procedencia"],
      "Ingresos / Ceses":["ingreso","cese","reemplaz","estado_postulante","estado_del_proceso"],
    }
    keys=keyword_map[title]
    cols=[c for c in df.columns if any(k.lower() in str(c).lower() for k in keys)]
    cols=[c for c in ["record_id",C["dni"],C["nom"],C["ape"],C["obra"]]+cols if c]
    cols=list(dict.fromkeys(cols))
    if cols:
        st.caption(f"{len(cols)} campos relacionados detectados automáticamente en la base.")
        st.dataframe(view[cols],use_container_width=True,hide_index=True)
    else: st.info("No se encontraron campos relacionados.")

# -------------------- REPORTES --------------------
elif st.session_state.page=="Reportes":
    show_metrics(view)
    st.subheader("Reportería gerencial")
    a,b=st.columns(2)
    with a:
        if C["obra"]:
            g=view.groupby(C["obra"]).size().reset_index(name="Postulantes").sort_values("Postulantes",ascending=False)
            fig=px.bar(g.head(20),x=C["obra"],y="Postulantes",title="Postulantes por obra")
            fig.update_layout(xaxis_tickangle=-35)
            st.plotly_chart(fig,use_container_width=True)
    with b:
        if C["esp"]:
            g=view.groupby(C["esp"]).size().reset_index(name="Postulantes").sort_values("Postulantes",ascending=False).head(15)
            fig=px.bar(g,x="Postulantes",y=C["esp"],orientation="h",title="Postulantes por especialidad")
            st.plotly_chart(fig,use_container_width=True)
    xbuf=io.BytesIO()
    with pd.ExcelWriter(xbuf,engine="openpyxl") as writer:
        view.to_excel(writer,index=False,sheet_name="Reporte")
    st.download_button("⬇️ Descargar reporte Excel",xbuf.getvalue(),"reporte_gerencial.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------------------- ADMIN --------------------
elif st.session_state.page=="Administración":
    if role!="admin":
        st.error("Solo el administrador puede acceder.")
        st.stop()
    st.subheader("Administración")
    tabs=st.tabs(["👤 Usuarios","📥 Migración / respaldo","🧾 Auditoría"])
    with tabs[0]:
        conn=get_conn()
        users=pd.read_sql_query("SELECT id,username,role,active,created_at FROM app_users",conn)
        conn.close()
        st.dataframe(users,use_container_width=True,hide_index=True)
        with st.form("new_user"):
            nu=st.text_input("Nuevo usuario")
            npw=st.text_input("Contraseña",type="password")
            nr=st.selectbox("Rol",["admin","reclutador","consulta"])
            create=st.form_submit_button("Crear usuario")
        if create:
            if not nu or not npw: st.error("Completa usuario y contraseña.")
            else:
                try:
                    execute("INSERT INTO app_users(username,password_hash,role,created_at) VALUES(?,?,?,?)",(nu,hash_password(npw),nr,datetime.now().isoformat(timespec="seconds")))
                    st.success("Usuario creado."); st.rerun()
                except sqlite3.IntegrityError: st.error("Ese usuario ya existe.")
        st.info("Para cambiar la contraseña del administrador, crea un usuario nuevo con rol admin y luego desactiva el anterior. En una siguiente versión se puede añadir un formulario de cambio de contraseña.")
    with tabs[1]:
        st.warning("La migración reemplaza la tabla operativa por el Excel seleccionado. Primero descarga un respaldo.")
        conn=get_conn()
        raw=pd.read_sql_query("SELECT * FROM personal",conn)
        conn.close()
        buf=io.BytesIO()
        with pd.ExcelWriter(buf,engine="openpyxl") as writer: raw.to_excel(writer,index=False,sheet_name="Datos")
        st.download_button("⬇️ Descargar respaldo actual",buf.getvalue(),"respaldo_personal.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        up=st.file_uploader("Excel para migrar",type=["xlsx"],key="migration")
        confirm=st.checkbox("Confirmo que deseo reemplazar la base operativa por este Excel.")
        if up and confirm and st.button("⚠️ Ejecutar migración"):
            nrows,ncols,sheet=migrate_excel(up.getvalue(),user)
            st.success(f"Migración terminada: {nrows} filas, {ncols} campos."); st.rerun()
    with tabs[2]:
        conn=get_conn()
        audit=pd.read_sql_query("SELECT * FROM audit_log ORDER BY id DESC LIMIT 500",conn)
        conn.close()
        st.dataframe(audit,use_container_width=True,hide_index=True)

st.markdown("---")
st.caption(f"Base local: {DB_PATH.name} · {len(df):,} registros · {len(df.columns)-1} campos del Excel + record_id · Última lectura dinámica.")
