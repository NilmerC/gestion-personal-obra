# Plataforma de Gestión de Personal de Obra — puesta en marcha

## Qué es esta versión

Es una versión operativa para comenzar a trabajar hoy con el Excel real.

- Conserva los 266 campos del Excel.
- Migra inicialmente los 1,373 registros a SQLite.
- El Excel deja de ser la fuente operativa una vez migrado.
- Los cambios hechos desde la plataforma quedan en la base.
- Tiene login, roles, CRUD de fichas, auditoría, filtros, dashboard, alertas y reportes.
- La apariencia está basada en el mockup generado para el proyecto.

## Requisito

Windows con Python 3.11 o superior.

## Inicio rápido

1. Descomprime este ZIP.
2. Mantén `202691814513JJC27174 (1).xlsx` en la misma carpeta de `app.py`.
3. Haz doble clic en `INICIAR_PLATAFORMA.bat`.
4. La primera ejecución instala dependencias y migra el Excel a `reclutamiento.db`.
5. Se abrirá Streamlit en el navegador.

## Acceso inicial

Usuario: `admin`
Contraseña: `admin123`

Cámbialo antes de usar el sistema con información real. La administración permite crear usuarios con roles `admin`, `reclutador` y `consulta`.

## Importante sobre la migración

La primera ejecución crea una base local a partir del Excel.

Después de esa primera migración, la plataforma trabaja sobre SQLite. Los cambios que hagas en la plataforma NO se escriben automáticamente de vuelta al Excel original.

Usa Administración > Migración / respaldo para:
- descargar un respaldo;
- importar una nueva versión del Excel si necesitas reemplazar la base.

## Qué puedes hacer hoy

### Inicio
- KPIs dinámicos.
- Embudo.
- Estado de procesos.
- Alertas.
- Últimos ingresos.
- Estado por obra.

### Requerimientos
- Agrupación por NroRequerimiento, obra, solicitante y reclutador.

### Reclutamiento
- Consultar postulantes.
- Crear postulante.
- Editar ficha.
- Registrar auditoría de cambios.
- Eliminar (solo admin).

### Personal
- Consultar toda la base.
- Exportar la vista filtrada.

### Seguridad / Capacitaciones / Evaluaciones / Documentos / EPP / Fotocheck / Movilización / Ingresos-Ceses
- Vista dinámica de los campos existentes relacionados con cada módulo.

### Reportes
- Gráficos por obra y especialidad.
- Exportación Excel.

## Siguiente fase

Esta versión es la base operativa. Para una instalación corporativa multiusuario se recomienda evolucionar a:

- PostgreSQL.
- Django.
- almacenamiento de documentos.
- permisos por obra/proyecto.
- workflow formal de estados.
- notificaciones por correo/WhatsApp/Teams según integración.
- Celery/Redis para tareas programadas.
- historial de cada etapa como entidades propias.
- documentos adjuntos.
- backups automáticos.
- despliegue en servidor/cloud.

La estructura actual evita perder los 266 campos mientras se realiza esa evolución.
