<%
'==============================================================================
' PoliMusic - Configuracion global
'==============================================================================
' Todos los valores sensibles de conexion viven aqui. Si necesitas apuntar a
' otra instancia de SQL Server, cambia unicamente este archivo.
'==============================================================================

' --- Servidor de base de datos ----------------------------------------------
Const DB_SERVER            = "localhost"
Const DB_NAME              = "BDD_PoliMusic"
Const DB_USER              = "usr_polimusic"
Const DB_PASSWORD          = "Politecnica1"
Const DB_ENCRYPT           = "Optional"
Const DB_TRUST_CERT        = "yes"

' --- Roles (TBL_CATALOG: COD_CATALOG_PARENT = 1) ----------------------------
Const ROLE_USER            = 0
Const ROLE_ADMIN           = 1

' --- Tipos de playlist (TBL_CATALOG: COD_CATALOG_PARENT = 2) ----------------
Const PLAYLIST_PUBLIC      = 0
Const PLAYLIST_PRIVATE     = 1

' --- Rutas de medios ---------------------------------------------------------
' La base de datos guarda rutas con el prefijo "~/" (estilo ASP.NET).
' Aqui se traducen a rutas web reales servidas por IIS.
Const MEDIA_PREFIX         = "~/"
Const MEDIA_BASE_URL       = "assets/media/"
Const MEDIA_ROOT_PREFIX    = "~/"

' --- Seguridad / sesion ------------------------------------------------------
Const MAX_LOGIN_ATTEMPTS   = 5
Const SESSION_TIMEOUT_MIN  = 30
Const UPLOAD_MAX_BYTES     = 10485760   ' 10 MB

' --- Etiquetas amigables -----------------------------------------------------
Const APP_NAME             = "PoliMusic"
Const APP_SUBTITLE         = "Escuela Politecnica Nacional"

'==============================================================================
' Devuelve la cadena de conexion principal (MSOLEDBSQL, driver moderno).
'==============================================================================
Function ConnectionString()
    Dim s
    s = "Provider=MSOLEDBSQL;" & _
        "Server=" & DB_SERVER & ";" & _
        "Database=" & DB_NAME & ";" & _
        "User Id=" & DB_USER & ";" & _
        "Password=" & DB_PASSWORD & ";" & _
        "Encrypt=" & DB_ENCRYPT & ";" & _
        "TrustServerCertificate=" & DB_TRUST_CERT & ";"
    ConnectionString = s
End Function

'==============================================================================
' Cadena de respaldo (SQLOLEDB, viene con Windows desde siempre).
' Se usa si MSOLEDBSQL no esta instalado en el servidor.
' Nota: SQLOLEDB NO acepta "MARS Connection=True"; si se agrega, la conexion
' falla con "El proveedor no admite la propiedad".
'==============================================================================
Function ConnectionStringFallback()
    Dim s
    s = "Provider=SQLOLEDB;" & _
        "Server=" & DB_SERVER & ";" & _
        "Database=" & DB_NAME & ";" & _
        "User Id=" & DB_USER & ";" & _
        "Password=" & DB_PASSWORD & ";"
    ConnectionStringFallback = s
End Function
%>
