<%
' =====================================================================
'  PoliMusic - Configuración global
'  Este archivo vive en /includes, carpeta bloqueada para el navegador
'  mediante requestFiltering/hiddenSegments (ver web.config).
' =====================================================================

Const APP_NAME = "PoliMusic"
Const APP_VERSION = "1.0.0"

' Cadena equivalente a la de la app de consola:
'   Server=localhost;Database=BDD_PoliMusic;User Id=usr_polimusic;Password=Politecnica1;Encrypt=true;TrustServerCertificate=true
' ADO clásico requiere un proveedor OLE DB: Microsoft OLE DB Driver 19 for SQL Server (MSOLEDBSQL19).
Const DB_CONN = "Provider=MSOLEDBSQL19;Data Source=localhost;Initial Catalog=BDD_PoliMusic;User ID=usr_polimusic;Password=Politecnica1;Use Encryption for Data=Mandatory;Trust Server Certificate=True;Application Name=PoliMusicWeb"

' Roles (TBL_CATALOG, COD_CATALOG_PARENT = '1')
Const ROLE_USER = 0
Const ROLE_ADMIN = 1

' Tipos de playlist (TBL_CATALOG, COD_CATALOG_PARENT = '2')
Const PLAYLIST_PUBLIC = 0
Const PLAYLIST_PRIVATE = 1

Const SESSION_MINUTES = 30
Const LOGIN_MAX_ATTEMPTS = 5
Const LOGIN_LOCK_MINUTES = 2

Const MAX_UPLOAD_BYTES = 52428800   ' 50 MB (audio)
Const MAX_IMAGE_BYTES = 5242880     ' 5 MB

Const DEFAULT_USER_PHOTO = "~/userImages/default.jpg"
Const DEFAULT_PLAYLIST_PHOTO = "~/playlistImages/default.jpg"

Const AUDIO_EXTS = ".mp3,.m4a,.ogg,.wav"
Const IMAGE_EXTS = ".jpg,.jpeg,.png,.webp,.gif"

Response.CodePage = 65001
Response.Charset = "utf-8"
Session.Timeout = SESSION_MINUTES
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Response.AddHeader "Pragma", "no-cache"
%>
