# PoliMusic — Reproductor de música web (ASP clásico)

Clon estilo Spotify construido con **ASP clásico (VBScript)**, **ADO** y **SQL Server** (`BDD_PoliMusic`).
Tema oscuro morado con contrastes WCAG AA, diseño responsive y reproductor persistente.

## Estructura

```
webapppolimusicv1_asp/
├── PoliMusic/                 ← raíz del sitio (lo que se publica en IIS)
│   ├── includes/              ← código de servidor (bloqueado al navegador)
│   │   ├── config.asp         ← cadena de conexión y constantes
│   │   ├── db.asp             ← ADO parametrizado + lotes transaccionales
│   │   ├── auth.asp           ← login, RBAC, CSRF
│   │   ├── form.asp           ← lector multipart (subida de archivos)
│   │   ├── ui.asp / utils.asp ← componentes y utilidades
│   │   └── layout_*.asp       ← estructura de página + reproductor
│   ├── assets/                ← CSS, JS, SVG
│   ├── login.asp, register.asp, logout.asp
│   ├── default.asp            ← pantalla principal
│   ├── songs/albums/album/artists/artist/playlists/playlist/search.asp
│   ├── play.asp               ← API: contador de reproducciones
│   ├── admin*.asp             ← CRUD (solo ADMIN)
│   └── songFiles/ albumImages/ ArtistImages/ userImages/ playlistImages/
├── database/01_seed_roles_usuarios.sql
├── deploy/deploy-iis.ps1
└── tools/run-sql.ps1
```

## Roles y permisos

| Funcionalidad | USER (0) | ADMIN (1) |
|---|:-:|:-:|
| Iniciar sesión / registrarse / cerrar sesión | ✔ | ✔ |
| Reproducir canciones, consultar álbumes y artistas | ✔ | ✔ |
| Ver playlists públicas y las propias | ✔ | ✔ (todas) |
| CRUD de Usuarios, Canciones, Artistas, Álbumes, Playlists | ✘ (HTTP 403) | ✔ |

- El rol se **revalida contra la BD en cada petición** (si un admin cambia el rol o elimina una cuenta, el efecto es inmediato).
- El registro público siempre crea usuarios con rol USER.
- Un admin no puede eliminarse ni quitarse el rol a sí mismo.

Cuentas: `userAdmin / 12345678` (ADMIN) y `paul.davila / 12345678` (USER).

## Seguridad

- Consultas 100 % parametrizadas (`ADODB.Command`) → sin inyección SQL.
- Contraseñas guardadas **sin cifrar** por ahora (texto plano, como en el script original). El login también acepta el hash SHA-1 que ya tiene `username1`.
- Token CSRF (generado con `CRYPT_GEN_RANDOM`) en todos los POST; cierre de sesión solo por POST.
- Bloqueo temporal tras 5 intentos fallidos de inicio de sesión.
- Escape HTML de toda salida; carpeta `includes/` oculta; las carpetas de medios solo sirven extensiones de audio o imagen (no se puede ejecutar un `.asp` subido).

## Despliegue en IIS

Requisitos: IIS con **ASP clásico**, **Microsoft OLE DB Driver 19 for SQL Server** (MSOLEDBSQL19) y SQL Server con `BDD_PoliMusic` creada.

```powershell
# PowerShell como Administrador, desde esta carpeta
powershell -ExecutionPolicy Bypass -File deploy\deploy-iis.ps1
```

Publica en `C:\inetpub\wwwroot\PoliMusic` → **http://localhost/PoliMusic/**

## Audio

Las canciones del script apuntan a `~/songFiles/bensound-*.mp3`. Copie esos MP3 a `PoliMusic/songFiles/` o súbalos desde **Administración → Canciones** (hasta 50 MB). El panel muestra qué canciones no tienen archivo.
