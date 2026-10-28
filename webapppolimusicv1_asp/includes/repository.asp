<%
'==============================================================================
' PoliMusic - Repositorio de consultas (SOLO LECTURA)
'==============================================================================
' Cada funcion devuelve un Recordset desconectado. Las paginas lo recorren y
' lo cierran; nunca se deja una conexion abierta durante el render.
'==============================================================================

'==============================================================================
' Columnas estandar de una cancion, con su artista principal y su album.
' Se usa un OUTER APPLY para traer solo un artista/album por cancion y evitar
' duplicar filas cuando una cancion pertenece a varias entidades.
'==============================================================================
Function SongSelectColumns()
    SongSelectColumns = _
        "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " & _
        "  (SELECT TOP 1 a.ID_ARTIST FROM TBL_PLAYER p " & _
        "     INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
        "     WHERE p.ID_SONG = s.ID_SONG ORDER BY a.ID_ARTIST) AS ID_ARTIST, " & _
        "  (SELECT TOP 1 a.ARTIST_NAME FROM TBL_PLAYER p " & _
        "     INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
        "     WHERE p.ID_SONG = s.ID_SONG ORDER BY a.ID_ARTIST) AS ARTIST_NAME, " & _
        "  (SELECT TOP 1 a.ARTIST_LASTNAME FROM TBL_PLAYER p " & _
        "     INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
        "     WHERE p.ID_SONG = s.ID_SONG ORDER BY a.ID_ARTIST) AS ARTIST_LASTNAME, " & _
        "  (SELECT TOP 1 al.ID_ALBUM FROM TBL_TRACKLIST t " & _
        "     INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
        "     WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ID_ALBUM, " & _
        "  (SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t " & _
        "     INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
        "     WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS TITLE_ALBUM, " & _
        "  (SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t " & _
        "     INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
        "     WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ALBUM_IMAGE_PATH " & _
        "FROM TBL_SONG s "
End Function

'==============================================================================
' SongsAll() - Todas las canciones, mas reproducidas primero.
'==============================================================================
Function SongsAll(limitCount)
    Dim sql, n
    n = CLng(limitCount)

    If n > 0 Then
        sql = "SELECT TOP " & n & " * FROM (" & SongSelectColumns() & ") q " & _
              "ORDER BY q.PLAYS DESC, q.SONG_NAME ASC"
    Else
        sql = "SELECT * FROM (" & SongSelectColumns() & ") q " & _
              "ORDER BY q.PLAYS DESC, q.SONG_NAME ASC"
    End If

    Set SongsAll = ExecuteQuery(sql, Empty)
End Function

'==============================================================================
' SongsTop() - Las N canciones mas reproducidas (para la portada).
'==============================================================================
Function SongsTop(limitCount)
    Dim n
    n = CLng(limitCount)
    If n <= 0 Then n = 5
    Set SongsTop = ExecuteQuery( _
        "SELECT TOP " & n & " * FROM (" & SongSelectColumns() & ") q " & _
        "ORDER BY q.PLAYS DESC, q.SONG_NAME ASC", Empty)
End Function

'==============================================================================
' SongById() - Una cancion concreta.
'==============================================================================
Function SongById(songId)
    Set SongById = ExecuteQuery( _
        "SELECT * FROM (" & SongSelectColumns() & ") q WHERE q.ID_SONG = ?", _
        Array(CLng(songId)))
End Function

'==============================================================================
' SongsByAlbum() - Canciones de un album.
'
' SongSelectColumns() genera una subconsulta con alias "s", por lo que las
' columnas de la tabla puente (ID_ALBUM, ID_PLAYLIST, ...) no se exponen.
' Por eso el filtro se hace con un IN sobre la tabla puente, no con un JOIN
' dentro del SELECT interno.
'==============================================================================
Function SongsByAlbum(albumId)
    Dim sql
    sql = "SELECT q.* FROM (" & SongSelectColumns() & ") q " & _
          "WHERE q.ID_SONG IN (SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = ?) " & _
          "ORDER BY q.SONG_NAME"
    Set SongsByAlbum = ExecuteQuery(sql, Array(CLng(albumId)))
End Function

'==============================================================================
' SongsByArtist() - Canciones de un artista.
'==============================================================================
Function SongsByArtist(artistId)
    Dim sql
    sql = "SELECT q.* FROM (" & SongSelectColumns() & ") q " & _
          "WHERE q.ID_SONG IN (SELECT ID_SONG FROM TBL_PLAYER WHERE ID_ARTIST = ?) " & _
          "ORDER BY q.SONG_NAME"
    Set SongsByArtist = ExecuteQuery(sql, Array(CLng(artistId)))
End Function

'==============================================================================
' SongsByPlaylist() - Canciones de una playlist.
'==============================================================================
Function SongsByPlaylist(playlistId)
    Dim sql
    sql = "SELECT q.* FROM (" & SongSelectColumns() & ") q " & _
          "WHERE q.ID_SONG IN (SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?) " & _
          "ORDER BY q.SONG_NAME"
    Set SongsByPlaylist = ExecuteQuery(sql, Array(CLng(playlistId)))
End Function

'==============================================================================
' SearchAll() - Busca canciones, albumes y artistas por nombre.
'==============================================================================
Function SearchAll(term)
    Set SearchAll = ExecuteQuery( _
        "SELECT * FROM (" & SongSelectColumns() & ") q " & _
        "WHERE q.SONG_NAME LIKE ? OR q.ARTIST_NAME LIKE ? OR q.TITLE_ALBUM LIKE ? " & _
        "ORDER BY q.SONG_NAME", _
        Array("%" & term & "%", "%" & term & "%", "%" & term & "%"))
End Function

'==============================================================================
' SongIncrementPlays() - Suma una reproduccion a la cancion.
'==============================================================================
Function SongIncrementPlays(songId)
    SongIncrementPlays = ExecuteNonQuery( _
        "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS,0) + 1 WHERE ID_SONG = ?", _
        Array(CLng(songId)))
End Function

'==============================================================================
' AlbumsAll() - Catalogo completo de albumes con numero de canciones.
'==============================================================================
Function AlbumsAll()
    Set AlbumsAll = ExecuteQuery( _
        "SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
        "  (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONG_COUNT " & _
        "FROM TBL_ALBUM al ORDER BY al.TITLE_ALBUM", Empty)
End Function

'==============================================================================
' AlbumById() - Un album concreto.
'==============================================================================
Function AlbumById(albumId)
    Set AlbumById = ExecuteQuery( _
        "SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
        "  (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONG_COUNT " & _
        "FROM TBL_ALBUM al WHERE al.ID_ALBUM = ?", _
        Array(CLng(albumId)))
End Function

'==============================================================================
' ArtistsAll() - Catalogo completo de artistas con numero de canciones.
'==============================================================================
Function ArtistsAll()
    Set ArtistsAll = ExecuteQuery( _
        "SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " & _
        "  (SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS SONG_COUNT, " & _
        "  (SELECT ISNULL(SUM(ISNULL(s.PLAYS,0)),0) FROM TBL_PLAYER p " & _
        "      INNER JOIN TBL_SONG s ON s.ID_SONG = p.ID_SONG " & _
        "      WHERE p.ID_ARTIST = ar.ID_ARTIST) AS TOTAL_PLAYS " & _
        "FROM TBL_ARTIST ar ORDER BY ar.ARTIST_NAME, ar.ARTIST_LASTNAME", Empty)
End Function

'==============================================================================
' ArtistById() - Un artista concreto.
'==============================================================================
Function ArtistById(artistId)
    Set ArtistById = ExecuteQuery( _
        "SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " & _
        "  (SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS SONG_COUNT, " & _
        "  (SELECT ISNULL(SUM(ISNULL(s.PLAYS,0)),0) FROM TBL_PLAYER p " & _
        "      INNER JOIN TBL_SONG s ON s.ID_SONG = p.ID_SONG " & _
        "      WHERE p.ID_ARTIST = ar.ID_ARTIST) AS TOTAL_PLAYS " & _
        "FROM TBL_ARTIST ar WHERE ar.ID_ARTIST = ?", _
        Array(CLng(artistId)))
End Function

'==============================================================================
' PlaylistsAll() - Todas las playlists con su dueño y numero de canciones.
'==============================================================================
Function PlaylistsAll()
    Set PlaylistsAll = ExecuteQuery(PlaylistSelectSql() & "ORDER BY pl.TITLE", Empty)
End Function

'==============================================================================
' PlaylistsByUser() - Playlists de un usuario.
'==============================================================================
Function PlaylistsByUser(userId)
    Set PlaylistsByUser = ExecuteQuery( _
        PlaylistSelectSql() & "WHERE pl.ID_USER = ? ORDER BY pl.TITLE", _
        Array(CLng(userId)))
End Function

'==============================================================================
' PlaylistsVisible() - Playlists publicas + las privadas del propio usuario.
'==============================================================================
Function PlaylistsVisible(userId)
    Set PlaylistsVisible = ExecuteQuery( _
        PlaylistSelectSql() & "WHERE pl.TYPE = " & PLAYLIST_PUBLIC & " OR pl.ID_USER = ? " & _
        "ORDER BY pl.TITLE", _
        Array(CLng(userId)))
End Function

'==============================================================================
' PlaylistById() - Una playlist concreta.
'==============================================================================
Function PlaylistById(playlistId)
    Set PlaylistById = ExecuteQuery( _
        PlaylistSelectSql() & "WHERE pl.ID_PLAYLIST = ?", _
        Array(CLng(playlistId)))
End Function

'==============================================================================
' PlaylistSelectSql() - SELECT comun de playlists.
'==============================================================================
Function PlaylistSelectSql()
    PlaylistSelectSql = _
        "SELECT pl.ID_PLAYLIST, pl.ID_USER, pl.TITLE, pl.CREATION_DATE, pl.TYPE, pl.PHOTO, " & _
        "  u.USERNAME, " & _
        "  (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG x WHERE x.ID_PLAYLIST = pl.ID_PLAYLIST) AS SONG_COUNT " & _
        "FROM TBL_PLAYLIST pl " & _
        "INNER JOIN TBL_USER u ON u.ID_USER = pl.ID_USER "
End Function

'==============================================================================
' UsersAll() - Listado de usuarios (solo administradores).
'==============================================================================
Function UsersAll()
    Set UsersAll = ExecuteQuery( _
        "SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE, " & _
        "  (SELECT COUNT(*) FROM TBL_PLAYLIST p WHERE p.ID_USER = u.ID_USER) AS PLAYLIST_COUNT " & _
        "FROM TBL_USER u ORDER BY USER_TYPE DESC, USERNAME", Empty)
End Function

'==============================================================================
' UserById() - Un usuario concreto.
'==============================================================================
Function UserById(userId)
    Set UserById = ExecuteQuery( _
        "SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE " & _
        "FROM TBL_USER WHERE ID_USER = ?", _
        Array(CLng(userId)))
End Function

'==============================================================================
' UserByUsername() - Busqueda por nombre de usuario.
'==============================================================================
Function UserByUsername(userName)
    Set UserByUsername = ExecuteQuery( _
        "SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE " & _
        "FROM TBL_USER WHERE USERNAME = ?", _
        Array(CStr(userName)))
End Function

'==============================================================================
' UsernameTaken() - True si el nombre ya existe (excluyendo un ID concreto).
'==============================================================================
Function UsernameTaken(userName, excludeId)
    Dim n
    n = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ? AND ID_USER <> ?", _
        Array(CStr(userName), CLng(excludeId))))
    UsernameTaken = (n > 0)
End Function

'==============================================================================
' EmailTaken() - True si el correo ya existe (excluyendo un ID concreto).
'==============================================================================
Function EmailTaken(email, excludeId)
    Dim n
    n = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", _
        Array(CStr(email), CLng(excludeId))))
    EmailTaken = (n > 0)
End Function

'==============================================================================
' CatalogLabel() - Etiqueta legible desde TBL_CATALOG.
'==============================================================================
Function CatalogLabel(parentCode, code)
    Dim v
    v = ExecuteScalar( _
        "SELECT TOP 1 VALUE FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = ? AND COD_CATALOG = ?", _
        Array(CStr(parentCode), CStr(code)))
    If IsNull(v) Then
        CatalogLabel = CStr(code)
    Else
        CatalogLabel = CStr(v)
    End If
End Function

'==============================================================================
' CountRows() - Cuenta filas de una tabla con filtro opcional.
' Solo se usa con nombres de tabla internos, nunca con entrada del usuario.
'==============================================================================
Function CountRows(tableName, whereClause)
    Dim sql
    sql = "SELECT COUNT(*) FROM " & tableName
    If Len(Trim(CStr(whereClause & ""))) > 0 Then
        sql = sql & " WHERE " & whereClause
    End If
    CountRows = CLng(ExecuteScalar(sql, Empty))
End Function

'==============================================================================
' TotalPlays() - Suma total de reproducciones del catalogo.
'==============================================================================
Function TotalPlays()
    Dim v
    v = ExecuteScalar("SELECT ISNULL(SUM(ISNULL(PLAYS,0)),0) FROM TBL_SONG", Empty)
    If IsNull(v) Then
        TotalPlays = 0
    Else
        TotalPlays = CLng(v)
    End If
End Function

'==============================================================================
' ArtistNameFull() - Nombre completo de un artista del Recordset actual.
'==============================================================================
Function ArtistNameFull(rs)
    Dim n, l
    n = CStr(rs("ARTIST_NAME") & "")
    l = CStr(rs("ARTIST_LASTNAME") & "")
    If Len(l) > 0 Then
        ArtistNameFull = n & " " & l
    Else
        ArtistNameFull = n
    End If
End Function
%>
