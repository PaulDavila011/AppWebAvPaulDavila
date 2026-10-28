<%
'==============================================================================
' PoliMusic - Operaciones de escritura (CRUD)
'==============================================================================
' Cada funcion devuelve el ID afectado o el numero de filas. Todas usan
' parametros enlazados. Las validaciones de negocio viven en las paginas; aqui
' solo se ejecuta la sentencia.
'==============================================================================

'==============================================================================
' CANCIONES
'==============================================================================
Function SongCreate(songName, songPath)
    SongCreate = ExecuteInsert( _
        "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, 0)", _
        Array(CStr(songName), CStr(songPath)))
End Function

Function SongUpdate(songId, songName, songPath)
    SongUpdate = ExecuteNonQuery( _
        "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ? WHERE ID_SONG = ?", _
        Array(CStr(songName), CStr(songPath), CLng(songId)))
End Function

'==============================================================================
' SongDelete() - Borra la cancion y todas sus relaciones (evita violar FKs).
'==============================================================================
Sub SongDelete(songId)
    Dim id
    id = CLng(songId)

    ExecuteNonQuery "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", Array(id)
    ExecuteNonQuery "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(id)
    ExecuteNonQuery "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(id)
    ExecuteNonQuery "DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array(id)
End Sub

'==============================================================================
' SongSetArtist() - Reemplaza la relacion cancion<->artista.
'==============================================================================
Sub SongSetArtist(songId, artistId)
    ExecuteNonQuery "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(CLng(songId))
    If CLng(artistId) > 0 Then
        ExecuteNonQuery "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", _
            Array(CLng(songId), CLng(artistId))
    End If
End Sub

'==============================================================================
' SongSetAlbum() - Reemplaza la relacion cancion<->album.
'==============================================================================
Sub SongSetAlbum(songId, albumId)
    ExecuteNonQuery "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(CLng(songId))
    If CLng(albumId) > 0 Then
        ExecuteNonQuery "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", _
            Array(CLng(albumId), CLng(songId))
    End If
End Sub

'==============================================================================
' SongArtistId() - ID del artista principal de una cancion (0 si no tiene).
'==============================================================================
Function SongArtistId(songId)
    Dim v
    v = ExecuteScalar( _
        "SELECT TOP 1 ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ? ORDER BY ID_ARTIST", _
        Array(CLng(songId)))
    If IsNull(v) Then
        SongArtistId = 0
    Else
        SongArtistId = CLng(v)
    End If
End Function

'==============================================================================
' SongAlbumId() - ID del album principal de una cancion (0 si no tiene).
'==============================================================================
Function SongAlbumId(songId)
    Dim v
    v = ExecuteScalar( _
        "SELECT TOP 1 ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = ? ORDER BY ID_ALBUM", _
        Array(CLng(songId)))
    If IsNull(v) Then
        SongAlbumId = 0
    Else
        SongAlbumId = CLng(v)
    End If
End Function

'==============================================================================
' ARTISTAS
'==============================================================================
Function ArtistCreate(artistName, artistLastname, artistImage)
    ArtistCreate = ExecuteInsert( _
        "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?)", _
        Array(CStr(artistName), CStr(artistLastname), CStr(artistImage)))
End Function

Function ArtistUpdate(artistId, artistName, artistLastname, artistImage)
    ArtistUpdate = ExecuteNonQuery( _
        "UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? " & _
        "WHERE ID_ARTIST = ?", _
        Array(CStr(artistName), CStr(artistLastname), CStr(artistImage), CLng(artistId)))
End Function

'==============================================================================
' ArtistDelete() - Borra el artista y desvincula sus canciones.
'==============================================================================
Sub ArtistDelete(artistId)
    Dim id
    id = CLng(artistId)
    ExecuteNonQuery "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(id)
    ExecuteNonQuery "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(id)
End Sub

'==============================================================================
' ArtistSongCount() - Cuantas canciones tiene asociadas un artista.
'==============================================================================
Function ArtistSongCount(artistId)
    ArtistSongCount = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(CLng(artistId))))
End Function

'==============================================================================
' ALBUMES
'==============================================================================
Function AlbumCreate(titleAlbum, releaseYear, albumImagePath)
    AlbumCreate = ExecuteInsert( _
        "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, ?, ?)", _
        Array(CStr(titleAlbum), CLng(releaseYear), CStr(albumImagePath)))
End Function

Function AlbumUpdate(albumId, titleAlbum, releaseYear, albumImagePath)
    AlbumUpdate = ExecuteNonQuery( _
        "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? " & _
        "WHERE ID_ALBUM = ?", _
        Array(CStr(titleAlbum), CLng(releaseYear), CStr(albumImagePath), CLng(albumId)))
End Function

'==============================================================================
' AlbumDelete() - Borra el album y desvincula sus canciones.
'==============================================================================
Sub AlbumDelete(albumId)
    Dim id
    id = CLng(albumId)
    ExecuteNonQuery "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(id)
    ExecuteNonQuery "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(id)
End Sub

'==============================================================================
' AlbumSongCount() - Cuantas canciones contiene un album.
'==============================================================================
Function AlbumSongCount(albumId)
    AlbumSongCount = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(CLng(albumId))))
End Function

'==============================================================================
' PLAYLISTS
'==============================================================================
Function PlaylistCreate(userId, title, playlistType, photo)
    PlaylistCreate = ExecuteInsert( _
        "INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) " & _
        "VALUES (?, ?, ?, ?, ?)", _
        Array(CLng(userId), CStr(title), Now(), CLng(playlistType), CStr(photo)))
End Function

Function PlaylistUpdate(playlistId, title, playlistType, photo)
    PlaylistUpdate = ExecuteNonQuery( _
        "UPDATE TBL_PLAYLIST SET TITLE = ?, TYPE = ?, PHOTO = ? WHERE ID_PLAYLIST = ?", _
        Array(CStr(title), CLng(playlistType), CStr(photo), CLng(playlistId)))
End Function

'==============================================================================
' PlaylistDelete() - Borra la playlist y su contenido asociado.
'==============================================================================
Sub PlaylistDelete(playlistId)
    Dim id
    id = CLng(playlistId)
    ExecuteNonQuery "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(id)
    ExecuteNonQuery "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(id)
End Sub

'==============================================================================
' PlaylistAddSong() - Agrega una cancion si no estaba ya.
'==============================================================================
Function PlaylistAddSong(playlistId, songId)
    Dim exists
    exists = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", _
        Array(CLng(playlistId), CLng(songId))))

    If exists > 0 Then
        PlaylistAddSong = 0
        Exit Function
    End If

    PlaylistAddSong = ExecuteNonQuery( _
        "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)", _
        Array(CLng(playlistId), CLng(songId)))
End Function

'==============================================================================
' PlaylistRemoveSong() - Quita una cancion de la playlist.
'==============================================================================
Function PlaylistRemoveSong(playlistId, songId)
    PlaylistRemoveSong = ExecuteNonQuery( _
        "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", _
        Array(CLng(playlistId), CLng(songId)))
End Function

'==============================================================================
' PlaylistSongCount() - Cuantas canciones tiene la playlist.
'==============================================================================
Function PlaylistSongCount(playlistId)
    PlaylistSongCount = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", _
        Array(CLng(playlistId))))
End Function

'==============================================================================
' PlaylistOwnedBy() - True si la playlist pertenece al usuario indicado.
'==============================================================================
Function PlaylistOwnedBy(playlistId, userId)
    Dim n
    n = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ? AND ID_USER = ?", _
        Array(CLng(playlistId), CLng(userId))))
    PlaylistOwnedBy = (n > 0)
End Function

'==============================================================================
' PlaylistIsPublic() - True si la playlist es publica.
'==============================================================================
Function PlaylistIsPublic(playlistId)
    Dim v
    v = ExecuteScalar("SELECT TOP 1 TYPE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", _
        Array(CLng(playlistId)))
    If IsNull(v) Then
        PlaylistIsPublic = False
    Else
        PlaylistIsPublic = (CLng(v) = PLAYLIST_PUBLIC)
    End If
End Function

'==============================================================================
' USUARIOS
'==============================================================================
Function UserCreate(userName, plainPassword, email, birthday, photo, userType)
    Dim birthVal
    If IsBlank(birthday) Then
        birthVal = Null
    Else
        birthVal = CDate(birthday)
    End If

    Dim photoVal
    If IsBlank(photo) Then
        photoVal = "~/userImages/default.jpg"
    Else
        photoVal = CStr(photo)
    End If

    UserCreate = ExecuteInsert( _
        "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
        "VALUES (?, ?, ?, ?, ?, ?)", _
        Array(CStr(userName), CStr(plainPassword), CStr(email), birthVal, CStr(photoVal), CLng(userType)))
End Function

'==============================================================================
' UserUpdate() - Actualiza datos del usuario (y opcionalmente la contrasena).
'==============================================================================
Function UserUpdate(userId, userName, email, birthday, photo, userType, newPassword)
    Dim birthVal
    If IsBlank(birthday) Then
        birthVal = Null
    Else
        birthVal = CDate(birthday)
    End If

    If Len(CStr(newPassword & "")) > 0 Then
        UserUpdate = ExecuteNonQuery( _
            "UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, " & _
            "USER_TYPE = ?, PASSWORD = ? WHERE ID_USER = ?", _
            Array(CStr(userName), CStr(email), birthVal, CStr(photo), CLng(userType), _
                  CStr(newPassword), CLng(userId)))
    Else
        UserUpdate = ExecuteNonQuery( _
            "UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, " & _
            "USER_TYPE = ? WHERE ID_USER = ?", _
            Array(CStr(userName), CStr(email), birthVal, CStr(photo), CLng(userType), CLng(userId)))
    End If
End Function

'==============================================================================
' UserUpdatePhoto() - Cambia solo la foto de perfil.
'==============================================================================
Function UserUpdatePhoto(userId, photo)
    UserUpdatePhoto = ExecuteNonQuery( _
        "UPDATE TBL_USER SET USER_PHOTO = ? WHERE ID_USER = ?", _
        Array(CStr(photo), CLng(userId)))
End Function

'==============================================================================
' UserDelete() - Borra el usuario y todo lo que depende de el.
'==============================================================================
Sub UserDelete(userId)
    Dim id
    id = CLng(userId)

    ExecuteNonQuery "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN " & _
                    "(SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = ?)", Array(id)
    ExecuteNonQuery "DELETE FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(id)
    ExecuteNonQuery "DELETE FROM TBL_USER WHERE ID_USER = ?", Array(id)
End Sub

'==============================================================================
' AdminCount() - Cuantos administradores hay (para no quedarnos sin ninguno).
'==============================================================================
Function AdminCount()
    AdminCount = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = ?", Array(ROLE_ADMIN)))
End Function

'==============================================================================
' UserPlaylistCount() - Cuantas playlists tiene el usuario.
'==============================================================================
Function UserPlaylistCount(userId)
    UserPlaylistCount = CLng(ExecuteScalar( _
        "SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(CLng(userId))))
End Function
%>
