/* ==========================================================================
   PoliMusic - Reproductor persistente
   --------------------------------------------------------------------------
   El reproductor vive en un elemento <audio> unico (#pmAudio) que esta
   presente en todas las paginas. Para que la musica no se corte al navegar,
   el estado se guarda en sessionStorage y se restaura al cargar la pagina
   siguiente: misma cancion, misma posicion, mismo volumen.

   API publica: window.PoliMusicPlayer
     load(tracks, startIndex)  Cola de reproduccion y arranque
     play(track)               Reproduce un objeto pista suelto
     toggle()                  Alterna reproducir/pausar
     getState()                Estado actual
   ========================================================================== */
(function () {
    'use strict';

    var STATE_KEY = 'pm_player_state';
    var COUNTED = 'pm_counted_plays';

    var audio = document.getElementById('pmAudio');
    if (!audio) return;

    /* --- Referencias del DOM ------------------------------------------- */
    var el = {
        cover:      document.getElementById('pmPlayerCover'),
        title:      document.getElementById('pmPlayerTitle'),
        artist:     document.getElementById('pmPlayerArtist'),
        toggle:     document.getElementById('pmToggle'),
        toggleIco:  document.getElementById('pmToggleIco'),
        prev:       document.getElementById('pmPrev'),
        next:       document.getElementById('pmNext'),
        shuffle:    document.getElementById('pmShuffle'),
        repeat:     document.getElementById('pmRepeat'),
        seek:       document.getElementById('pmSeek'),
        timeNow:    document.getElementById('pmTimeNow'),
        timeTotal:  document.getElementById('pmTimeTotal'),
        volume:     document.getElementById('pmVolume'),
        mute:       document.getElementById('pmMute'),
        volumeIco:  document.getElementById('pmVolumeIco')
    };

    /* --- Estado --------------------------------------------------------- */
    var state = {
        tracks:   [],
        index:    -1,
        playing:  false,
        shuffle:  false,
        repeat:   'off',          // 'off' | 'all' | 'one'
        volume:   0.8,
        muted:    false,
        position: 0
    };

    var seeking = false;
    var countedIds = {};

    /* --- Iconos (los mismos SVG que sirve el servidor) ------------------ */
    var ICONS = {
        play:  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true" class="pm-ico" width="24" height="24"><path d="M8 5.5v13a.6.6 0 0 0 .93.5l10-6.5a.6.6 0 0 0 0-1L8.93 5a.6.6 0 0 0-.93.5Z"/></svg>',
        pause: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true" class="pm-ico" width="24" height="24"><rect x="7" y="5" width="3.6" height="14" rx="1.1"/><rect x="13.4" y="5" width="3.6" height="14" rx="1.1"/></svg>',
        volume:'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" class="pm-ico" width="24" height="24"><path d="M4 9.5h3.2L12 5.4v13.2L7.2 14.5H4Z"/><path d="M16 9a4.5 4.5 0 0 1 0 6"/><path d="M18.6 6.4a8 8 0 0 1 0 11.2"/></svg>',
        mute:  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" class="pm-ico" width="24" height="24"><path d="M4 9.5h3.2L12 5.4v13.2L7.2 14.5H4Z"/><path d="M16.5 9.5l5 5"/><path d="M21.5 9.5l-5 5"/></svg>'
    };

    /* --- Utilidades ----------------------------------------------------- */
    function fmtTime(sec) {
        if (!isFinite(sec) || sec < 0) sec = 0;
        var m = Math.floor(sec / 60);
        var s = Math.floor(sec % 60);
        return m + ':' + (s < 10 ? '0' : '') + s;
    }

    function saveState() {
        try {
            sessionStorage.setItem(STATE_KEY, JSON.stringify({
                tracks:   state.tracks,
                index:    state.index,
                playing:  state.playing,
                shuffle:  state.shuffle,
                repeat:   state.repeat,
                volume:   state.volume,
                muted:    state.muted,
                position: audio.currentTime || 0
            }));
        } catch (e) { /* sessionStorage puede estar bloqueado */ }
    }

    function loadState() {
        try {
            var raw = sessionStorage.getItem(STATE_KEY);
            return raw ? JSON.parse(raw) : null;
        } catch (e) { return null; }
    }

    function post(url, data) {
        if (!url) return;
        try {
            var body = [];
            for (var k in data) {
                if (Object.prototype.hasOwnProperty.call(data, k)) {
                    body.push(encodeURIComponent(k) + '=' + encodeURIComponent(data[k]));
                }
            }
            if (window.fetch) {
                fetch(url, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: body.join('&'),
                    credentials: 'same-origin'
                }).catch(function () {});
            } else {
                var xhr = new XMLHttpRequest();
                xhr.open('POST', url, true);
                xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
                xhr.send(body.join('&'));
            }
        } catch (e) { /* silencioso: contar reproducciones no es critico */ }
    }

    /* --- API publica ---------------------------------------------------- */

    function currentTrack() {
        if (state.index < 0 || state.index >= state.tracks.length) return null;
        return state.tracks[state.index];
    }

    function renderTrack() {
        var t = currentTrack();
        if (!t) {
            if (el.title)  el.title.textContent = 'Nada en reproduccion';
            if (el.artist) el.artist.textContent = 'Elige una cancion del catalogo';
            if (el.cover)  el.cover.innerHTML = '<span class="pm-player-cover-fallback">' + ICONS.volume + '</span>';
            highlightRows();
            return;
        }

        if (el.title)  el.title.textContent = t.title || 'Sin titulo';
        if (el.artist) el.artist.textContent = t.artist || 'Artista desconocido';
        if (el.cover) {
            if (t.cover) {
                el.cover.innerHTML = '<img src="' + t.cover + '" alt="">';
            } else {
                el.cover.innerHTML = '<span class="pm-player-cover-fallback">' + ICONS.volume + '</span>';
            }
        }
        document.title = (t.title || 'PoliMusic') + ' - PoliMusic';
        highlightRows();
    }

    function highlightRows() {
        var t = currentTrack();
        var id = t ? String(t.id) : '';
        var rows = document.querySelectorAll('.pm-track-row[data-play]');
        for (var i = 0; i < rows.length; i++) {
            var idx = rows[i].getAttribute('data-play');
            var track = state.tracks[parseInt(idx, 10)];
            var match = track && String(track.id) === id;
            if (match) rows[i].classList.add('is-current');
            else rows[i].classList.remove('is-current');
        }
    }

    function renderPlayIcon() {
        if (!el.toggleIco) return;
        el.toggleIco.innerHTML = state.playing ? ICONS.pause : ICONS.play;
        if (el.toggle) {
            el.toggle.setAttribute('aria-label', state.playing ? 'Pausar' : 'Reproducir');
            if (state.playing) el.toggle.classList.add('is-playing');
            else el.toggle.classList.remove('is-playing');
        }
    }

    function countPlay(trackId) {
        if (!trackId) return;
        if (countedIds[trackId]) return;
        countedIds[trackId] = true;
        post('play_count.asp', {
            song: trackId,
            csrf: window.PM_CSRF || ''
        });
    }

    function setTrack(index, autoplay) {
        if (index < 0 || index >= state.tracks.length) return;
        state.index = index;
        var t = state.tracks[index];
        if (!t || !t.src) return;

        audio.src = t.src;
        audio.load();
        renderTrack();

        if (autoplay) {
            var p = audio.play();
            if (p && typeof p.catch === 'function') {
                p.catch(function () {
                    // El navegador bloqueo la reproduccion automatica
                    state.playing = false;
                    renderPlayIcon();
                    saveState();
                });
            }
        }
        saveState();
    }

    function play(track) {
        if (!track) return;
        // Si la pista ya esta en la cola, salta a ella
        for (var i = 0; i < state.tracks.length; i++) {
            if (String(state.tracks[i].id) === String(track.id) && state.tracks[i].src === track.src) {
                setTrack(i, true);
                return;
            }
        }
        state.tracks = [track];
        setTrack(0, true);
    }

    function load(tracks, startIndex) {
        if (!tracks || !tracks.length) return;
        state.tracks = tracks;
        startIndex = startIndex || 0;
        if (startIndex < 0) startIndex = 0;
        if (startIndex >= tracks.length) startIndex = 0;
        setTrack(startIndex, true);
    }

    function toggle() {
        if (!currentTrack()) {
            // Nada cargado pero hay filas en la pagina: empieza por la primera
            var first = document.querySelector('.pm-track-row[data-play="0"]');
            if (first) first.click();
            return;
        }
        if (audio.paused) {
            var p = audio.play();
            if (p && typeof p.catch === 'function') p.catch(function () {});
        } else {
            audio.pause();
        }
    }

    function next() {
        if (!state.tracks.length) return;

        if (state.repeat === 'one') {
            audio.currentTime = 0;
            var p = audio.play();
            if (p && typeof p.catch === 'function') p.catch(function () {});
            return;
        }

        if (state.shuffle && state.tracks.length > 1) {
            var r = state.index;
            while (r === state.index) {
                r = Math.floor(Math.random() * state.tracks.length);
            }
            setTrack(r, true);
            return;
        }

        if (state.index + 1 < state.tracks.length) {
            setTrack(state.index + 1, true);
        } else if (state.repeat === 'all') {
            setTrack(0, true);
        } else {
            state.playing = false;
            renderPlayIcon();
            saveState();
        }
    }

    function prev() {
        if (!state.tracks.length) return;
        // Reinicia si llevamos mas de 3 segundos
        if (audio.currentTime > 3) {
            audio.currentTime = 0;
            return;
        }
        if (state.index > 0) {
            setTrack(state.index - 1, true);
        } else {
            setTrack(state.tracks.length - 1, true);
        }
    }

    /* --- Sincronizacion con el elemento audio --------------------------- */
    audio.addEventListener('play', function () {
        state.playing = true;
        renderPlayIcon();
        saveState();
        var t = currentTrack();
        if (t) countPlay(t.id);
    });

    audio.addEventListener('pause', function () {
        state.playing = false;
        renderPlayIcon();
        saveState();
    });

    audio.addEventListener('ended', function () {
        saveState();
        next();
    });

    audio.addEventListener('timeupdate', function () {
        if (seeking) return;
        var d = audio.duration || 0;
        var t = audio.currentTime || 0;
        if (el.timeNow)   el.timeNow.textContent = fmtTime(t);
        if (el.timeTotal) el.timeTotal.textContent = fmtTime(d);
        if (el.seek && d > 0) el.seek.value = (t / d) * 100;
    });

    audio.addEventListener('loadedmetadata', function () {
        if (el.timeTotal) el.timeTotal.textContent = fmtTime(audio.duration);
    });

    audio.addEventListener('error', function () {
        state.playing = false;
        renderPlayIcon();
        if (el.artist) el.artist.textContent = 'No se pudo cargar el archivo de audio';
    });

    /* --- Controles ------------------------------------------------------ */
    if (el.toggle) el.toggle.addEventListener('click', toggle);
    if (el.next)   el.next.addEventListener('click', next);
    if (el.prev)   el.prev.addEventListener('click', prev);

    if (el.shuffle) {
        el.shuffle.addEventListener('click', function () {
            state.shuffle = !state.shuffle;
            el.shuffle.classList.toggle('is-active', state.shuffle);
            saveState();
        });
    }

    if (el.repeat) {
        el.repeat.addEventListener('click', function () {
            state.repeat = state.repeat === 'off' ? 'all'
                         : state.repeat === 'all' ? 'one' : 'off';
            el.repeat.classList.toggle('is-active', state.repeat !== 'off');
            el.repeat.classList.toggle('is-one', state.repeat === 'one');
            el.repeat.setAttribute('title',
                state.repeat === 'off' ? 'Repetir: desactivado'
                : state.repeat === 'all' ? 'Repetir: toda la cola'
                : 'Repetir: solo esta cancion');
            saveState();
        });
    }

    if (el.seek) {
        el.seek.addEventListener('input', function () {
            seeking = true;
            var d = audio.duration || 0;
            if (d > 0 && el.timeNow) {
                el.timeNow.textContent = fmtTime((el.seek.value / 100) * d);
            }
        });
        el.seek.addEventListener('change', function () {
            var d = audio.duration || 0;
            if (d > 0) audio.currentTime = (el.seek.value / 100) * d;
            seeking = false;
        });
    }

    if (el.volume) {
        el.volume.addEventListener('input', function () {
            state.volume = parseFloat(el.volume.value);
            state.muted = (state.volume === 0);
            audio.volume = state.volume;
            audio.muted = state.muted;
            renderVolumeIcon();
            saveState();
        });
    }

    function renderVolumeIcon() {
        if (!el.volumeIco) return;
        el.volumeIco.innerHTML = (state.muted || state.volume === 0) ? ICONS.mute : ICONS.volume;
        if (el.mute) el.mute.setAttribute('aria-label', state.muted ? 'Activar sonido' : 'Silenciar');
    }

    if (el.mute) {
        el.mute.addEventListener('click', function () {
            state.muted = !state.muted;
            audio.muted = state.muted;
            renderVolumeIcon();
            saveState();
        });
    }

    /* --- Delegacion de clics en las filas de canciones ------------------ */
    document.addEventListener('click', function (ev) {
        // Ignora clics sobre enlaces o botones dentro de la fila
        if (ev.target.closest && ev.target.closest('a, button, input, select')) return;

        var row = ev.target.closest ? ev.target.closest('.pm-track-row[data-play]') : null;
        if (!row) return;

        var idx = parseInt(row.getAttribute('data-play'), 10);
        var queue = window.PM_QUEUE || [];
        if (queue.length) {
            // Si la fila pertenece a la cola que la pagina expuso, reproduce esa cola
            if (row.getAttribute('data-queue') === 'page') {
                setTrack(idx, true);
                return;
            }
        }
        // Comportamiento por defecto: reproduce la cola embebida en la pagina
        if (window.PM_QUEUE && window.PM_QUEUE.length) {
            state.tracks = window.PM_QUEUE;
            setTrack(idx, true);
        }
    });

    document.addEventListener('keydown', function (ev) {
        var row = ev.target.closest ? ev.target.closest('.pm-track-row[data-play]') : null;
        if (!row) return;
        if (ev.key === 'Enter' || ev.key === ' ') {
            ev.preventDefault();
            row.click();
        }
    });

    /* Barra espaciadora: reproducir/pausar cuando no se escribe en un campo */
    document.addEventListener('keydown', function (ev) {
        if (ev.key !== ' ' && ev.code !== 'Space') return;
        var t = ev.target;
        if (!t) return;
        var tag = (t.tagName || '').toLowerCase();
        if (tag === 'input' || tag === 'textarea' || tag === 'select' || t.isContentEditable) return;
        if (t.closest && t.closest('.pm-track-row')) return;
        ev.preventDefault();
        toggle();
    });

    /* --- Restauracion del estado ---------------------------------------- */
    function restore() {
        var saved = loadState();
        if (!saved || !saved.tracks || !saved.tracks.length) {
            // Sin estado previo: intenta la cola embebida y precarga la primera
            if (window.PM_QUEUE && window.PM_QUEUE.length) {
                state.tracks = window.PM_QUEUE;
                state.index = 0;
                var t0 = state.tracks[0];
                if (t0 && t0.src) {
                    audio.src = t0.src;
                    renderTrack();
                }
            }
            return;
        }

        state.tracks  = saved.tracks;
        state.index   = typeof saved.index === 'number' ? saved.index : 0;
        state.shuffle = !!saved.shuffle;
        state.repeat  = saved.repeat || 'off';
        state.volume  = typeof saved.volume === 'number' ? saved.volume : 0.8;
        state.muted   = !!saved.muted;
        state.position = saved.position || 0;

        if (el.shuffle) el.shuffle.classList.toggle('is-active', state.shuffle);
        if (el.repeat) {
            el.repeat.classList.toggle('is-active', state.repeat !== 'off');
            el.repeat.classList.toggle('is-one', state.repeat === 'one');
        }
        if (el.volume) el.volume.value = state.volume;
        audio.volume = state.volume;
        audio.muted = state.muted;
        renderVolumeIcon();

        var t = currentTrack();
        if (!t || !t.src) return;

        audio.src = t.src;
        renderTrack();

        // Restaura la posicion de reproduccion
        if (state.position > 0) {
            var onMeta = function () {
                try { audio.currentTime = state.position; } catch (e) {}
                audio.removeEventListener('loadedmetadata', onMeta);
            };
            audio.addEventListener('loadedmetadata', onMeta);
        }

        // Si estaba sonando, reanuda al primer gesto del usuario
        if (state.playing) {
            var resume = function () {
                var p = audio.play();
                if (p && typeof p.catch === 'function') p.catch(function () {});
                document.removeEventListener('click', resume);
                document.removeEventListener('keydown', resume);
            };
            audio.play().then(function () {
                // Se pudo reanudar sin gesto
            }).catch(function () {
                document.addEventListener('click', resume);
                document.addEventListener('keydown', resume);
            });
        }
    }

    /* --- Guarda la posicion antes de abandonar la pagina ---------------- */
    window.addEventListener('beforeunload', saveState);
    document.addEventListener('visibilitychange', function () {
        if (document.visibilityState === 'hidden') saveState();
    });

    /* --- Exponer la API ------------------------------------------------- */
    window.PoliMusicPlayer = {
        load: load,
        play: play,
        toggle: toggle,
        next: next,
        prev: prev,
        countPlay: countPlay,
        getState: function () { return state; }
    };

    renderPlayIcon();
    restore();
    highlightRows();
})();

/* ==========================================================================
   Modal de confirmacion para eliminaciones
   Se usa desde los formularios generados por RenderDeleteForm().
   ========================================================================== */
var pmPendingForm = null;

function pmConfirmDelete(ev, label) {
    var modal = document.getElementById('pmConfirmModal');
    if (!modal) {
        return window.confirm('Seguro que quieres eliminar "' + label + '"? Esta accion no se puede deshacer.');
    }

    ev.preventDefault();
    pmPendingForm = ev.target;

    var text = document.getElementById('pmConfirmText');
    if (text) {
        text.textContent = 'Se eliminara "' + label + '" de forma permanente. Esta accion no se puede deshacer.';
    }

    modal.hidden = false;
    modal.setAttribute('aria-hidden', 'false');
    var ok = modal.querySelector('[data-pm-modal-ok]');
    if (ok) ok.focus();
    return false;
}

document.addEventListener('DOMContentLoaded', function () {
    var modal = document.getElementById('pmConfirmModal');
    if (!modal) return;

    var close = function () {
        modal.hidden = true;
        modal.setAttribute('aria-hidden', 'true');
        pmPendingForm = null;
    };

    var cancel = modal.querySelector('[data-pm-modal-cancel]');
    var ok = modal.querySelector('[data-pm-modal-ok]');

    if (cancel) cancel.addEventListener('click', close);

    if (ok) ok.addEventListener('click', function () {
        if (pmPendingForm) {
            var form = pmPendingForm;
            pmPendingForm = null;
            modal.hidden = true;
            form.submit();
        } else {
            close();
        }
    });

    modal.addEventListener('click', function (ev) {
        if (ev.target === modal) close();
    });

    document.addEventListener('keydown', function (ev) {
        if (ev.key === 'Escape' && !modal.hidden) close();
    });
});

/* ==========================================================================
   Utilidades de interfaz
   ========================================================================== */
document.addEventListener('DOMContentLoaded', function () {
    /* Alternar la barra lateral en pantallas pequenas */
    var toggles = document.querySelectorAll('[data-pm-sidebar-toggle]');
    var sidebar = document.getElementById('pmSidebar');
    var backdrop = document.querySelector('.pm-sidebar-backdrop');

    for (var i = 0; i < toggles.length; i++) {
        toggles[i].addEventListener('click', function () {
            if (!sidebar) return;
            sidebar.classList.toggle('is-open');
            if (backdrop) backdrop.classList.toggle('is-open');
        });
    }

    /* Nombre del archivo elegido en los inputs de tipo file */
    var fileInputs = document.querySelectorAll('.pm-file input[type="file"]');
    for (var f = 0; f < fileInputs.length; f++) {
        fileInputs[f].addEventListener('change', function () {
            var nameEl = this.parentNode.querySelector('.pm-file-name');
            if (nameEl) {
                nameEl.textContent = this.files && this.files.length
                    ? this.files[0].name
                    : 'Ningun archivo seleccionado';
            }
        });
    }

    /* Ocultar mensajes flash automaticamente */
    var flashes = document.querySelectorAll('[data-pm-autohide]');
    for (var g = 0; g < flashes.length; g++) {
        (function (node) {
            setTimeout(function () {
                node.style.transition = 'opacity .4s ease, transform .4s ease';
                node.style.opacity = '0';
                node.style.transform = 'translateY(-8px)';
                setTimeout(function () { if (node.parentNode) node.parentNode.removeChild(node); }, 420);
            }, 6000);
        })(flashes[g]);
    }
});

/* Rellena el formulario de login con las credenciales de demostracion */
function pmFillLogin(user, pass) {
    var u = document.getElementById('username');
    var p = document.getElementById('password');
    if (u) u.value = user;
    if (p) p.value = pass;
    if (u) u.focus();
}
