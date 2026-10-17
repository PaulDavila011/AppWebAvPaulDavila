/* =====================================================================
   PoliMusic — Cliente
   - Reproductor persistente (cola, aleatorio, repetir, volumen, Media Session)
   - Navegación parcial (PJAX) para que la música no se corte al cambiar de página
   - Utilidades de formularios (confirmaciones, vista previa, filtros)
   ===================================================================== */
(function () {
  'use strict';

  var $ = function (sel, ctx) { return (ctx || document).querySelector(sel); };
  var $$ = function (sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); };
  var STORE_KEY = 'polimusic.player.v1';
  var csrf = ($('meta[name="csrf-token"]') || {}).content || '';

  function storageGet() { try { return JSON.parse(sessionStorage.getItem(STORE_KEY) || 'null'); } catch (e) { return null; } }
  function storageSet(v) { try { sessionStorage.setItem(STORE_KEY, JSON.stringify(v)); } catch (e) { /* sin almacenamiento */ } }
  function fmtTime(s) {
    if (!isFinite(s) || s < 0) s = 0;
    var m = Math.floor(s / 60), r = Math.floor(s % 60);
    return m + ':' + (r < 10 ? '0' : '') + r;
  }
  function iconSvg(name) { return '<svg class="ico" aria-hidden="true"><use href="#i-' + name + '"/></svg>'; }

  function toast(msg) {
    var box = $('#toasts');
    if (!box) return;
    var t = document.createElement('div');
    t.className = 'toast';
    t.innerHTML = iconSvg('alert') + '<span></span>';
    t.querySelector('span').textContent = msg;
    box.appendChild(t);
    setTimeout(function () { t.remove(); }, 4500);
  }

  /* ------------------------------ Reproductor ------------------------------ */
  var Player = (function () {
    var el = $('#player');
    if (!el) return null;

    var audio = new Audio();
    audio.preload = 'metadata';
    var ui = {
      cover: $('#pl-cover'), title: $('#pl-title'), artist: $('#pl-artist'),
      toggle: $('#pl-toggle'), prev: $('#pl-prev'), next: $('#pl-next'),
      shuffle: $('#pl-shuffle'), repeat: $('#pl-repeat'),
      seek: $('#pl-seek'), time: $('#pl-time'), dur: $('#pl-dur'),
      mute: $('#pl-mute'), volume: $('#pl-volume')
    };
    var state = { queue: [], index: -1, shuffle: false, repeat: 'off', volume: 0.8, muted: false };
    var order = [];
    var seeking = false;
    var errorsInRow = 0;
    var countedFor = null;

    function current() { return state.queue[state.index] || null; }

    function buildOrder() {
      order = state.queue.map(function (_, i) { return i; });
      if (state.shuffle) {
        for (var i = order.length - 1; i > 0; i--) {
          var j = Math.floor(Math.random() * (i + 1)), t = order[i]; order[i] = order[j]; order[j] = t;
        }
        var pos = order.indexOf(state.index);
        if (pos > 0) { order.splice(pos, 1); order.unshift(state.index); }
      }
    }

    function save() {
      storageSet({
        queue: state.queue, index: state.index, shuffle: state.shuffle, repeat: state.repeat,
        volume: state.volume, muted: state.muted, time: audio.currentTime || 0
      });
    }

    function setRange(r, pct) { r.style.setProperty('--p', pct + '%'); }

    function renderMeta() {
      var t = current();
      el.classList.toggle('is-empty', !t);
      if (!t) return;
      ui.title.textContent = t.title;
      ui.artist.textContent = t.artist;
      ui.cover.src = t.cover || ui.cover.dataset.fallback;
      document.title = document.title.replace(/^▶ /, '');
      if ('mediaSession' in navigator) {
        try {
          navigator.mediaSession.metadata = new MediaMetadata({
            title: t.title, artist: t.artist, album: 'PoliMusic',
            artwork: t.cover ? [{ src: new URL(t.cover, location.href).href, sizes: '512x512' }] : []
          });
        } catch (e) { /* navegador sin soporte */ }
      }
    }

    function renderState() {
      var playing = !audio.paused && !audio.ended;
      ui.toggle.innerHTML = iconSvg(playing ? 'pause' : 'play');
      ui.toggle.setAttribute('aria-label', playing ? 'Pausar' : 'Reproducir');
      ui.shuffle.setAttribute('aria-pressed', state.shuffle ? 'true' : 'false');
      ui.repeat.setAttribute('aria-pressed', state.repeat !== 'off' ? 'true' : 'false');
      ui.repeat.classList.toggle('repeat-1', state.repeat === 'one');
      ui.repeat.title = state.repeat === 'off' ? 'Repetir' : (state.repeat === 'all' ? 'Repetir todo' : 'Repetir una');
      ui.mute.innerHTML = iconSvg(state.muted || state.volume === 0 ? 'volume-x' : 'volume');
      ui.volume.value = Math.round((state.muted ? 0 : state.volume) * 100);
      setRange(ui.volume, ui.volume.value);
      highlightRows();
    }

    function highlightRows() {
      var t = current();
      var playing = !audio.paused && !audio.ended;
      $$('[data-track]').forEach(function (row) {
        var isCur = !!t && row.dataset.id === String(t.id);
        row.classList.toggle('is-current', isCur);
        row.classList.toggle('is-playing', isCur && playing);
        var b = row.querySelector('[data-play-track]');
        if (b) b.innerHTML = iconSvg(isCur && playing ? 'pause' : 'play');
      });
      $$('[data-play-list]').forEach(function (b) {
        if (b.dataset.label === undefined) b.dataset.label = b.textContent.trim();
        var list = $(b.getAttribute('data-play-list'));
        var mine = !!t && !!list && !!list.querySelector('[data-track][data-id="' + t.id + '"]') && sameQueue(list);
        b.classList.toggle('is-playing', mine && playing);
        b.innerHTML = iconSvg(mine && playing ? 'pause' : 'play') + (b.classList.contains('btn-play-big') ? '' : (b.dataset.label || ''));
      });
    }

    function sameQueue(list) {
      var ids = $$('[data-track]', list).map(function (r) { return r.dataset.id; });
      if (ids.length !== state.queue.length) return false;
      for (var i = 0; i < ids.length; i++) if (ids[i] !== String(state.queue[i].id)) return false;
      return true;
    }

    function load(index, autoplay, startAt) {
      if (index < 0 || index >= state.queue.length) return;
      state.index = index;
      var t = current();
      countedFor = null;
      audio.src = t.src;
      if (startAt) {
        audio.addEventListener('loadedmetadata', function once() {
          audio.removeEventListener('loadedmetadata', once);
          try { audio.currentTime = startAt; } catch (e) { /* ignora */ }
        });
      }
      renderMeta();
      if (autoplay) {
        var p = audio.play();
        if (p && p.catch) p.catch(function () { renderState(); });
      }
      renderState();
      save();
    }

    function playQueue(tracks, startIndex) {
      if (!tracks.length) return;
      state.queue = tracks;
      state.index = startIndex || 0;
      buildOrder();
      errorsInRow = 0;
      load(state.index, true);
    }

    function toggle() {
      if (!current()) {
        var firstList = $('[data-tracklist]');
        if (firstList) playQueue(tracksFrom(firstList), 0);
        return;
      }
      if (audio.paused) { var p = audio.play(); if (p && p.catch) p.catch(function () {}); } else { audio.pause(); }
    }

    function step(dir, fromEnd) {
      if (!state.queue.length) return;
      if (dir < 0 && audio.currentTime > 3) { audio.currentTime = 0; return; }
      var pos = order.indexOf(state.index) + dir;
      if (pos >= order.length) {
        if (state.repeat === 'all' || !fromEnd) pos = 0; else { audio.pause(); audio.currentTime = 0; renderState(); return; }
      }
      if (pos < 0) pos = state.repeat === 'all' ? order.length - 1 : 0;
      load(order[pos], true);
    }

    function countPlay() {
      var t = current();
      if (!t || countedFor === t.id) return;
      countedFor = t.id;
      var body = 'id=' + encodeURIComponent(t.id) + '&csrf=' + encodeURIComponent(csrf);
      fetch('play.asp', { method: 'POST', credentials: 'same-origin', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: body })
        .then(function (r) { return r.ok ? r.json() : null; })
        .then(function (j) {
          if (!j || !j.ok) return;
          $$('[data-track][data-id="' + j.id + '"] .tl-plays').forEach(function (c) { c.textContent = j.plays.toLocaleString('es'); });
        })
        .catch(function () {});
    }

    /* Eventos del elemento de audio */
    audio.addEventListener('play', renderState);
    audio.addEventListener('pause', function () { renderState(); save(); });
    audio.addEventListener('playing', function () { errorsInRow = 0; countPlay(); });
    audio.addEventListener('ended', function () {
      if (state.repeat === 'one') { audio.currentTime = 0; audio.play(); return; }
      step(1, true);
    });
    audio.addEventListener('loadedmetadata', function () { ui.dur.textContent = fmtTime(audio.duration); });
    audio.addEventListener('timeupdate', function () {
      if (seeking) return;
      var pct = audio.duration ? (audio.currentTime / audio.duration) * 1000 : 0;
      ui.seek.value = pct;
      setRange(ui.seek, pct / 10);
      ui.time.textContent = fmtTime(audio.currentTime);
    });
    audio.addEventListener('error', function () {
      var t = current();
      if (!t || !audio.getAttribute('src')) return;
      errorsInRow++;
      toast('No se pudo reproducir «' + t.title + '»: el archivo de audio no está disponible.');
      renderState();
      if (errorsInRow < state.queue.length && state.queue.length > 1) setTimeout(function () { step(1, true); }, 900);
    });

    /* Controles */
    ui.toggle.addEventListener('click', toggle);
    ui.next.addEventListener('click', function () { step(1); });
    ui.prev.addEventListener('click', function () { step(-1); });
    ui.shuffle.addEventListener('click', function () { state.shuffle = !state.shuffle; buildOrder(); renderState(); save(); });
    ui.repeat.addEventListener('click', function () {
      state.repeat = state.repeat === 'off' ? 'all' : (state.repeat === 'all' ? 'one' : 'off');
      renderState(); save();
    });
    ui.seek.addEventListener('input', function () {
      seeking = true;
      setRange(ui.seek, ui.seek.value / 10);
      if (audio.duration) ui.time.textContent = fmtTime(audio.duration * ui.seek.value / 1000);
    });
    ui.seek.addEventListener('change', function () {
      if (audio.duration) audio.currentTime = audio.duration * ui.seek.value / 1000;
      seeking = false;
    });
    ui.volume.addEventListener('input', function () {
      state.volume = ui.volume.value / 100;
      state.muted = state.volume === 0;
      audio.volume = state.volume;
      audio.muted = state.muted;
      renderState(); save();
    });
    ui.mute.addEventListener('click', function () {
      state.muted = !state.muted;
      if (!state.muted && state.volume === 0) state.volume = 0.5;
      audio.muted = state.muted;
      audio.volume = state.volume;
      renderState(); save();
    });

    if ('mediaSession' in navigator) {
      try {
        navigator.mediaSession.setActionHandler('play', toggle);
        navigator.mediaSession.setActionHandler('pause', toggle);
        navigator.mediaSession.setActionHandler('previoustrack', function () { step(-1); });
        navigator.mediaSession.setActionHandler('nexttrack', function () { step(1); });
      } catch (e) { /* sin soporte */ }
    }

    window.addEventListener('beforeunload', save);
    setInterval(function () { if (!audio.paused) save(); }, 5000);

    /* Restaurar estado tras una recarga completa (queda en pausa) */
    var saved = storageGet();
    if (saved && saved.queue && saved.queue.length) {
      state.queue = saved.queue;
      state.shuffle = !!saved.shuffle;
      state.repeat = saved.repeat || 'off';
      state.volume = typeof saved.volume === 'number' ? saved.volume : 0.8;
      state.muted = !!saved.muted;
      buildOrder();
      audio.volume = state.volume;
      audio.muted = state.muted;
      load(Math.min(saved.index || 0, state.queue.length - 1), false, saved.time || 0);
    } else {
      audio.volume = state.volume;
      el.classList.add('is-empty');
      renderState();
    }

    return {
      playQueue: playQueue, toggle: toggle, highlightRows: highlightRows,
      isCurrentList: function (list) { return !!current() && sameQueue(list); },
      currentId: function () { var t = current(); return t ? String(t.id) : null; },
      isPlaying: function () { return !audio.paused && !audio.ended; }
    };
  })();

  function tracksFrom(list) {
    return $$('[data-track]', list).map(function (r) {
      return { id: r.dataset.id, src: r.dataset.src, title: r.dataset.title, artist: r.dataset.artist, cover: r.dataset.cover };
    });
  }

  /* Delegación de clics de reproducción */
  document.addEventListener('click', function (e) {
    if (!Player) return;
    var btn = e.target.closest('[data-play-track]');
    if (btn) {
      var row = btn.closest('[data-track]');
      var list = btn.closest('[data-tracklist]');
      if (Player.currentId() === row.dataset.id && Player.isCurrentList(list)) { Player.toggle(); return; }
      var rows = $$('[data-track]', list);
      Player.playQueue(tracksFrom(list), rows.indexOf(row));
      return;
    }
    var all = e.target.closest('[data-play-list]');
    if (all) {
      var target = $(all.getAttribute('data-play-list'));
      var tl = target && (target.matches('[data-tracklist]') ? target : target.querySelector('[data-tracklist]'));
      if (!tl) return;
      if (Player.isCurrentList(tl)) { Player.toggle(); return; }
      Player.playQueue(tracksFrom(tl), 0);
    }
  });
  document.addEventListener('dblclick', function (e) {
    var row = e.target.closest('[data-track]');
    if (row && Player && !e.target.closest('a,button')) {
      var list = row.closest('[data-tracklist]');
      Player.playQueue(tracksFrom(list), $$('[data-track]', list).indexOf(row));
    }
  });
  document.addEventListener('keydown', function (e) {
    if (!Player || e.code !== 'Space') return;
    var tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || tag === 'select' || tag === 'button' || e.target.isContentEditable) return;
    e.preventDefault();
    Player.toggle();
  });
  /* ------------------------- Navegación parcial (PJAX) ------------------------- */
  var content = $('#content');
  var bar = document.createElement('div');
  bar.className = 'loading-bar';
  document.body.appendChild(bar);

  function canPjax(url) {
    return content && Player && url.origin === location.origin && /\.asp$/i.test(url.pathname) &&
      !/(logout|login|register)\.asp$/i.test(url.pathname);
  }

  function navigate(href, push) {
    var url = new URL(href, location.href);
    bar.style.opacity = '1'; bar.style.width = '35%';
    fetch(url.href, { credentials: 'same-origin', headers: { 'X-Requested-With': 'pjax' } })
      .then(function (r) {
        if (/\/(login|register)\.asp/i.test(r.url)) throw new Error('auth');
        return r.text().then(function (html) { return { html: html, url: r.url }; });
      })
      .then(function (res) {
        var doc = new DOMParser().parseFromString(res.html, 'text/html');
        var next = doc.getElementById('content');
        if (!next || !doc.getElementById('player')) throw new Error('layout');
        content.innerHTML = next.innerHTML;
        content.setAttribute('data-nav', next.getAttribute('data-nav') || '');
        document.title = doc.title;
        var token = doc.querySelector('meta[name="csrf-token"]');
        if (token) csrf = token.content;
        var nav = content.getAttribute('data-nav');
        $$('.sidebar [data-nav]').forEach(function (a) { a.classList.toggle('active', a.classList.contains('nav-link') && a.dataset.nav === nav); });
        var search = $('.top-search input');
        if (search && nav !== 'search') search.value = '';
        if (push) history.pushState({ pjax: true }, '', res.url || url.href);
        content.scrollTop = 0;
        document.body.classList.remove('menu-open');
        closeMenus();
        initPage();
        content.focus({ preventScroll: true });
        bar.style.width = '100%';
        setTimeout(function () { bar.style.opacity = '0'; bar.style.width = '0'; }, 250);
      })
      .catch(function () { location.href = url.href; });
  }

  document.addEventListener('click', function (e) {
    var a = e.target.closest('a[href]');
    if (!a || e.defaultPrevented || e.button !== 0 || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
    if (a.target || a.hasAttribute('download') || a.hasAttribute('data-no-pjax')) return;
    var url = new URL(a.href, location.href);
    if (url.hash && url.pathname === location.pathname && url.search === location.search) return;
    if (!canPjax(url)) return;
    e.preventDefault();
    navigate(url.href, true);
  });

  document.addEventListener('submit', function (e) {
    var f = e.target;
    if (f.hasAttribute('data-confirm') && !confirm(f.getAttribute('data-confirm'))) { e.preventDefault(); return; }
    if (f.hasAttribute('data-pjax-form') && (f.method || 'get').toLowerCase() === 'get') {
      var url = new URL(f.getAttribute('action') || location.pathname, location.href);
      url.search = new URLSearchParams(new FormData(f)).toString();
      if (canPjax(url)) { e.preventDefault(); navigate(url.href, true); }
      return;
    }
    if ((f.method || '').toLowerCase() === 'post') {
      var b = f.querySelector('[type=submit]');
      if (b) setTimeout(function () { b.disabled = true; b.classList.add('is-busy'); }, 0);
    }
  });

  window.addEventListener('popstate', function () { if (content && Player) navigate(location.href, false); });

  /* ----------------------------- Interfaz general ----------------------------- */
  function closeMenus() { $$('details.user-menu[open]').forEach(function (d) { d.removeAttribute('open'); }); }
  document.addEventListener('click', function (e) {
    if (e.target.closest('[data-open-menu]')) { document.body.classList.add('menu-open'); return; }
    if (e.target.closest('[data-close-menu]')) { document.body.classList.remove('menu-open'); return; }
    if (!e.target.closest('details.user-menu')) closeMenus();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { closeMenus(); document.body.classList.remove('menu-open'); }
  });

  function initPage() {
    // Vista previa de imágenes antes de subir
    $$('input[type=file][data-preview]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        var img = $(inp.getAttribute('data-preview'));
        var f = inp.files && inp.files[0];
        inp.closest('.file-drop').classList.toggle('has-file', !!f);
        if (img && f && /^image\//.test(f.type)) img.src = URL.createObjectURL(f);
      });
    });
    // Nombre del archivo de audio seleccionado
    $$('input[type=file][data-file-name]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        var f = inp.files && inp.files[0];
        var drop = inp.closest('.file-drop');
        drop.classList.toggle('has-file', !!f);
        var label = drop.querySelector('span');
        if (f) label.textContent = f.name + ' (' + (f.size / 1048576).toFixed(1) + ' MB)';
        var path = $('#song_path');
        if (f && path) path.value = '';
      });
    });
    // Filtro y contador en listas de casillas
    $$('[data-checklist]').forEach(function (box) {
      var input = $('[data-checklist-filter]', box);
      var count = $('[data-checklist-count]', box);
      var items = $$('.check-item', box);
      function refresh() {
        var n = items.filter(function (i) { return i.querySelector('input').checked; }).length;
        count.textContent = n + ' seleccionada' + (n === 1 ? '' : 's');
      }
      input.addEventListener('input', function () {
        var q = input.value.trim().toLowerCase();
        items.forEach(function (i) { i.hidden = q && i.textContent.toLowerCase().indexOf(q) === -1; });
      });
      box.addEventListener('change', refresh);
      refresh();
    });
    if (Player) Player.highlightRows();
  }

  initPage();
})();
