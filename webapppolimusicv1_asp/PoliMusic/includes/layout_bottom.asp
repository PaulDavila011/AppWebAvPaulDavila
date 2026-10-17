    </main>
  </div>

  <footer class="player" id="player" aria-label="Reproductor">
    <div class="player-track">
      <img class="player-cover" id="pl-cover" src="assets/img/fallback-cover.svg" data-fallback="assets/img/fallback-cover.svg" alt="">
      <div class="player-meta">
        <strong id="pl-title">Nada en reproducción</strong>
        <small id="pl-artist">Elige una canción para empezar</small>
      </div>
    </div>
    <div class="player-center">
      <div class="player-controls">
        <button class="ctrl" type="button" id="pl-shuffle" aria-label="Aleatorio" aria-pressed="false" title="Aleatorio"><%= Icon("shuffle") %></button>
        <button class="ctrl" type="button" id="pl-prev" aria-label="Anterior" title="Anterior"><%= Icon("prev") %></button>
        <button class="ctrl ctrl-main" type="button" id="pl-toggle" aria-label="Reproducir" title="Reproducir / Pausa (Espacio)"><%= Icon("play") %></button>
        <button class="ctrl" type="button" id="pl-next" aria-label="Siguiente" title="Siguiente"><%= Icon("next") %></button>
        <button class="ctrl" type="button" id="pl-repeat" aria-label="Repetir" aria-pressed="false" title="Repetir"><%= Icon("repeat") %><span class="repeat-one">1</span></button>
      </div>
      <div class="player-progress">
        <span id="pl-time" class="time">0:00</span>
        <input type="range" id="pl-seek" min="0" max="1000" value="0" step="1" aria-label="Posición de la canción" class="range">
        <span id="pl-dur" class="time">0:00</span>
      </div>
    </div>
    <div class="player-extra">
      <button class="ctrl" type="button" id="pl-mute" aria-label="Silenciar" title="Silenciar"><%= Icon("volume") %></button>
      <input type="range" id="pl-volume" min="0" max="100" value="80" aria-label="Volumen" class="range range-vol">
    </div>
  </footer>
</div>
<div class="toasts" id="toasts" aria-live="polite"></div>
<script src="assets/js/app.js?v=<%= APP_VERSION %>"></script>
</body>
</html>
<% DbClose %>
