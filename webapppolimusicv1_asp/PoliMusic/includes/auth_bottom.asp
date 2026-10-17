    </div>
  </main>
</div>
<script>
document.querySelectorAll('[data-toggle-password]').forEach(function (b) {
  b.addEventListener('click', function () {
    var i = document.getElementById(b.getAttribute('data-toggle-password'));
    var show = i.type === 'password';
    i.type = show ? 'text' : 'password';
    b.textContent = show ? 'Ocultar' : 'Mostrar';
    b.setAttribute('aria-pressed', show ? 'true' : 'false');
  });
});
document.querySelectorAll('form[data-busy]').forEach(function (f) {
  f.addEventListener('submit', function () { var b = f.querySelector('[type=submit]'); if (b) { b.disabled = true; b.classList.add('is-busy'); } });
});
</script>
</body>
</html>
<% DbClose %>
