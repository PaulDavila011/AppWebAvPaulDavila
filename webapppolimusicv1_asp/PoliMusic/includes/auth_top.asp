<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="dark">
<meta name="theme-color" content="#0d0817">
<title><%= H(pageTitle) %> · <%= APP_NAME %></title>
<link rel="icon" href="assets/img/logo.svg" type="image/svg+xml">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="assets/css/app.css?v=<%= APP_VERSION %>">
</head>
<body class="auth-body">
<!--#include file="icons.asp"-->
<div class="auth">
  <section class="auth-hero" aria-hidden="true">
    <div class="auth-brand"><img src="assets/img/logo.svg" alt="" width="44" height="44"><span>Poli<b>Music</b></span></div>
    <div class="auth-hero-body">
      <div class="bars"><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i></div>
      <h2>Tu música,<br>sin límites.</h2>
      <p>Descubre álbumes, sigue a tus artistas favoritos y escucha tus playlists en cualquier momento.</p>
    </div>
    <p class="auth-foot">Escuela Politécnica Nacional · <%= APP_NAME %> v<%= APP_VERSION %></p>
  </section>
  <main class="auth-panel" id="content">
    <div class="auth-card">
      <div class="auth-brand auth-brand-sm"><img src="assets/img/logo.svg" alt="" width="36" height="36"><span>Poli<b>Music</b></span></div>
