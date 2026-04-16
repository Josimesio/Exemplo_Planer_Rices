function normalizeUsuario(usuario) {
  return String(usuario || '').trim().toLowerCase();
}

function fazerLogin(usuario, senha) {
  const credenciais = {
    conecta: 'pluma@2026'
  };

  const usuarioNormalizado = normalizeUsuario(usuario);
  const senhaInformada = String(senha || '');

  if (!Object.prototype.hasOwnProperty.call(credenciais, usuarioNormalizado)) {
    return { sucesso: false, erro: 'usuario' };
  }

  if (credenciais[usuarioNormalizado] !== senhaInformada) {
    return { sucesso: false, erro: 'senha' };
  }

  sessionStorage.setItem('painel_autenticado', 'true');
  sessionStorage.setItem('painel_usuario', usuarioNormalizado);
  return { sucesso: true };
}

function protegerPagina(loginPage) {
  if (sessionStorage.getItem('painel_autenticado') !== 'true') {
    window.location.href = loginPage;
  }
}

function fazerLogout() {
  sessionStorage.removeItem('painel_autenticado');
  sessionStorage.removeItem('painel_usuario');
  window.location.href = 'login.html';
}

function limparSessao() {
  sessionStorage.removeItem('painel_autenticado');
  sessionStorage.removeItem('painel_usuario');
}
