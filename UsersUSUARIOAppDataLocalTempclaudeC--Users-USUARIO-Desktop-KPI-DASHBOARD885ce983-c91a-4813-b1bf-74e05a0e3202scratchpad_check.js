
    // Aplicar tema guardado ANTES del primer paint, para evitar flash de color equivocado.
    (function(){
      try {
        var saved = localStorage.getItem('remitera-theme');
        var theme = saved || (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
        document.documentElement.setAttribute('data-theme', theme);
      } catch (e) {}
    })();
  