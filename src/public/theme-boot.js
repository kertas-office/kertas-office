// Apply the saved theme before first paint to avoid a flash
    try {
      const t = localStorage.getItem('kertas-theme');
      const dark = t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
      document.documentElement.dataset.theme = dark ? 'dark' : 'light';
    } catch { document.documentElement.dataset.theme = 'light'; }
