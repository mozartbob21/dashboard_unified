(function () {
  'use strict';
  const button = document.querySelector('[data-autumn-pause]');
  if (!button) return;
  const motion = window.matchMedia('(prefers-reduced-motion: reduce)');
  const update = () => { button.hidden = motion.matches; };
  update();
  motion.addEventListener('change', update);
  button.addEventListener('click', () => {
    const paused = document.body.classList.toggle('autumn-paused');
    button.setAttribute('aria-pressed', String(paused));
    button.textContent = paused ? 'Продолжить анимацию' : 'Пауза фона';
  });
})();
