const loginMotionCanvas = document.getElementById("loginMotionCanvas");
let loginMotionCanvasCleanup = null;

export function startLoginMotionCanvas() {
  if (!loginMotionCanvas || loginMotionCanvasCleanup) return;
  const canvas = loginMotionCanvas;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;
  const reduceMotion = window.matchMedia?.("(prefers-reduced-motion: reduce)")?.matches === true;
  const pointer = { x: 0, y: 0, tx: 0, ty: 0, active: false };
  let width = 0;
  let height = 0;
  let dpr = 1;
  let raf = 0;
  let lastFrame = 0;
  let nodes = [];
  const colors = ["#3e4095", "#5661d6", "#7d7fe2", "#d1ae6c"];
  const buildNodes = () => {
    const spacing = width < 700 ? 56 : 74;
    const cols = Math.max(5, Math.ceil(width / spacing) + 2);
    const rows = Math.max(5, Math.ceil(height / spacing) + 2);
    const xGap = width / Math.max(1, cols - 1);
    const yGap = height / Math.max(1, rows - 1);
    nodes = [];
    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        const jitterX = (Math.random() - 0.5) * xGap * 0.44;
        const jitterY = (Math.random() - 0.5) * yGap * 0.44;
        const homeX = col * xGap + jitterX;
        const homeY = row * yGap + jitterY;
        nodes.push({
          homeX,
          homeY,
          x: homeX,
          y: homeY,
          vx: 0,
          vy: 0,
          phase: Math.random() * Math.PI * 2,
          drift: 3 + Math.random() * 8,
          size: 1.2 + Math.random() * 2.1,
          color: colors[(row + col) % colors.length],
          alpha: 0.26 + Math.random() * 0.42
        });
      }
    }
  };
  const resize = () => {
    const rect = canvas.getBoundingClientRect();
    dpr = Math.min(1.35, window.devicePixelRatio || 1);
    width = Math.max(1, Math.floor(rect.width));
    height = Math.max(1, Math.floor(rect.height));
    canvas.width = Math.floor(width * dpr);
    canvas.height = Math.floor(height * dpr);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    pointer.x = pointer.tx = width * 0.5;
    pointer.y = pointer.ty = height * 0.5;
    buildNodes();
  };
  const move = (event) => {
    pointer.tx = event.clientX;
    pointer.ty = event.clientY;
    pointer.active = true;
  };
  const leave = () => {
    pointer.active = false;
  };
  const draw = (ts = 0) => {
    if (document.hidden) {
      raf = requestAnimationFrame(draw);
      return;
    }
    if (ts - lastFrame < 16) {
      raf = requestAnimationFrame(draw);
      return;
    }
    lastFrame = ts;
    ctx.clearRect(0, 0, width, height);
    const time = ts * 0.001;
    pointer.x += (pointer.tx - pointer.x) * 0.17;
    pointer.y += (pointer.ty - pointer.y) * 0.17;
    ctx.save();
    ctx.globalCompositeOperation = "source-over";
    const linkDistance = width < 700 ? 84 : 108;
    const pullRadius = Math.min(360, Math.max(210, width * 0.26));
    if (pointer.active && !reduceMotion) {
      const pulseRadius = pullRadius * 0.78;
      const pulse = ctx.createRadialGradient(pointer.x, pointer.y, 0, pointer.x, pointer.y, pulseRadius);
      pulse.addColorStop(0, "rgba(209, 174, 108, 0.22)");
      pulse.addColorStop(0.55, "rgba(126, 140, 230, 0.09)");
      pulse.addColorStop(1, "rgba(62, 64, 149, 0)");
      ctx.globalAlpha = 1;
      ctx.fillStyle = pulse;
      ctx.beginPath();
      ctx.arc(pointer.x, pointer.y, pulseRadius, 0, Math.PI * 2);
      ctx.fill();
    }
    nodes.forEach((node) => {
      const idleX = Math.cos(time * 0.94 + node.phase) * node.drift;
      const idleY = Math.sin(time * 0.86 + node.phase * 1.17) * node.drift;
      const targetX = node.homeX + idleX;
      const targetY = node.homeY + idleY;
      node.vx += (targetX - node.x) * 0.02;
      node.vy += (targetY - node.y) * 0.02;

      if (pointer.active && !reduceMotion) {
        const dx = pointer.x - node.x;
        const dy = pointer.y - node.y;
        const dist = Math.max(1, Math.hypot(dx, dy));
        if (dist < pullRadius) {
          const force = (1 - dist / pullRadius) ** 2.2;
          node.vx += (dx / dist) * force * 2.25;
          node.vy += (dy / dist) * force * 2.25;
        }
      }
      node.vx *= 0.84;
      node.vy *= 0.84;
      node.x += node.vx;
      node.y += node.vy;
    });

    for (let i = 0; i < nodes.length; i += 1) {
      const a = nodes[i];
      for (let j = i + 1; j < nodes.length; j += 1) {
        const b = nodes[j];
        const dx = b.x - a.x;
        const dy = b.y - a.y;
        const dist = Math.hypot(dx, dy);
        if (dist > linkDistance) continue;
        ctx.globalAlpha = (1 - dist / linkDistance) * 0.22;
        ctx.strokeStyle = "rgba(139, 141, 201, 0.9)";
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(a.x, a.y);
        ctx.lineTo(b.x, b.y);
        ctx.stroke();
      }
    }

    nodes.forEach((node) => {
      const pointerBoost = pointer.active
        ? Math.max(0, 1 - Math.hypot(pointer.x - node.x, pointer.y - node.y) / pullRadius)
        : 0;
      ctx.globalAlpha = Math.min(0.94, node.alpha + pointerBoost * 0.5);
      ctx.fillStyle = node.color;
      ctx.beginPath();
      ctx.arc(node.x, node.y, node.size + pointerBoost * 2.2, 0, Math.PI * 2);
      ctx.fill();
    });
    ctx.restore();
    raf = requestAnimationFrame(draw);
  };
  resize();
  window.addEventListener("resize", resize);
  window.addEventListener("pointermove", move, { passive: true });
  window.addEventListener("pointerleave", leave, { passive: true });
  raf = requestAnimationFrame(draw);
  loginMotionCanvasCleanup = () => {
    cancelAnimationFrame(raf);
    window.removeEventListener("resize", resize);
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerleave", leave);
    loginMotionCanvasCleanup = null;
  };
}

export function stopLoginMotionCanvas() {
  if (typeof loginMotionCanvasCleanup === "function") loginMotionCanvasCleanup();
}
