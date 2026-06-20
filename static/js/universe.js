// THREE.JS PREMIUM SPIRAL GALAXY / SUN ENGINE
document.addEventListener("DOMContentLoaded", () => {
    let canvas = document.getElementById('bg-canvas');
    if (!canvas) {
        canvas = document.createElement('canvas');
        canvas.id = 'bg-canvas';
        canvas.style.position = 'fixed';
        canvas.style.top = '0';
        canvas.style.left = '0';
        canvas.style.width = '100vw';
        canvas.style.height = '100vh';
        canvas.style.zIndex = '-1';
        canvas.style.pointerEvents = 'none';
        canvas.style.transition = 'opacity 0.6s ease';
        document.body.prepend(canvas);
    }

    const scene = new THREE.Scene();
    scene.background = new THREE.Color(0x050510);
    scene.fog = new THREE.FogExp2(0x050510, 0.0015);

    const camera = new THREE.PerspectiveCamera(60, window.innerWidth / window.innerHeight, 0.1, 2000);
    camera.position.z = 400;
    camera.position.y = 150; 
    camera.lookAt(0, 0, 0);

    const renderer = new THREE.WebGLRenderer({ canvas: canvas, antialias: true, alpha: false });
    renderer.setSize(window.innerWidth, window.innerHeight);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

    const PARTICLE_COUNT = 45000;
    const geometry = new THREE.BufferGeometry();
    
    const positions = new Float32Array(PARTICLE_COUNT * 3);
    const colorsDark = new Float32Array(PARTICLE_COUNT * 3);
    const colorsLight = new Float32Array(PARTICLE_COUNT * 3);
    const sizes = new Float32Array(PARTICLE_COUNT);
    const phases = new Float32Array(PARTICLE_COUNT);

    const colorTeal = new THREE.Color(0x34D399);
    const colorBlue = new THREE.Color(0x2563EB);
    const colorPurple = new THREE.Color(0x7C3AED);

    const colorYellow = new THREE.Color(0xFBBF24);
    const colorOrange = new THREE.Color(0xEA580C);
    const colorRed = new THREE.Color(0xEF4444);

    const ARMS = 5;
    
    for (let i = 0; i < PARTICLE_COUNT; i++) {
        const i3 = i * 3;
        
        // Dense core, sparse outer
        const radius = Math.pow(Math.random(), 2.5) * 700; 
        const armAngle = (i % ARMS) * ((Math.PI * 2) / ARMS);
        const spin = radius * 0.015;
        
        // Spread is tighter near the center
        const spread = Math.pow(Math.random(), 2) * (50 + radius * 0.1);
        const randomAngle = Math.random() * Math.PI * 2;
        const randomX = Math.cos(randomAngle) * spread;
        const randomZ = Math.sin(randomAngle) * spread;
        
        const theta = armAngle + spin;
        
        positions[i3] = Math.cos(theta) * radius + randomX;
        
        // Y thickness bulge at center
        const thickness = Math.max(2, 120 - radius * 0.4);
        positions[i3 + 1] = (Math.random() - 0.5) * thickness + (Math.random() - 0.5) * 5;
        positions[i3 + 2] = Math.sin(theta) * radius + randomZ;

        // Dark Colors
        const mixRatioD = Math.random();
        const cd = new THREE.Color().lerpColors(colorPurple, colorBlue, mixRatioD);
        if (Math.random() > 0.7) cd.lerp(colorTeal, 0.8);
        if (radius < 80) cd.lerp(new THREE.Color(0xffffff), 0.6);
        colorsDark[i3] = cd.r; colorsDark[i3 + 1] = cd.g; colorsDark[i3 + 2] = cd.b;

        // Light Colors (Sun)
        const mixRatioL = Math.random();
        const cl = new THREE.Color().lerpColors(colorOrange, colorYellow, mixRatioL);
        if (Math.random() > 0.6) cl.lerp(colorRed, 0.8);
        if (radius < 120) cl.lerp(new THREE.Color(0xffffff), 0.4); 
        colorsLight[i3] = cl.r; colorsLight[i3 + 1] = cl.g; colorsLight[i3 + 2] = cl.b;

        sizes[i] = Math.random() * 2.2 + 0.5;
        phases[i] = Math.random() * Math.PI * 2;
    }

    geometry.setAttribute('position', new THREE.BufferAttribute(positions, 3));
    geometry.setAttribute('colorDark', new THREE.BufferAttribute(colorsDark, 3));
    geometry.setAttribute('colorLight', new THREE.BufferAttribute(colorsLight, 3));
    geometry.setAttribute('size', new THREE.BufferAttribute(sizes, 1));
    geometry.setAttribute('phase', new THREE.BufferAttribute(phases, 1));

    const vertexShader = `
        attribute vec3 colorDark;
        attribute vec3 colorLight;
        attribute float size;
        attribute float phase;
        
        varying vec3 vColor;
        
        uniform float uTime;
        uniform float uThemeMix; 
        
        void main() {
            vColor = mix(colorDark, colorLight, uThemeMix);
            vec3 pos = position;
            
            float wave = sin(pos.x * 0.01 + uTime * 0.5) * 8.0 + cos(pos.z * 0.01 + uTime * 0.4) * 8.0;
            pos.y += wave;

            float dist = length(pos);
            if (dist < 200.0) {
                // Sun pulsing effect in light mode
                pos.y += mix(0.0, sin(phase + uTime * 4.0) * 15.0, uThemeMix); 
            }
            
            float pulse = sin(phase + uTime * 2.0) * 0.5 + 0.5;
            
            vec4 mvPosition = modelViewMatrix * vec4(pos, 1.0);
            gl_PointSize = size * (300.0 / -mvPosition.z) * (1.0 + pulse);
            gl_Position = projectionMatrix * mvPosition;
        }
    `;

    const fragmentShader = `
        varying vec3 vColor;
        void main() {
            float dist = distance(gl_PointCoord, vec2(0.5));
            if (dist > 0.5) discard;
            float alpha = smoothstep(0.5, 0.1, dist);
            // Lower opacity (0.2) so normal blending overlaps smoothly like soft fire/clouds
            gl_FragColor = vec4(vColor, alpha * 0.2);
        }
    `;

    const material = new THREE.ShaderMaterial({
        uniforms: {
            uTime: { value: 0 },
            uThemeMix: { value: 0 } 
        },
        vertexShader: vertexShader,
        fragmentShader: fragmentShader,
        transparent: true,
        depthWrite: false,
        blending: THREE.AdditiveBlending
    });

    const particles = new THREE.Points(geometry, material);
    scene.add(particles);

    let isLightMode = false;
    function updateThemeColors() {
        isLightMode = document.body.getAttribute('data-theme') === 'light';
        const bgColor = isLightMode ? 0xFAFAFA : 0x050510;
        
        scene.background = new THREE.Color(bgColor);
        scene.fog.color.setHex(bgColor);
        
        material.blending = isLightMode ? THREE.NormalBlending : THREE.AdditiveBlending;
        material.needsUpdate = true;
        
        if (typeof gsap !== 'undefined') {
            gsap.to(material.uniforms.uThemeMix, {
                value: isLightMode ? 1.0 : 0.0,
                duration: 1.5,
                ease: "power2.inOut"
            });
        } else {
            material.uniforms.uThemeMix.value = isLightMode ? 1.0 : 0.0;
        }
        
        canvas.style.opacity = '1'; 
    }
    
    setTimeout(updateThemeColors, 100);
    
    const observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
            if (mutation.attributeName === 'data-theme') {
                updateThemeColors();
            }
        });
    });
    observer.observe(document.body, { attributes: true });

    let mouseX = 0;
    let mouseY = 0;
    document.addEventListener('mousemove', (e) => {
        mouseX = (e.clientX - window.innerWidth / 2) * 0.001;
        mouseY = (e.clientY - window.innerHeight / 2) * 0.001;
    });

    const clock = new THREE.Clock();

    function animate() {
        requestAnimationFrame(animate);
        const time = clock.getElapsedTime();
        material.uniforms.uTime.value = time;

        const rotSpeed = isLightMode ? 0.08 : 0.04;
        particles.rotation.y += rotSpeed * 0.05;
        
        particles.rotation.z = Math.sin(time * 0.2) * 0.05;

        camera.position.x += (mouseX * 150 - camera.position.x) * 0.05;
        
        // In dark mode we look across the galaxy edge, in light mode we look slightly more top-down at the sun
        const targetY = isLightMode ? 350 + (-mouseY * 50) : 120 + (-mouseY * 100);
        camera.position.y += (targetY - camera.position.y) * 0.03;
        
        // Zoom in slightly for the sun
        const targetZ = isLightMode ? 350 : 450;
        camera.position.z += (targetZ - camera.position.z) * 0.03;
        
        camera.lookAt(scene.position);
        renderer.render(scene, camera);
    }

    animate();

    window.addEventListener('resize', () => {
        camera.aspect = window.innerWidth / window.innerHeight;
        camera.updateProjectionMatrix();
        renderer.setSize(window.innerWidth, window.innerHeight);
    });
});
