document.addEventListener('DOMContentLoaded', function() {
    // Initialize Mermaid with zoom settings
    mermaid.initialize({
        startOnLoad: true,
        theme: 'neutral',
        securityLevel: 'loose',
        flowchart: {
            curve: 'basis',
            htmlLabels: true
        }
    });

    // Add sidebar toggle button to DOM
    const toggleButton = document.createElement('button');
    toggleButton.id = 'sidebarToggle';
    toggleButton.innerHTML = '☰';
    document.body.appendChild(toggleButton);

    // Sidebar toggle functionality
    const sidebar = document.getElementById('sidebar');
    const content = document.getElementById('content');
    const sidebarToggle = document.getElementById('sidebarToggle');

    sidebarToggle.addEventListener('click', function() {
        sidebar.classList.toggle('active');
    });

    // Close sidebar when clicking outside on mobile
    document.addEventListener('click', function(e) {
        if (window.innerWidth < 768) {
            if (!sidebar.contains(e.target) && e.target !== sidebarToggle) {
                sidebar.classList.remove('active');
            }
        }
    });

    // Add zoom controls to each mermaid diagram after they're rendered
    function addZoomControls() {
        document.querySelectorAll('.mermaid-container').forEach((container, index) => {
            const diagram = container.querySelector('.mermaid');
            const svg = diagram.querySelector('svg');
            if (!svg || container.querySelector('.zoom-controls')) return;

            // Create zoom controls
            const controls = document.createElement('div');
            controls.className = 'zoom-controls';
            controls.innerHTML = `
                <button class="zoom-btn zoom-in" title="Zoom In">+</button>
                <button class="zoom-btn zoom-out" title="Zoom Out">-</button>
                <button class="zoom-btn zoom-reset" title="Reset">Reset</button>
            `;

            // Insert controls before the diagram
            container.insertBefore(controls, diagram);

            // Initialize zoom state
            let currentZoom = 1;
            let isDragging = false;
            let startX = 0;
            let startY = 0;
            let translateX = 0;
            let translateY = 0;

            // Add zoom functionality
            container.querySelector('.zoom-in').addEventListener('click', () => {
                currentZoom *= 1.2;
                updateTransform();
            });

            container.querySelector('.zoom-out').addEventListener('click', () => {
                currentZoom *= 0.8;
                updateTransform();
            });

            container.querySelector('.zoom-reset').addEventListener('click', () => {
                currentZoom = 1;
                translateX = 0;
                translateY = 0;
                updateTransform();
            });

            // Add pan functionality
            svg.addEventListener('mousedown', (e) => {
                isDragging = true;
                startX = e.clientX - translateX;
                startY = e.clientY - translateY;
                svg.style.cursor = 'grabbing';
            });

            document.addEventListener('mousemove', (e) => {
                if (!isDragging) return;
                translateX = e.clientX - startX;
                translateY = e.clientY - startY;
                updateTransform();
            });

            document.addEventListener('mouseup', () => {
                isDragging = false;
                svg.style.cursor = 'grab';
            });

            // Add wheel zoom
            container.addEventListener('wheel', (e) => {
                if (e.ctrlKey) {
                    e.preventDefault();
                    const delta = e.deltaY > 0 ? 0.8 : 1.2;
                    currentZoom *= delta;
                    updateTransform();
                }
            });

            function updateTransform() {
                // Limit zoom levels
                currentZoom = Math.min(Math.max(0.1, currentZoom), 4);
                svg.style.transform = `translate(${translateX}px, ${translateY}px) scale(${currentZoom})`;
            }

            // Initial setup
            svg.style.cursor = 'grab';
            svg.style.transformOrigin = 'center';
            svg.style.transition = 'transform 0.1s';
        });
    }

    // Wait for Mermaid diagrams to be rendered then add controls
    setTimeout(addZoomControls, 1000);

    // Smooth scrolling for navigation links
    document.querySelectorAll('.nav-link').forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            const targetId = this.getAttribute('href').substring(1);
            const targetElement = document.getElementById(targetId);
            if (targetElement) {
                targetElement.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
                // Close sidebar on mobile after clicking a link
                if (window.innerWidth < 768) {
                    sidebar.classList.remove('active');
                }
            }
        });
    });

    // Active section highlighting
    const sections = document.querySelectorAll('section');
    const navLinks = document.querySelectorAll('.nav-link');

    function updateActiveSection() {
        const scrollPosition = window.scrollY;

        sections.forEach(section => {
            const sectionTop = section.offsetTop - 100;
            const sectionBottom = sectionTop + section.offsetHeight;

            if (scrollPosition >= sectionTop && scrollPosition < sectionBottom) {
                const currentId = section.getAttribute('id');
                navLinks.forEach(link => {
                    link.classList.remove('active');
                    if (link.getAttribute('href') === `#${currentId}`) {
                        link.classList.add('active');
                    }
                });
            }
        });
    }

    // Update active section on scroll
    window.addEventListener('scroll', updateActiveSection);
    
    // Handle window resize
    window.addEventListener('resize', function() {
        if (window.innerWidth >= 768) {
            sidebar.classList.remove('active');
        }
    });

    // Initial call to set active section
    updateActiveSection();
});