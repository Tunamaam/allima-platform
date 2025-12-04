document.addEventListener('DOMContentLoaded', () => {
    // Sidebar Toggle
    const menuToggle = document.querySelector('.menu-toggle');
    const sidebar = document.querySelector('.sidebar');
    const appContainer = document.querySelector('.app-container');

    if (menuToggle) {
        menuToggle.addEventListener('click', () => {
            sidebar.classList.toggle('active');
        });
    }

    // Close sidebar when clicking outside on mobile
    document.addEventListener('click', (e) => {
        if (window.innerWidth <= 1024) {
            if (!sidebar.contains(e.target) && !menuToggle.contains(e.target) && sidebar.classList.contains('active')) {
                sidebar.classList.remove('active');
            }
        }
    });

    // View Switching Logic
    const navItems = document.querySelectorAll('.nav-item');
    const views = {
        'Inicio': document.getElementById('view-home'),
        'Herramientas': document.getElementById('view-tools')
    };

    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            const text = item.querySelector('span').textContent;

            if (views[text]) {
                e.preventDefault(); // Prevent default anchor behavior

                // Update active state in sidebar
                navItems.forEach(nav => nav.classList.remove('active'));
                item.classList.add('active');

                // Hide all views
                Object.values(views).forEach(view => view.classList.add('hidden'));

                // Show target view
                views[text].classList.remove('hidden');
            }
        });
    });

    // Tab Filtering (Visual only for now)
    const tabs = document.querySelectorAll('.tab');

    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            // Remove active class from all tabs
            tabs.forEach(t => t.classList.remove('active'));
            // Add active class to clicked tab
            tab.classList.add('active');

            // Here you would add logic to filter the tool cards
            // For now, we just animate the cards to simulate a refresh
            const cards = document.querySelectorAll('.tool-card');
            cards.forEach(card => {
                card.style.opacity = '0.5';
                card.style.transform = 'scale(0.98)';
                setTimeout(() => {
                    card.style.opacity = '1';
                    card.style.transform = 'scale(1)';
                }, 200);
            });
        });
    });

    // Add simple hover effects for "Create Session" card
    const createSessionCard = document.querySelector('.action-card.primary');
    if (createSessionCard) {
        createSessionCard.addEventListener('click', () => {
            // Simulate navigation or action
            console.log('Navigating to Create Session...');
        });
    }
});
