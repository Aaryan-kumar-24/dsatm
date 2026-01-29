function toggleTheme() {
    const body = document.body;
    const icon = document.getElementById('theme-icon');
    
    body.classList.toggle('dark-theme');
    
    if (body.classList.contains('dark-theme')) {
        icon.className = 'fas fa-sun';
        localStorage.setItem('theme', 'dark');
    } else {
        icon.className = 'fas fa-moon';
        localStorage.setItem('theme', 'light');
    }
}

// Load saved theme
document.addEventListener('DOMContentLoaded', function() {
    if (localStorage.getItem('theme') === 'dark') {
        document.body.classList.add('dark-theme');
        const icon = document.getElementById('theme-icon');
        if (icon) icon.className = 'fas fa-sun';
    }
});