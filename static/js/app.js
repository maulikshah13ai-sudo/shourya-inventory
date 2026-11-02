// Common frontend behaviors

document.addEventListener('DOMContentLoaded', () => {
    // Auto dismiss flash messages after 4 seconds
    setTimeout(() => {
        const flashes = document.querySelectorAll('.max-w-7xl .p-4.rounded-xl');
        flashes.forEach(flash => {
            flash.style.transition = 'opacity 0.5s ease-out, transform 0.5s ease-out';
            flash.style.opacity = '0';
            flash.style.transform = 'translateY(-10px)';
            setTimeout(() => flash.remove(), 500);
        });
    }, 4500);
});
