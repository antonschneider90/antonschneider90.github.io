// Password protection check - runs on every page
(function() {
    // Skip check if we're on the password page itself
    if (window.location.pathname.endsWith('index.html') || window.location.pathname === '/') {
        return;
    }
    
    // Check if user has authenticated
    if (sessionStorage.getItem('siteAccess') !== 'granted') {
        // Not authenticated - redirect to password page
        window.location.href = '/index.html';
    }
})();
