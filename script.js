// Blog posts data
const blogPosts = [
    {
        title: "Why Vehicle-to-Grid Makes Sense for Public Transport",
        date: "2026-04-05",
        category: "Energy Systems",
        excerpt: "Electric buses sit idle for 16+ hours per day. With bidirectional charging, they can become distributed energy storage assets. Here's the economic case.",
        slug: "v2g-public-transport"
    },
    {
        title: "From ISPs to Energy: The Surprising Similarities",
        date: "2026-03-28",
        category: "Business Models",
        excerpt: "Fiber networks and power grids have more in common than you'd think. Both require massive upfront capital and monetize through subscriptions. The playbook transfers.",
        slug: "isp-to-energy-transition"
    },
    {
        title: "Forecasting in Energy Markets: Lessons from Tech",
        date: "2026-03-15",
        category: "Analytics",
        excerpt: "At Google Fiber, forecasting drove profitability. In energy, it's even more critical—imbalances cost money. Here's how the technical skills translate.",
        slug: "forecasting-energy-markets"
    }
];

// Load recent blog posts on homepage
function loadRecentBlogPosts() {
    const blogContainer = document.getElementById('blog-posts');
    if (!blogContainer) return;

    const recentPosts = blogPosts.slice(0, 3);

    blogContainer.innerHTML = recentPosts.map(post => `
        <article class="blog-card">
            <div class="blog-meta">
                <span>${new Date(post.date).toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })}</span>
                <span>•</span>
                <span>${post.category}</span>
            </div>
            <h3>${post.title}</h3>
            <p>${post.excerpt}</p>
            <a href="blog/${post.slug}.html">Read More →</a>
        </article>
    `).join('');
}

// Smooth scroll for navigation links
document.addEventListener('DOMContentLoaded', () => {
    loadRecentBlogPosts();

    // Smooth scroll
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });

    // Navbar scroll effect
    const navbar = document.querySelector('.navbar');
    window.addEventListener('scroll', () => {
        if (window.pageYOffset > 100) {
            navbar.style.boxShadow = '0 4px 12px rgba(10, 17, 40, 0.08)';
        } else {
            navbar.style.boxShadow = 'none';
        }
    });

    // Intersection Observer for animations
    const observerOptions = {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    };

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.opacity = '1';
                entry.target.style.transform = 'translateY(0)';
            }
        });
    }, observerOptions);

    // Observe cards for animation
    document.querySelectorAll('.featured-model, .blog-card').forEach(card => {
        card.style.opacity = '0';
        card.style.transform = 'translateY(20px)';
        card.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        observer.observe(card);
    });
});
