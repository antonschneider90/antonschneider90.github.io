// Blog posts data - Update this array to add new blog posts
const blogPosts = [
    {
        title: "The Future of Energy Storage",
        date: "2026-03-15",
        category: "Energy Storage",
        excerpt: "Exploring how battery technology is reshaping the energy landscape and enabling higher renewable penetration.",
        slug: "future-energy-storage"
    },
    {
        title: "Modeling Grid Reliability in High-Renewable Scenarios",
        date: "2026-02-28",
        category: "Grid Analysis",
        excerpt: "A deep dive into the challenges and solutions for maintaining grid stability with 80%+ renewable energy.",
        slug: "grid-reliability-renewables"
    },
    {
        title: "Data-Driven Approaches to Energy Policy",
        date: "2026-02-10",
        category: "Policy",
        excerpt: "How advanced analytics and modeling can inform better energy policy decisions and accelerate transitions.",
        slug: "data-driven-energy-policy"
    }
];

// Load blog posts on homepage
function loadRecentBlogPosts() {
    const blogContainer = document.getElementById('blog-posts');
    if (!blogContainer) return;

    // Show only the 3 most recent posts on homepage
    const recentPosts = blogPosts.slice(0, 3);

    blogContainer.innerHTML = recentPosts.map(post => `
        <article class="blog-card">
            <div class="blog-card-content">
                <div class="blog-meta">
                    <span>${new Date(post.date).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}</span>
                    <span>•</span>
                    <span>${post.category}</span>
                </div>
                <h3>${post.title}</h3>
                <p>${post.excerpt}</p>
                <a href="blog/${post.slug}.html">Read More →</a>
            </div>
        </article>
    `).join('');
}

// Smooth scroll for navigation links
document.addEventListener('DOMContentLoaded', () => {
    // Load blog posts
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
    let lastScroll = 0;

    window.addEventListener('scroll', () => {
        const currentScroll = window.pageYOffset;

        if (currentScroll > 100) {
            navbar.style.boxShadow = '0 4px 6px rgba(0, 0, 0, 0.1)';
        } else {
            navbar.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.1)';
        }

        lastScroll = currentScroll;
    });

    // Add animation on scroll for sections
    const observerOptions = {
        threshold: 0.1,
        rootMargin: '0px 0px -100px 0px'
    };

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.opacity = '1';
                entry.target.style.transform = 'translateY(0)';
            }
        });
    }, observerOptions);

    // Observe all cards
    document.querySelectorAll('.model-card, .blog-card').forEach(card => {
        card.style.opacity = '0';
        card.style.transform = 'translateY(20px)';
        card.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        observer.observe(card);
    });
});
