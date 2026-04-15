// Blog posts data
const blogPosts = [
   {
    title: "Driving Impact in the Energy Markets",
    slug: "driving-impact-in-the-energy-markets",
    date: "2026-04-15",
    category: "Energy Markets",
    excerpt: "Energy markets are in turmoil. Again. The 2026 Iran war, the unresolved fallout from Russia's invasion of Ukraine, and political backlash against climate policy — why I got involved and what V2G can do."
},
{
    title: "Financial Templates for Innovative Ideas",
    slug: "financial-templates-for-innovative-ideas",
    date: "2026-04-15",
    category: "Finance & Entrepreneurship",
    excerpt: "I spent years analyzing pre-IPO business plans at UBS. Then I became employee #1 at a startup and had to build my own from scratch. Here's why I'm sharing the playbook."
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
