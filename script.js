// Download counter functionality
class DownloadCounter {
    constructor() {
        this.storageKey = 'docsmartDownloadCount';
        this.count = this.getStoredCount();
        this.updateDisplay();
        this.setupEventListeners();
    }

    getStoredCount() {
        const stored = localStorage.getItem(this.storageKey);
        return stored ? parseInt(stored, 10) : 0;
    }

    saveCount() {
        localStorage.setItem(this.storageKey, this.count.toString());
    }

    incrementCount() {
        this.count++;
        this.saveCount();
        this.updateDisplay();
        this.showDownloadAnimation();
    }

    updateDisplay() {
        const countElements = document.querySelectorAll('#downloadCount, #downloadCountDisplay');
        countElements.forEach(element => {
            if (element) {
                element.textContent = this.count.toLocaleString();
            }
        });
    }

    showDownloadAnimation() {
        const button = document.getElementById('mainDownloadBtn');
        if (button) {
            button.style.transform = 'scale(0.95)';
            setTimeout(() => {
                button.style.transform = 'scale(1)';
            }, 150);
        }

        // Show success message
        this.showNotification('Download started! Thank you for trying Doc-smart v0.1');
    }

    showNotification(message) {
        // Create notification element
        const notification = document.createElement('div');
        notification.className = 'download-notification';
        notification.textContent = message;
        notification.style.cssText = `
            position: fixed;
            top: 100px;
            right: 20px;
            background: #27ae60;
            color: white;
            padding: 1rem 1.5rem;
            border-radius: 5px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
            z-index: 1001;
            transform: translateX(100%);
            transition: transform 0.3s ease;
        `;

        document.body.appendChild(notification);

        // Animate in
        setTimeout(() => {
            notification.style.transform = 'translateX(0)';
        }, 100);

        // Remove after 3 seconds
        setTimeout(() => {
            notification.style.transform = 'translateX(100%)';
            setTimeout(() => {
                document.body.removeChild(notification);
            }, 300);
        }, 3000);
    }

    setupEventListeners() {
        // Main download buttons
        const downloadButtons = document.querySelectorAll('#downloadBtn, #mainDownloadBtn');
        downloadButtons.forEach(button => {
            button.addEventListener('click', () => {
                this.handleDownload();
            });
        });

        // Demo buttons (just for show)
        const demoButtons = document.querySelectorAll('.demo-btn');
        demoButtons.forEach(button => {
            button.addEventListener('click', () => {
                this.showDemoMessage(button.textContent);
            });
        });
    }

    handleDownload() {
        // Increment counter
        this.incrementCount();

        // In a real scenario, you would trigger the actual download here
        // For now, we'll simulate it
        this.simulateDownload();
    }

    simulateDownload() {
        // Create a temporary link to simulate download
        const link = document.createElement('a');
        link.href = './Doc-smart.exe'; // Path to your exe file
        link.download = 'Doc-smart-v0.1.exe';
        
        // Try to download the actual file if it exists
        // If not, show a message
        link.addEventListener('error', () => {
            this.showNotification('Download simulation - In production, this would download Doc-smart.exe');
        });

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    showDemoMessage(buttonText) {
        const messages = {
            'Add Document': 'Opens dialog to add new Word documents from files or URLs',
            'Import Folder': 'Bulk import Word documents from selected folder',
            'Open Selected': 'Opens selected documents in Microsoft Word',
            'Close Selected': 'Closes selected documents in Word',
            'Mark Favorite': 'Toggles favorite status for quick access',
            'Remove Document': 'Removes documents from the manager',
            'Add Team': 'Creates new debate team for organization',
            'Rename Team': 'Renames existing team',
            'Delete Team': 'Removes team (documents become ungrouped)',
            'Open Team Docs': 'Opens all documents assigned to selected team',
            'Close All Open': 'Closes all currently open Word documents',
            'Export Data': 'Exports all data to JSON file for backup',
            'Toggle Favorites': 'Batch toggle favorite status for selected documents'
        };

        const message = messages[buttonText] || `${buttonText} functionality`;
        this.showNotification(message);
    }
}

// Smooth scrolling for navigation links
function setupSmoothScrolling() {
    const navLinks = document.querySelectorAll('a[href^="#"]');
    navLinks.forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const targetId = link.getAttribute('href');
            const targetElement = document.querySelector(targetId);
            
            if (targetElement) {
                const offsetTop = targetElement.offsetTop - 80; // Account for fixed header
                window.scrollTo({
                    top: offsetTop,
                    behavior: 'smooth'
                });
            }
        });
    });
}

// Add scroll effect to header
function setupHeaderScroll() {
    const header = document.querySelector('header');
    let lastScrollY = window.scrollY;

    window.addEventListener('scroll', () => {
        const currentScrollY = window.scrollY;
        
        if (currentScrollY > 100) {
            header.style.background = 'rgba(44, 62, 80, 0.95)';
            header.style.backdropFilter = 'blur(10px)';
        } else {
            header.style.background = '#2c3e50';
            header.style.backdropFilter = 'none';
        }

        lastScrollY = currentScrollY;
    });
}

// Animate elements on scroll
function setupScrollAnimations() {
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

    // Observe feature cards
    const featureCards = document.querySelectorAll('.feature-card');
    featureCards.forEach((card, index) => {
        card.style.opacity = '0';
        card.style.transform = 'translateY(30px)';
        card.style.transition = `opacity 0.6s ease ${index * 0.1}s, transform 0.6s ease ${index * 0.1}s`;
        observer.observe(card);
    });

    // Observe other sections
    const sections = document.querySelectorAll('.func-category, .download-card');
    sections.forEach(section => {
        section.style.opacity = '0';
        section.style.transform = 'translateY(30px)';
        section.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        observer.observe(section);
    });
}

// Initialize everything when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    // Initialize download counter
    new DownloadCounter();
    
    // Setup other functionality
    setupSmoothScrolling();
    setupHeaderScroll();
    setupScrollAnimations();
    
    // Add some interactive feedback to buttons
    const allButtons = document.querySelectorAll('button, .btn-secondary');
    allButtons.forEach(button => {
        button.addEventListener('mouseenter', () => {
            button.style.transform = 'translateY(-2px)';
        });
        
        button.addEventListener('mouseleave', () => {
            button.style.transform = 'translateY(0)';
        });
    });
});

// Add some stats animation
function animateStats() {
    const statsElement = document.getElementById('downloadCount');
    if (statsElement) {
        const finalCount = parseInt(statsElement.textContent);
        let currentCount = 0;
        const increment = Math.ceil(finalCount / 50);
        
        const timer = setInterval(() => {
            currentCount += increment;
            if (currentCount >= finalCount) {
                currentCount = finalCount;
                clearInterval(timer);
            }
            statsElement.textContent = currentCount.toLocaleString();
        }, 30);
    }
}

// Run stats animation after a short delay
setTimeout(animateStats, 1000);