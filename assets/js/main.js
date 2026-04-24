// ハンバーガーメニュー
const hamburger = document.querySelector('.hamburger');
const mobileMenu = document.querySelector('.mobile-menu');

if (hamburger && mobileMenu) {
    hamburger.addEventListener('click', () => {
        hamburger.classList.toggle('active');
        mobileMenu.classList.toggle('active');
    });

    // モバイルメニューのリンククリック時にメニューを閉じる
    document.querySelectorAll('.mobile-menu a').forEach(link => {
        link.addEventListener('click', () => {
            hamburger.classList.remove('active');
            mobileMenu.classList.remove('active');
        });
    });
}

// ヘッダーのスクロール効果（passive: true でパフォーマンス向上）
window.addEventListener('scroll', () => {
    const header = document.querySelector('header');
    if (window.scrollY > 50) {
        header.classList.add('scrolled');
    } else {
        header.classList.remove('scrolled');
    }
}, { passive: true });

// スクロール時のフェードインアニメーション
const observerOptions = {
    root: null,
    rootMargin: '0px 0px -30px 0px',
    threshold: 0.1
};

const fadeInTargets = document.querySelectorAll('.fade-in-up');

const showFadeInTargets = () => {
    fadeInTargets.forEach(el => {
        el.classList.add('is-visible');
    });
};

if ('IntersectionObserver' in window) {
    window.setTimeout(showFadeInTargets, 1500);

    const observer = new IntersectionObserver((entries, observer) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('is-visible');
                // 一度表示されたら監視を解除（パフォーマンス向上）
                observer.unobserve(entry.target);
            }
        });
    }, observerOptions);

    fadeInTargets.forEach(el => {
        observer.observe(el);
    });
} else {
    showFadeInTargets();
}
