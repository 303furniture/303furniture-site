// ハンバーガーメニュー
const hamburger = document.querySelector('.hamburger');
const mobileMenu = document.querySelector('.mobile-menu');

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

// ヘッダーのスクロール効果
window.addEventListener('scroll', () => {
    const header = document.querySelector('header');
    if (window.scrollY > 50) {
        header.classList.add('scrolled');
    } else {
        header.classList.remove('scrolled');
    }
});

// スクロール時のフェードインアニメーション
const observerOptions = {
    root: null,
    rootMargin: '0px',
    threshold: 0.15
};

const observer = new IntersectionObserver((entries, observer) => {
    entries.forEach(entry => {
        if (entry.isIntersecting) {
            entry.target.classList.add('is-visible');
            // 一度発火したら監視を解除する場合（解除しない場合はコメントアウトを外さない）
            // observer.unobserve(entry.target);
        }
    });
}, observerOptions);

document.querySelectorAll('.fade-in-up').forEach(el => {
    observer.observe(el);
});
