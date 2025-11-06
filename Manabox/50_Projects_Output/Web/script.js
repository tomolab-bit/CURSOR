// モバイルメニューの制御
document.addEventListener('DOMContentLoaded', function() {
    const navToggle = document.getElementById('nav-toggle');
    const navMenu = document.getElementById('nav-menu');

    navToggle.addEventListener('click', function() {
        navMenu.classList.toggle('active');
        navToggle.classList.toggle('active');
    });

    // メニューリンクをクリックしたときにメニューを閉じる
    const navLinks = document.querySelectorAll('.nav-link');
    navLinks.forEach(link => {
        link.addEventListener('click', function() {
            navMenu.classList.remove('active');
            navToggle.classList.remove('active');
        });
    });

    // メニュー外をクリックしたときにメニューを閉じる
    document.addEventListener('click', function(event) {
        if (!navToggle.contains(event.target) && !navMenu.contains(event.target)) {
            navMenu.classList.remove('active');
            navToggle.classList.remove('active');
        }
    });
});

// スクロールアニメーション
function observeElements() {
    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('visible');
            }
        });
    }, {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    });

    // アニメーション対象の要素を取得
    const animatedElements = document.querySelectorAll('.service-card, .trial-card, .project-card, .about-content, .contact-content');
    
    animatedElements.forEach(el => {
        el.classList.add('fade-in');
        observer.observe(el);
    });
}

// ページ読み込み完了後にアニメーションを開始
document.addEventListener('DOMContentLoaded', observeElements);

// スムーススクロール
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            const headerHeight = document.querySelector('.header').offsetHeight;
            const targetPosition = target.offsetTop - headerHeight - 20;
            
            window.scrollTo({
                top: targetPosition,
                behavior: 'smooth'
            });
        }
    });
});

// ヘッダーの背景変更（スクロール時）
window.addEventListener('scroll', function() {
    const header = document.querySelector('.header');
    if (window.scrollY > 100) {
        header.style.background = 'rgba(255, 255, 255, 0.98)';
        header.style.boxShadow = '0 2px 20px rgba(0, 0, 0, 0.15)';
    } else {
        header.style.background = 'rgba(255, 255, 255, 0.95)';
        header.style.boxShadow = '0 2px 20px rgba(0, 0, 0, 0.1)';
    }
});

// フォーム送信処理
document.addEventListener('DOMContentLoaded', function() {
    const form = document.querySelector('.form');
    if (form) {
        form.addEventListener('submit', function(e) {
            e.preventDefault();
            
            // フォームデータを取得
            const formData = new FormData(form);
            const data = Object.fromEntries(formData);
            
            // バリデーション
            if (!data.name || !data.email || !data.message) {
                alert('必須項目を入力してください。');
                return;
            }
            
            // メールアドレスの形式チェック
            const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
            if (!emailRegex.test(data.email)) {
                alert('正しいメールアドレスを入力してください。');
                return;
            }
            
            // 送信処理（実際の実装ではサーバーサイドで処理）
            showSuccessMessage();
        });
    }
});

// 成功メッセージ表示
function showSuccessMessage() {
    const form = document.querySelector('.form');
    const successMessage = document.createElement('div');
    successMessage.className = 'success-message';
    successMessage.innerHTML = `
        <div style="
            background: #d4edda;
            color: #155724;
            padding: 1rem;
            border-radius: 10px;
            margin-top: 1rem;
            text-align: center;
            border: 1px solid #c3e6cb;
        ">
            <i class="fas fa-check-circle" style="margin-right: 0.5rem;"></i>
            お問い合わせありがとうございます。内容を確認次第、担当者よりご連絡いたします。
        </div>
    `;
    
    form.appendChild(successMessage);
    form.reset();
    
    // 3秒後にメッセージを削除
    setTimeout(() => {
        successMessage.remove();
    }, 5000);
}

// カウンターアニメーション
function animateCounters() {
    const counters = document.querySelectorAll('.stat-number');
    
    counters.forEach(counter => {
        const target = parseInt(counter.textContent.replace('+', ''));
        const increment = target / 100;
        let current = 0;
        
        const updateCounter = () => {
            if (current < target) {
                current += increment;
                counter.textContent = Math.ceil(current) + (counter.textContent.includes('+') ? '+' : '');
                requestAnimationFrame(updateCounter);
            } else {
                counter.textContent = target + (counter.textContent.includes('+') ? '+' : '');
            }
        };
        
        // Intersection Observerでカウンターが表示されたときにアニメーション開始
        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    updateCounter();
                    observer.unobserve(entry.target);
                }
            });
        });
        
        observer.observe(counter);
    });
}

// ページ読み込み完了後にカウンターアニメーションを開始
document.addEventListener('DOMContentLoaded', animateCounters);

// パフォーマンス最適化: 画像の遅延読み込み
function lazyLoadImages() {
    const images = document.querySelectorAll('img[data-src]');
    
    const imageObserver = new IntersectionObserver((entries, observer) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const img = entry.target;
                img.src = img.dataset.src;
                img.classList.remove('lazy');
                imageObserver.unobserve(img);
            }
        });
    });
    
    images.forEach(img => imageObserver.observe(img));
}

// ページ読み込み完了後に遅延読み込みを開始
document.addEventListener('DOMContentLoaded', lazyLoadImages);

// アクセシビリティ向上: キーボードナビゲーション
document.addEventListener('keydown', function(e) {
    // ESCキーでメニューを閉じる
    if (e.key === 'Escape') {
        const navMenu = document.getElementById('nav-menu');
        const navToggle = document.getElementById('nav-toggle');
        navMenu.classList.remove('active');
        navToggle.classList.remove('active');
    }
});

// パフォーマンス監視
window.addEventListener('load', function() {
    // ページ読み込み時間の測定
    const loadTime = performance.now();
    console.log(`ページ読み込み時間: ${Math.round(loadTime)}ms`);
    
    // 大きな画像の検出と警告
    const images = document.querySelectorAll('img');
    images.forEach(img => {
        if (img.naturalWidth > 1920 || img.naturalHeight > 1080) {
            console.warn(`大きな画像が検出されました: ${img.src} (${img.naturalWidth}x${img.naturalHeight})`);
        }
    });
});

// エラーハンドリング
window.addEventListener('error', function(e) {
    console.error('JavaScript エラー:', e.error);
});

// サービスワーカーの登録（PWA対応）
if ('serviceWorker' in navigator) {
    window.addEventListener('load', function() {
        navigator.serviceWorker.register('/sw.js')
            .then(function(registration) {
                console.log('ServiceWorker 登録成功:', registration.scope);
            })
            .catch(function(error) {
                console.log('ServiceWorker 登録失敗:', error);
            });
    });
}

