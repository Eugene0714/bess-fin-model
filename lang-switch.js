// 全局语言切换核心逻辑
(function() {
    // 1. 初始化语言（优先读取localStorage）
    let currentLang = localStorage.getItem('selectedLang') || 'zh';

    // 2. 动态创建语言切换按钮（无需手动加HTML）
    function createLangSwitcher() {
        // 如果已存在，不再重复创建
        if (document.querySelector('.language-switcher')) return;

        const switcher = document.createElement('div');
        switcher.className = 'language-switcher';
        switcher.innerHTML = `
            <button class="lang-btn" data-lang="zh">中文</button>
            <button class="lang-btn" data-lang="en">English</button>
            <button class="lang-btn" data-lang="de">Deutsch</button>
        `;
        document.body.appendChild(switcher);

        // 绑定点击事件
        bindLangClickEvent();
    }

    // 3. 绑定语言切换点击事件
    function bindLangClickEvent() {
        document.querySelectorAll('.lang-btn').forEach(btn => {
            btn.removeEventListener('click', handleLangClick);
            btn.addEventListener('click', handleLangClick);
        });
    }

    // 4. 语言切换处理函数
    function handleLangClick() {
        currentLang = this.dataset.lang;
        localStorage.setItem('selectedLang', currentLang);
        updateLangBtnStatus();
        updatePageLang(); // 刷新页面所有翻译文字
    }

    // 5. 更新按钮激活状态
    function updateLangBtnStatus() {
        document.querySelectorAll('.lang-btn').forEach(btn => {
            btn.classList.remove('active');
            if (btn.dataset.lang === currentLang) {
                btn.classList.add('active');
            }
        });
    }

    // 6. 核心：更新页面所有翻译文字
    function updatePageLang() {
        // 遍历所有带data-i18n标记的元素
        document.querySelectorAll('[data-i18n]').forEach(el => {
            const key = el.dataset.i18n;
            const [module, field] = key.split('.'); // 格式：index.dashboardTitle

            // 从全局翻译表取值
            if (window.langMap[module] && window.langMap[module][field]) {
                el.innerHTML = window.langMap[module][field][currentLang];
            }

            // 处理占位符（如果有）
            if (el.dataset.i18nPlaceholder) {
                const [phModule, phField] = el.dataset.i18nPlaceholder.split('.');
                if (window.langMap[phModule] && window.langMap[phModule][phField]) {
                    el.placeholder = window.langMap[phModule][phField][currentLang];
                }
            }
        });
    }

    // 7. 初始化函数（页面加载后执行）
    function initLangSwitch() {
        createLangSwitcher(); // 动态创建按钮
        updateLangBtnStatus(); // 初始化按钮状态
        updatePageLang(); // 初始化页面文字
    }

    // 页面加载完成后初始化
    if (document.readyState === 'complete') {
        initLangSwitch();
    } else {
        window.addEventListener('load', initLangSwitch);
    }

    // 暴露全局方法（可选，方便其他脚本调用）
    window.switchLanguage = function(lang) {
        currentLang = lang;
        localStorage.setItem('selectedLang', lang);
        updateLangBtnStatus();
        updatePageLang();
    };
})();
