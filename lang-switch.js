// 全局语言切换核心逻辑
(function() {
    // 1. 初始化语言（优先读取localStorage）
    let currentLang = localStorage.getItem('selectedLang') || 'zh';

    // 2. 获取翻译数据（支持两种格式：window.languages 和 languages）
    function getLangData(lang) {
        if (window.languages && window.languages[lang]) {
            return window.languages[lang];
        } else if (typeof languages !== 'undefined' && languages[lang]) {
            return languages[lang];
        }
        return {};
    }

    // 3. 动态创建语言切换按钮（如果不存在）
    function createLangSwitcher() {
        // 如果已存在，直接绑定点击事件
        if (document.querySelector('.language-switcher')) {
            bindLangClickEvent();
            return;
        }

        // 如果不存在，创建新的语言切换按钮
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

    // 4. 绑定语言切换点击事件
    function bindLangClickEvent() {
        document.querySelectorAll('.lang-btn').forEach(btn => {
            btn.removeEventListener('click', handleLangClick);
            btn.addEventListener('click', handleLangClick);
        });
    }

    // 5. 语言切换处理函数
    function handleLangClick() {
        currentLang = this.dataset.lang;
        localStorage.setItem('selectedLang', currentLang);
        updateLangBtnStatus();
        updatePageLang(); // 刷新页面所有翻译文字
    }

    // 6. 更新按钮激活状态
    function updateLangBtnStatus() {
        document.querySelectorAll('.lang-btn').forEach(btn => {
            btn.classList.remove('active');
            if (btn.dataset.lang === currentLang) {
                btn.classList.add('active');
            }
        });
    }

    // 7. 核心：更新页面所有翻译文字
    function updatePageLang() {
        const langData = getLangData(currentLang);
        
        // 遍历所有带data-i18n标记的元素
        const i18nElements = document.querySelectorAll('[data-i18n]');
        console.log('Found ' + i18nElements.length + ' elements with data-i18n attribute');
        
        i18nElements.forEach(el => {
            let key = el.dataset.i18n;
            // 移除 "common." 前缀
            if (key.startsWith('common.')) {
                key = key.substring(7);
            }
            if (langData[key]) {
                console.log('Updating element with key: ' + key + ' to: ' + langData[key]);
                // 对于optgroup元素，更新label属性
                if (el.tagName === 'OPTGROUP') {
                    el.label = langData[key];
                } else {
                    el.innerHTML = langData[key];
                }
            }
        });

        // 处理占位符（支持两种格式：data-i18nPlaceholder 和 data-i18n-placeholder）
        const placeholderElements = document.querySelectorAll('[data-i18nPlaceholder], [data-i18n-placeholder]');
        console.log('Found ' + placeholderElements.length + ' elements with placeholder attributes');
        
        placeholderElements.forEach(el => {
            let key = el.dataset.i18nPlaceholder || el.dataset['i18n-placeholder'];
            // 移除 "common." 前缀
            if (key && key.startsWith('common.')) {
                key = key.substring(7);
            }
            if (langData[key]) {
                console.log('Updating placeholder with key: ' + key + ' to: ' + langData[key]);
                el.placeholder = langData[key];
            }
        });
    }

    // 8. 初始化函数（页面加载后执行）
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

    // 监听DOM变化，支持动态加载的内容
    if (typeof MutationObserver !== 'undefined') {
        const observer = new MutationObserver(function(mutations) {
            mutations.forEach(function(mutation) {
                // 检查是否有新的元素被添加
                if (mutation.addedNodes.length > 0) {
                    // 延迟执行，确保新元素已经完全加载
                    setTimeout(updatePageLang, 100);
                }
            });
        });

        // 观察整个文档的变化
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    }

    // 暴露全局方法（可选，方便其他脚本调用）
    window.switchLanguage = function(lang) {
        currentLang = lang;
        localStorage.setItem('selectedLang', lang);
        updateLangBtnStatus();
        updatePageLang();
    };

    // 暴露全局方法，方便内容面板加载后调用
    window.updatePageLanguage = function() {
        updatePageLang();
    };
})();
