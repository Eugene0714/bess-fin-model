// 全局语言映射表（包含登录/注册/首页所有翻译文案）
window.langMap = {
    // 通用文案（登录/注册/首页都用得到）
    common: {
        logout: { zh: "退出登录", en: "Logout", de: "Abmelden" },
        save: { zh: "保存", en: "Save", de: "Speichern" },
        cancel: { zh: "取消", en: "Cancel", de: "Abbrechen" },
        confirm: { zh: "确认", en: "Confirm", de: "Bestätigen" }
    },
    // 首页专属文案
    index: {
        dashboardTitle: { zh: "储能投资分析仪表盘", en: "Energy Storage Investment Analysis Dashboard", de: "Energiespeicher-Investitionsanalyse-Dashboard" },
        projectList: { zh: "项目列表", en: "Project List", de: "Projektliste" },
        newProject: { zh: "新建项目", en: "New Project", de: "Neues Projekt" },
        dataAnalysis: { zh: "数据分析", en: "Data Analysis", de: "Datenanalyse" },
        reportExport: { zh: "报告导出", en: "Report Export", de: "Berichtsexport" },
        userCenter: { zh: "个人中心", en: "User Center", de: "Benutzerzentrum" },
        projectName: { zh: "项目名称", en: "Project Name", de: "Projektname" },
        investmentAmount: { zh: "投资金额", en: "Investment Amount", de: "Investitionsbetrag" },
        irr: { zh: "内部收益率", en: "IRR", de: "Interner Zinssatz" },
        npv: { zh: "净现值", en: "NPV", de: "Nettobarwert" },
        // 补充首页其他需要翻译的文案...
    },
    // 登录/注册页文案（已有的可以保留，确保全站统一）
    login: { /* 原有登录页翻译... */ },
    register: { /* 原有注册页翻译... */ }
};
