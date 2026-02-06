/**
 * API Service - 德国独立储能电站投资测算系统
 * 对接 Replit 后端真实接口 + 保留前端认证模拟
 * 
 * 真实后端基础路径: https://bess-fin-backend--1423127256.replit.app
 * 认证相关逻辑仍使用前端模拟
 * 
 * @version 1.0.0
 * @author TEMAX Energy Solutions
 */

// ==================== 核心配置 ====================
// 替换为你的 Replit 后端地址
const BACKEND_BASE_URL = 'https://bess-fin-backend--1423127256.replit.app';

/**
 * 本地存储键名
 */
const STORAGE_KEYS = {
    USERS: 'api_users',
    VERIFY_CODES: 'api_verify_codes',
    ORDERS: 'api_orders',
    PLANS: 'api_plans',
    TOKENS: 'api_tokens'
};

// ==================== 原有认证逻辑（保留不变） ====================
function initPlans() {
    if (!localStorage.getItem(STORAGE_KEYS.PLANS)) {
        const plans = [
            {
                plan_id: 1,
                name: '免费版',
                code: 'free',
                price: 0,
                duration: 'forever',
                features: [
                    '基础财务计算',
                    '单项目限制',
                    '社区支持'
                ],
                api_calls: 100,
                concurrent: 1
            },
            {
                plan_id: 2,
                name: '专业版',
                code: 'pro',
                price: 99.00,
                duration: 'month',
                features: [
                    '完整财务模型',
                    '无限项目',
                    '敏感性分析',
                    '融资报告生成',
                    '优先技术支持'
                ],
                api_calls: 10000,
                concurrent: 5
            },
            {
                plan_id: 3,
                name: '企业版',
                code: 'enterprise',
                price: 499.00,
                duration: 'month',
                features: [
                    '专业版全部功能',
                    '多用户协作',
                    'API完全访问',
                    '专属客户经理',
                    '定制化开发支持'
                ],
                api_calls: -1, // unlimited
                concurrent: -1 // unlimited
            }
        ];
        localStorage.setItem(STORAGE_KEYS.PLANS, JSON.stringify(plans));
    }
}

function generateToken(userId) {
    const header = btoa(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
    const payload = btoa(JSON.stringify({
        user_id: userId,
        iat: Date.now(),
        exp: Date.now() + 24 * 60 * 60 * 1000 // 24小时有效期
    }));
    const signature = btoa(Math.random().toString(36).substring(2));
    return `${header}.${payload}.${signature}`;
}

function verifyToken(token) {
    if (!token) return null;
    
    try {
        const parts = token.split('.');
        if (parts.length !== 3) return null;
        
        const payload = JSON.parse(atob(parts[1]));
        
        if (payload.exp < Date.now()) return null;
        
        const users = JSON.parse(localStorage.getItem(STORAGE_KEYS.USERS) || '[]');
        return users.find(u => u.user_id === payload.user_id) || null;
    } catch (e) {
        return null;
    }
}

function generateVerifyCode() {
    return Math.floor(100000 + Math.random() * 900000).toString();
}

function generateOrderId() {
    const timestamp = Date.now().toString(36).toUpperCase();
    const random = Math.random().toString(36).substring(2, 6).toUpperCase();
    return `ORD${timestamp}${random}`;
}

function hashPassword(password) {
    return btoa(password + '_salt_temax');
}

function verifyPassword(password, hash) {
    return hashPassword(password) === hash;
}

// ==================== 原有认证API（保留不变） ====================
class AuthMock {
    constructor() {
        initPlans();
        this.rateLimits = {};
    }

    async sendCode(body) {
        const { email } = body;
        
        if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
            return {
                status: 400,
                data: { error: 'Invalid email format' }
            };
        }
        
        const now = Date.now();
        const lastSent = this.rateLimits[`code_${email}`];
        if (lastSent && now - lastSent < 60000) {
            const waitTime = Math.ceil((60000 - (now - lastSent)) / 1000);
            return {
                status: 429,
                data: { error: `Rate limit exceeded. Please wait ${waitTime} seconds.` }
            };
        }
        
        const code = generateVerifyCode();
        const codes = JSON.parse(localStorage.getItem(STORAGE_KEYS.VERIFY_CODES) || '{}');
        codes[email] = {
            code,
            expires: Date.now() + 5 * 60 * 1000
        };
        localStorage.setItem(STORAGE_KEYS.VERIFY_CODES, JSON.stringify(codes));
        
        this.rateLimits[`code_${email}`] = now;
        
        await new Promise(r => setTimeout(r, 500));
        
        console.log(`[API Mock] 验证码已发送到 ${email}: ${code}`);
        
        return {
            status: 200,
            data: { message: 'Code sent successfully.' }
        };
    }

    async register(body) {
        const { email, password, code, username, plan_id = 1 } = body;
        
        if (!email || !password || !code) {
            return {
                status: 400,
                data: { error: 'Missing required fields' }
            };
        }
        
        const codes = JSON.parse(localStorage.getItem(STORAGE_KEYS.VERIFY_CODES) || '{}');
        const storedCode = codes[email];
        
        if (!storedCode || storedCode.code !== code || storedCode.expires < Date.now()) {
            return {
                status: 400,
                data: { error: 'Invalid or expired verification code' }
            };
        }
        
        const users = JSON.parse(localStorage.getItem(STORAGE_KEYS.USERS) || '[]');
        if (users.some(u => u.email === email)) {
            return {
                status: 409,
                data: { error: 'Email already registered' }
            };
        }
        
        if (password.length < 8) {
            return {
                status: 400,
                data: { error: 'Password must be at least 8 characters' }
            };
        }
        
        const user = {
            user_id: `USR${Date.now().toString(36).toUpperCase()}`,
            email,
            username: username || email.split('@')[0],
            password_hash: hashPassword(password),
            current_plan_id: plan_id,
            plan_expiry_date: plan_id === 1 ? null : new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString(),
            created_at: new Date().toISOString(),
            status: 'active'
        };
        
        users.push(user);
        localStorage.setItem(STORAGE_KEYS.USERS, JSON.stringify(users));
        
        delete codes[email];
        localStorage.setItem(STORAGE_KEYS.VERIFY_CODES, JSON.stringify(codes));
        
        const token = generateToken(user.user_id);
        
        return {
            status: 201,
            data: {
                token,
                user_id: user.user_id,
                message: 'Registration successful'
            }
        };
    }

    async login(body) {
        const { email, password } = body;
        
        if (!email || !password) {
            return {
                status: 400,
                data: { error: 'Email and password required' }
            };
        }
        
        const now = Date.now();
        const attempts = this.rateLimits[`login_${email}`] || [];
        const recentAttempts = attempts.filter(t => now - t < 60000);
        
        if (recentAttempts.length >= 5) {
            return {
                status: 429,
                data: { error: 'Too many login attempts. Please try again later.' }
            };
        }
        
        const users = JSON.parse(localStorage.getItem(STORAGE_KEYS.USERS) || '[]');
        const user = users.find(u => u.email === email);
        
        if (!user || !verifyPassword(password, user.password_hash)) {
            this.rateLimits[`login_${email}`] = [...recentAttempts, now];
            
            return {
                status: 401,
                data: { error: 'Invalid email or password' }
            };
        }
        
        delete this.rateLimits[`login_${email}`];
        
        const token = generateToken(user.user_id);
        
        return {
            status: 200,
            data: {
                token,
                user_id: user.user_id
            }
        };
    }

    async getProfile(token) {
        const user = verifyToken(token);
        
        if (!user) {
            return {
                status: 401,
                data: { error: 'Unauthorized' }
            };
        }
        
        const plans = JSON.parse(localStorage.getItem(STORAGE_KEYS.PLANS) || '[]');
        const currentPlan = plans.find(p => p.plan_id === user.current_plan_id);
        
        return {
            status: 200,
            data: {
                user_id: user.user_id,
                email: user.email,
                username: user.username,
                current_plan_id: user.current_plan_id,
                current_plan: currentPlan,
                plan_expiry_date: user.plan_expiry_date,
                created_at: user.created_at
            }
        };
    }

    async getPlans() {
        const plans = JSON.parse(localStorage.getItem(STORAGE_KEYS.PLANS) || '[]');
        
        return {
            status: 200,
            data: plans
        };
    }

    async createOrder(body, token) {
        const user = verifyToken(token);
        
        if (!user) {
            return {
                status: 401,
                data: { error: 'Unauthorized' }
            };
        }
        
        const { plan_id, duration = 'month', payment_method = 'wechat' } = body;
        
        const plans = JSON.parse(localStorage.getItem(STORAGE_KEYS.PLANS) || '[]');
        const plan = plans.find(p => p.plan_id === plan_id);
        
        if (!plan) {
            return {
                status: 404,
                data: { error: 'Plan not found' }
            };
        }
        
        if (plan.price === 0) {
            return {
                status: 400,
                data: { error: 'Cannot create order for free plan' }
            };
        }
        
        let amount = plan.price;
        let months = 1;
        
        if (duration === 'quarter') {
            amount = plan.price * 3 * 0.9;
            months = 3;
        } else if (duration === 'year') {
            amount = plan.price * 12 * 0.8;
            months = 12;
        }
        
        const order = {
            order_id: generateOrderId(),
            user_id: user.user_id,
            plan_id,
            plan_name: plan.name,
            amount: Math.round(amount * 100) / 100,
            duration,
            months,
            payment_method,
            status: 'pending',
            created_at: new Date().toISOString(),
            expires_at: new Date(Date.now() + 30 * 60 * 1000).toISOString()
        };
        
        const orders = JSON.parse(localStorage.getItem(STORAGE_KEYS.ORDERS) || '[]');
        orders.push(order);
        localStorage.setItem(STORAGE_KEYS.ORDERS, JSON.stringify(orders));
        
        const prepay_url = `https://pay.example.com/${payment_method}?order=${order.order_id}`;
        
        return {
            status: 201,
            data: {
                order_id: order.order_id,
                prepay_url,
                amount: order.amount,
                status: order.status
            }
        };
    }

    async paymentNotify(body) {
        const { order_id, transaction_id, status = 'success' } = body;
        
        const orders = JSON.parse(localStorage.getItem(STORAGE_KEYS.ORDERS) || '[]');
        const orderIndex = orders.findIndex(o => o.order_id === order_id);
        
        if (orderIndex === -1) {
            return {
                status: 404,
                data: { error: 'Order not found' }
            };
        }
        
        const order = orders[orderIndex];
        
        if (order.status !== 'pending') {
            return {
                status: 400,
                data: { error: 'Order already processed' }
            };
        }
        
        if (status === 'success') {
            order.status = 'paid';
            order.transaction_id = transaction_id || `TXN${Date.now()}`;
            order.paid_at = new Date().toISOString();
            orders[orderIndex] = order;
            localStorage.setItem(STORAGE_KEYS.ORDERS, JSON.stringify(orders));
            
            const users = JSON.parse(localStorage.getItem(STORAGE_KEYS.USERS) || '[]');
            const userIndex = users.findIndex(u => u.user_id === order.user_id);
            
            if (userIndex !== -1) {
                const user = users[userIndex];
                const currentExpiry = user.plan_expiry_date ? new Date(user.plan_expiry_date) : new Date();
                const baseDate = currentExpiry > new Date() ? currentExpiry : new Date();
                
                const newExpiry = new Date(baseDate);
                newExpiry.setMonth(newExpiry.getMonth() + order.months);
                
                user.current_plan_id = order.plan_id;
                user.plan_expiry_date = newExpiry.toISOString();
                users[userIndex] = user;
                localStorage.setItem(STORAGE_KEYS.USERS, JSON.stringify(users));
            }
            
            return {
                status: 200,
                data: { status: 'success' }
            };
        } else {
            order.status = 'failed';
            orders[orderIndex] = order;
            localStorage.setItem(STORAGE_KEYS.ORDERS, JSON.stringify(orders));
            
            return {
                status: 200,
                data: { status: 'failed' }
            };
        }
    }

    async getOrder(orderId, token) {
        const user = verifyToken(token);
        
        if (!user) {
            return {
                status: 401,
                data: { error: 'Unauthorized' }
            };
        }
        
        const orders = JSON.parse(localStorage.getItem(STORAGE_KEYS.ORDERS) || '[]');
        const order = orders.find(o => o.order_id === orderId && o.user_id === user.user_id);
        
        if (!order) {
            return {
                status: 404,
                data: { error: 'Order not found' }
            };
        }
        
        if (order.status === 'pending' && new Date(order.expires_at) < new Date()) {
            order.status = 'closed';
            localStorage.setItem(STORAGE_KEYS.ORDERS, JSON.stringify(orders));
        }
        
        return {
            status: 200,
            data: {
                order_id: order.order_id,
                status: order.status,
                plan_id: order.plan_id,
                amount: order.amount,
                expiry_date: order.status === 'paid' ? 
                    new Date(new Date().setMonth(new Date().getMonth() + order.months)).toISOString() : null
            }
        };
    }

    async checkAccess(token, feature) {
        const user = verifyToken(token);
        
        if (!user) {
            return {
                status: 401,
                data: { error: 'Unauthorized', has_access: false }
            };
        }
        
        if (user.current_plan_id > 1 && user.plan_expiry_date) {
            if (new Date(user.plan_expiry_date) < new Date()) {
                return {
                    status: 403,
                    data: { 
                        error: 'Subscription expired', 
                        has_access: false,
                        message: '您的订阅已过期，请续费后继续使用'
                    }
                };
            }
        }
        
        const plans = JSON.parse(localStorage.getItem(STORAGE_KEYS.PLANS) || '[]');
        const currentPlan = plans.find(p => p.plan_id === user.current_plan_id);
        
        const featureAccess = {
            'sensitivity_analysis': [2, 3],
            'financing_report': [2, 3],
            'api_access': [3],
            'multi_user': [3],
            'basic_calculation': [1, 2, 3]
        };
        
        const allowedPlans = featureAccess[feature] || [];
        const hasAccess = allowedPlans.includes(user.current_plan_id);
        
        if (!hasAccess) {
            return {
                status: 403,
                data: {
                    error: 'Feature not available',
                    has_access: false,
                    required_plan: allowedPlans[0],
                    message: '此功能需要升级到更高版本'
                }
            };
        }
        
        return {
            status: 200,
            data: {
                has_access: true,
                plan: currentPlan
            }
        };
    }
}

// ==================== 新增：对接真实后端的API ====================
class BackendAPI {
    constructor() {
        this.baseUrl = BACKEND_BASE_URL;
    }

    /**
     * 测试后端连接
     */
    async testConnection() {
        try {
            const response = await fetch(`${this.baseUrl}/test`);
            const data = await response.json();
            return {
                status: response.status,
                data
            };
        } catch (error) {
            console.error('后端连接失败:', error);
            return {
                status: 500,
                data: { error: 'Failed to connect to backend', message: '后端服务暂时不可用' }
            };
        }
    }

    /**
     * 生成财务测算Excel
     * @param {Object} params - 测算参数
     * @param {string} token - 认证Token
     */
    async generateExcel(params, token) {
        // 先检查权限
        const authMock = new AuthMock();
        const accessCheck = await authMock.checkAccess(token, 'basic_calculation');
        
        if (accessCheck.status !== 200) {
            return accessCheck;
        }

        try {
            const response = await fetch(`${this.baseUrl}/generate-excel`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${token}`
                },
                body: JSON.stringify(params)
            });
            
            const data = await response.json();
            return {
                status: response.status,
                data
            };
        } catch (error) {
            console.error('生成Excel失败:', error);
            return {
                status: 500,
                data: { error: 'Failed to generate Excel', message: 'Excel生成失败，请重试' }
            };
        }
    }

    /**
     * 同步Excel数据
     * @param {Object} params - 同步参数
     * @param {string} token - 认证Token
     */
    async syncExcel(params, token) {
        // 先检查权限
        const authMock = new AuthMock();
        const accessCheck = await authMock.checkAccess(token, 'basic_calculation');
        
        if (accessCheck.status !== 200) {
            return accessCheck;
        }

        try {
            const response = await fetch(`${this.baseUrl}/sync-excel`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${token}`
                },
                body: JSON.stringify(params)
            });
            
            const data = await response.json();
            return {
                status: response.status,
                data
            };
        } catch (error) {
            console.error('同步Excel失败:', error);
            return {
                status: 500,
                data: { error: 'Failed to sync Excel', message: '数据同步失败，请重试' }
            };
        }
    }
}

// ==================== 整合API客户端 ====================
class APIClient {
    constructor() {
        this.authMock = new AuthMock();
        this.backendAPI = new BackendAPI();
    }

    getToken() {
        return sessionStorage.getItem('auth_token');
    }

    setToken(token) {
        sessionStorage.setItem('auth_token', token);
    }

    clearToken() {
        sessionStorage.removeItem('auth_token');
    }

    // 认证相关（沿用原有模拟逻辑）
    async sendCode(email) {
        return this.authMock.sendCode({ email });
    }

    async register(data) {
        const result = await this.authMock.register(data);
        if (result.status === 201) {
            this.setToken(result.data.token);
        }
        return result;
    }

    async login(email, password) {
        const result = await this.authMock.login({ email, password });
        if (result.status === 200) {
            this.setToken(result.data.token);
        }
        return result;
    }

    logout() {
        this.clearToken();
    }

    async getProfile() {
        return this.authMock.getProfile(this.getToken());
    }

    async getPlans() {
        return this.authMock.getPlans();
    }

    async createOrder(planId, duration, paymentMethod) {
        return this.authMock.createOrder(
            { plan_id: planId, duration, payment_method: paymentMethod },
            this.getToken()
        );
    }

    async simulatePayment(orderId) {
        return this.authMock.paymentNotify({
            order_id: orderId,
            status: 'success'
        });
    }

    async getOrder(orderId) {
        return this.authMock.getOrder(orderId, this.getToken());
    }

    async checkAccess(feature) {
        return this.authMock.checkAccess(this.getToken(), feature);
    }

    isLoggedIn() {
        const token = this.getToken();
        return token && verifyToken(token) !== null;
    }

    // 新增：对接真实后端的方法
    async testBackend() {
        return this.backendAPI.testConnection();
    }

    async generateExcel(params) {
        return this.backendAPI.generateExcel(params, this.getToken());
    }

    async syncExcel(params) {
        return this.backendAPI.syncExcel(params, this.getToken());
    }
}

// ==================== 导出 ====================
const api = new APIClient();

if (typeof module !== 'undefined' && module.exports) {
    module.exports = { AuthMock, BackendAPI, APIClient, api };
}

if (typeof window !== 'undefined') {
    window.AuthMock = AuthMock;
    window.BackendAPI = BackendAPI;
    window.APIClient = APIClient;
    window.api = api;
}

console.log('[API Service] 已对接Replit后端:', BACKEND_BASE_URL);
console.log('[API Service] 测试连接: api.testBackend()');
console.log('[API Service] 生成Excel: api.generateExcel({...参数})');
