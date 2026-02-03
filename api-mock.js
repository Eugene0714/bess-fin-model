/**
 * API Mock Service - 德国独立储能电站投资测算系统
 * 
 * 本文件模拟后端API服务，实现用户注册、登录、订阅付费等功能
 * 基于 前后端开发逻辑说明 (API & Logic Specification) 文档设计
 * 
 * API基础路径: /api/v1
 * 
 * @version 1.0.0
 * @author TEMAX Energy Solutions
 */

// ==================== 数据存储模拟 ====================

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

/**
 * 初始化订阅方案数据
 */
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

// ==================== 工具函数 ====================

/**
 * 生成JWT Token（模拟）
 * @param {string} userId - 用户ID
 * @returns {string} 模拟的JWT Token
 */
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

/**
 * 验证Token
 * @param {string} token - JWT Token
 * @returns {Object|null} 用户信息或null
 */
function verifyToken(token) {
    if (!token) return null;
    
    try {
        const parts = token.split('.');
        if (parts.length !== 3) return null;
        
        const payload = JSON.parse(atob(parts[1]));
        
        // 检查是否过期
        if (payload.exp < Date.now()) return null;
        
        const users = JSON.parse(localStorage.getItem(STORAGE_KEYS.USERS) || '[]');
        return users.find(u => u.user_id === payload.user_id) || null;
    } catch (e) {
        return null;
    }
}

/**
 * 生成6位验证码
 * @returns {string} 6位数字验证码
 */
function generateVerifyCode() {
    return Math.floor(100000 + Math.random() * 900000).toString();
}

/**
 * 生成订单号
 * @returns {string} 订单号
 */
function generateOrderId() {
    const timestamp = Date.now().toString(36).toUpperCase();
    const random = Math.random().toString(36).substring(2, 6).toUpperCase();
    return `ORD${timestamp}${random}`;
}

/**
 * 密码哈希（模拟Bcrypt）
 * @param {string} password - 明文密码
 * @returns {string} 哈希后的密码
 */
function hashPassword(password) {
    // 实际应用中使用Bcrypt或Argon2
    return btoa(password + '_salt_temax');
}

/**
 * 验证密码
 * @param {string} password - 明文密码
 * @param {string} hash - 哈希值
 * @returns {boolean} 是否匹配
 */
function verifyPassword(password, hash) {
    return hashPassword(password) === hash;
}

// ==================== API 实现 ====================

/**
 * API Mock 服务类
 */
class APIMock {
    constructor() {
        initPlans();
        this.rateLimits = {}; // IP级别速率限制
    }

    // ==================== A. 注册与认证 API ====================

    /**
     * POST /api/v1/auth/send-code
     * 发送邮箱验证码
     * 
     * @param {Object} body - 请求体 { email: string }
     * @returns {Promise<Object>} 响应
     */
    async sendCode(body) {
        const { email } = body;
        
        // 验证邮箱格式
        if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
            return {
                status: 400,
                data: { error: 'Invalid email format' }
            };
        }
        
        // 速率限制检查（每分钟最多1次）
        const now = Date.now();
        const lastSent = this.rateLimits[`code_${email}`];
        if (lastSent && now - lastSent < 60000) {
            const waitTime = Math.ceil((60000 - (now - lastSent)) / 1000);
            return {
                status: 429,
                data: { error: `Rate limit exceeded. Please wait ${waitTime} seconds.` }
            };
        }
        
        // 生成验证码
        const code = generateVerifyCode();
        const codes = JSON.parse(localStorage.getItem(STORAGE_KEYS.VERIFY_CODES) || '{}');
        codes[email] = {
            code,
            expires: Date.now() + 5 * 60 * 1000 // 5分钟有效
        };
        localStorage.setItem(STORAGE_KEYS.VERIFY_CODES, JSON.stringify(codes));
        
        // 记录速率限制
        this.rateLimits[`code_${email}`] = now;
        
        // 模拟发送延迟
        await new Promise(r => setTimeout(r, 500));
        
        console.log(`[API Mock] 验证码已发送到 ${email}: ${code}`);
        
        return {
            status: 200,
            data: { message: 'Code sent successfully.' }
        };
    }

    /**
     * POST /api/v1/auth/register
     * 用户注册
     * 
     * @param {Object} body - 请求体 { email, password, code, username?, plan_id? }
     * @returns {Promise<Object>} 响应
     */
    async register(body) {
        const { email, password, code, username, plan_id = 1 } = body;
        
        // 参数验证
        if (!email || !password || !code) {
            return {
                status: 400,
                data: { error: 'Missing required fields' }
            };
        }
        
        // 验证验证码
        const codes = JSON.parse(localStorage.getItem(STORAGE_KEYS.VERIFY_CODES) || '{}');
        const storedCode = codes[email];
        
        if (!storedCode || storedCode.code !== code || storedCode.expires < Date.now()) {
            return {
                status: 400,
                data: { error: 'Invalid or expired verification code' }
            };
        }
        
        // 检查邮箱是否已注册
        const users = JSON.parse(localStorage.getItem(STORAGE_KEYS.USERS) || '[]');
        if (users.some(u => u.email === email)) {
            return {
                status: 409,
                data: { error: 'Email already registered' }
            };
        }
        
        // 密码强度验证
        if (password.length < 8) {
            return {
                status: 400,
                data: { error: 'Password must be at least 8 characters' }
            };
        }
        
        // 创建用户
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
        
        // 清除验证码
        delete codes[email];
        localStorage.setItem(STORAGE_KEYS.VERIFY_CODES, JSON.stringify(codes));
        
        // 生成Token
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

    /**
     * POST /api/v1/auth/login
     * 用户登录
     * 
     * @param {Object} body - 请求体 { email, password }
     * @returns {Promise<Object>} 响应
     */
    async login(body) {
        const { email, password } = body;
        
        if (!email || !password) {
            return {
                status: 400,
                data: { error: 'Email and password required' }
            };
        }
        
        // 速率限制检查（每分钟最多5次）
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
            // 记录失败尝试
            this.rateLimits[`login_${email}`] = [...recentAttempts, now];
            
            return {
                status: 401,
                data: { error: 'Invalid email or password' }
            };
        }
        
        // 清除失败记录
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

    /**
     * GET /api/v1/user/profile
     * 获取当前用户信息
     * 
     * @param {string} token - Authorization Header中的Bearer Token
     * @returns {Promise<Object>} 响应
     */
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

    // ==================== B. 订阅与支付 API ====================

    /**
     * GET /api/v1/plans
     * 获取所有订阅方案
     * 
     * @returns {Promise<Object>} 响应
     */
    async getPlans() {
        const plans = JSON.parse(localStorage.getItem(STORAGE_KEYS.PLANS) || '[]');
        
        return {
            status: 200,
            data: plans
        };
    }

    /**
     * POST /api/v1/payment/create-order
     * 创建支付订单
     * 
     * @param {Object} body - 请求体 { plan_id, duration, payment_method }
     * @param {string} token - Authorization Token
     * @returns {Promise<Object>} 响应
     */
    async createOrder(body, token) {
        const user = verifyToken(token);
        
        if (!user) {
            return {
                status: 401,
                data: { error: 'Unauthorized' }
            };
        }
        
        const { plan_id, duration = 'month', payment_method = 'wechat' } = body;
        
        // 验证方案
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
        
        // 计算金额
        let amount = plan.price;
        let months = 1;
        
        if (duration === 'quarter') {
            amount = plan.price * 3 * 0.9; // 季度9折
            months = 3;
        } else if (duration === 'year') {
            amount = plan.price * 12 * 0.8; // 年付8折
            months = 12;
        }
        
        // 创建订单
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
            expires_at: new Date(Date.now() + 30 * 60 * 1000).toISOString() // 30分钟有效
        };
        
        const orders = JSON.parse(localStorage.getItem(STORAGE_KEYS.ORDERS) || '[]');
        orders.push(order);
        localStorage.setItem(STORAGE_KEYS.ORDERS, JSON.stringify(orders));
        
        // 模拟支付URL
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

    /**
     * POST /api/v1/payment/notify
     * 支付回调（模拟支付网关通知）
     * 
     * @param {Object} body - 支付网关通知
     * @returns {Promise<Object>} 响应
     */
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
        
        // 检查订单是否已处理
        if (order.status !== 'pending') {
            return {
                status: 400,
                data: { error: 'Order already processed' }
            };
        }
        
        if (status === 'success') {
            // 事务操作开始
            
            // 1. 更新订单状态
            order.status = 'paid';
            order.transaction_id = transaction_id || `TXN${Date.now()}`;
            order.paid_at = new Date().toISOString();
            orders[orderIndex] = order;
            localStorage.setItem(STORAGE_KEYS.ORDERS, JSON.stringify(orders));
            
            // 2. 更新用户订阅
            const users = JSON.parse(localStorage.getItem(STORAGE_KEYS.USERS) || '[]');
            const userIndex = users.findIndex(u => u.user_id === order.user_id);
            
            if (userIndex !== -1) {
                const user = users[userIndex];
                const currentExpiry = user.plan_expiry_date ? new Date(user.plan_expiry_date) : new Date();
                const baseDate = currentExpiry > new Date() ? currentExpiry : new Date();
                
                // 如果原到期时间未过期，则在原基础上顺延；否则从当前时间开始
                const newExpiry = new Date(baseDate);
                newExpiry.setMonth(newExpiry.getMonth() + order.months);
                
                user.current_plan_id = order.plan_id;
                user.plan_expiry_date = newExpiry.toISOString();
                users[userIndex] = user;
                localStorage.setItem(STORAGE_KEYS.USERS, JSON.stringify(users));
            }
            
            // 事务操作结束
            
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

    /**
     * GET /api/v1/payment/order/:order_id
     * 查询订单状态
     * 
     * @param {string} orderId - 订单ID
     * @param {string} token - Authorization Token
     * @returns {Promise<Object>} 响应
     */
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
        
        // 检查是否超时
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

    // ==================== 功能访问控制 ====================

    /**
     * 检查用户是否有权限访问付费功能
     * 
     * @param {string} token - Authorization Token
     * @param {string} feature - 功能标识
     * @returns {Promise<Object>} 响应
     */
    async checkAccess(token, feature) {
        const user = verifyToken(token);
        
        if (!user) {
            return {
                status: 401,
                data: { error: 'Unauthorized', has_access: false }
            };
        }
        
        // 检查订阅是否有效
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
        
        // 检查功能权限
        const plans = JSON.parse(localStorage.getItem(STORAGE_KEYS.PLANS) || '[]');
        const currentPlan = plans.find(p => p.plan_id === user.current_plan_id);
        
        const featureAccess = {
            'sensitivity_analysis': [2, 3],  // 专业版、企业版
            'financing_report': [2, 3],
            'api_access': [3],  // 仅企业版
            'multi_user': [3],
            'basic_calculation': [1, 2, 3]  // 所有版本
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

// ==================== API 客户端 ====================

/**
 * API客户端类 - 封装API调用
 */
class APIClient {
    constructor() {
        this.mock = new APIMock();
        this.baseUrl = '/api/v1';
    }

    /**
     * 获取存储的Token
     * @returns {string|null}
     */
    getToken() {
        return sessionStorage.getItem('auth_token');
    }

    /**
     * 设置Token
     * @param {string} token
     */
    setToken(token) {
        sessionStorage.setItem('auth_token', token);
    }

    /**
     * 清除Token
     */
    clearToken() {
        sessionStorage.removeItem('auth_token');
    }

    /**
     * 发送验证码
     * @param {string} email
     * @returns {Promise<Object>}
     */
    async sendCode(email) {
        return this.mock.sendCode({ email });
    }

    /**
     * 用户注册
     * @param {Object} data - { email, password, code, username?, plan_id? }
     * @returns {Promise<Object>}
     */
    async register(data) {
        const result = await this.mock.register(data);
        if (result.status === 201) {
            this.setToken(result.data.token);
        }
        return result;
    }

    /**
     * 用户登录
     * @param {string} email
     * @param {string} password
     * @returns {Promise<Object>}
     */
    async login(email, password) {
        const result = await this.mock.login({ email, password });
        if (result.status === 200) {
            this.setToken(result.data.token);
        }
        return result;
    }

    /**
     * 退出登录
     */
    logout() {
        this.clearToken();
    }

    /**
     * 获取用户信息
     * @returns {Promise<Object>}
     */
    async getProfile() {
        return this.mock.getProfile(this.getToken());
    }

    /**
     * 获取订阅方案
     * @returns {Promise<Object>}
     */
    async getPlans() {
        return this.mock.getPlans();
    }

    /**
     * 创建订单
     * @param {number} planId
     * @param {string} duration - 'month' | 'quarter' | 'year'
     * @param {string} paymentMethod - 'wechat' | 'alipay'
     * @returns {Promise<Object>}
     */
    async createOrder(planId, duration, paymentMethod) {
        return this.mock.createOrder(
            { plan_id: planId, duration, payment_method: paymentMethod },
            this.getToken()
        );
    }

    /**
     * 模拟支付成功
     * @param {string} orderId
     * @returns {Promise<Object>}
     */
    async simulatePayment(orderId) {
        return this.mock.paymentNotify({
            order_id: orderId,
            status: 'success'
        });
    }

    /**
     * 查询订单
     * @param {string} orderId
     * @returns {Promise<Object>}
     */
    async getOrder(orderId) {
        return this.mock.getOrder(orderId, this.getToken());
    }

    /**
     * 检查功能访问权限
     * @param {string} feature
     * @returns {Promise<Object>}
     */
    async checkAccess(feature) {
        return this.mock.checkAccess(this.getToken(), feature);
    }

    /**
     * 检查是否已登录
     * @returns {boolean}
     */
    isLoggedIn() {
        const token = this.getToken();
        return token && verifyToken(token) !== null;
    }
}

// ==================== 导出 ====================

// 创建全局API客户端实例
const api = new APIClient();

// 如果在Node.js环境
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { APIMock, APIClient, api };
}

// 如果在浏览器环境，挂载到window
if (typeof window !== 'undefined') {
    window.APIMock = APIMock;
    window.APIClient = APIClient;
    window.api = api;
}

console.log('[API Mock] Service initialized. Use `api` object to make API calls.');
console.log('[API Mock] Example: api.sendCode("user@example.com")');
