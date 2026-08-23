const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "ShopAgent-X Team";

// Color palette - Teal Trust (电商感)
const C = {
  teal: "028090",
  seafoam: "00A896",
  mint: "02C39A",
  white: "FFFFFF",
  dark: "1B2A3A",
  light: "F0F7F6",
  gray: "6B7B8D",
  accent: "F39C12",
  red: "E74C3C",
  green: "27AE60"
};

// ===== Slide 1: Cover =====
{
  const s = pptx.addSlide();
  s.background = { color: C.teal };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.mint } });
  s.addText("ShopAgent-X", { x: 0.8, y: 1.5, w: "85%", h: 1.2, fontSize: 54, fontFace: "Arial Black", color: C.white, bold: true });
  s.addText("AI 智能导购系统 · 项目答辩", { x: 0.8, y: 2.7, w: "85%", h: 0.6, fontSize: 22, fontFace: "Calibri", color: C.mint });
  s.addShape(pptx.ShapeType.rect, { x: 0.8, y: 3.4, w: 3, h: 0.04, fill: { color: C.accent } });
  s.addText("姚昆鹏  ·  Gaogxx", { x: 0.8, y: 3.7, w: 5, h: 0.5, fontSize: 16, fontFace: "Calibri", color: C.seafoam });
  s.addText("三端 AI 电商导购平台", { x: 0.8, y: 4.2, w: 6, h: 0.4, fontSize: 13, fontFace: "Calibri", color: C.gray });
}

// ===== Slide 2: 项目背景与目标 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.teal } });
  s.addText("项目背景与目标", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.dark });

  s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.3, w: 6, h: 3.2, fill: { color: C.light }, rectRadius: 0.12 });
  s.addText("🎯 项目目标", { x: 0.6, y: 1.4, w: 5, h: 0.5, fontSize: 18, fontFace: "Arial Black", color: C.teal });
  s.addText([
    { text: "打造一款", options: { fontSize: 14 } },
    { text: " AI 驱动的电商智能导购平台\n\n", options: { bold: true, fontSize: 14 } },
    { text: "• 用户自然语言提问 → AI 理解意图\n", options: { fontSize: 13 } },
    { text: "• 智能检索商品库 + 知识库\n", options: { fontSize: 13 } },
    { text: "• 生成个性化推荐 + 对比分析\n", options: { fontSize: 13 } },
    { text: "• 购物车管理 + 订单全流程\n", options: { fontSize: 13 } },
    { text: "• 管理后台数据大屏 + 知识巡检\n", options: { fontSize: 13 } },
    { text: "• AI 对话式购物体验\n", options: { fontSize: 13 } }
  ], { x: 0.8, y: 2.0, w: 5.4, h: 2.4, fontFace: "Calibri", color: C.dark, lineSpacingMultiple: 1.35 });

  s.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: 1.3, w: 5.6, h: 3.2, fill: { color: C.teal }, rectRadius: 0.12 });
  s.addText("💡 解决什么痛点", { x: 7.0, y: 1.4, w: 5, h: 0.5, fontSize: 18, fontFace: "Arial Black", color: C.white });
  s.addText([
    { text: "❌ 传统电商搜索体验差", options: { fontSize: 12, color: C.seafoam } },
    { text: "\n   关键词搜索 → 无法理解用户真实需求\n\n", options: { fontSize: 11, color: C.white } },
    { text: "❌ 商品对比决策困难", options: { fontSize: 12, color: C.seafoam } },
    { text: "\n   多个商品参数难以横向比较\n\n", options: { fontSize: 11, color: C.white } },
    { text: "❌ 知识库管理低效", options: { fontSize: 12, color: C.seafoam } },
    { text: "\n   无法识别知识缺口和质量问题\n\n", options: { fontSize: 11, color: C.white } },
    { text: "→ AI导购 + 知识库闭环，解决全部痛点", options: { fontSize: 13, color: C.accent, bold: true } }
  ], { x: 7.2, y: 2.0, w: 5.0, h: 2.4, fontFace: "Calibri", lineSpacingMultiple: 1.2 });
}

// ===== Slide 3: 系统架构 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.teal } });
  s.addText("系统架构", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.dark });

  // Frontend layer
  s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.2, w: 12, h: 0.9, fill: { color: C.light }, rectRadius: 0.08 });
  s.addText("📱 前端层", { x: 0.6, y: 1.3, w: 2, h: 0.7, fontSize: 14, fontFace: "Arial Black", color: C.teal });
  s.addText("Android App (Kotlin Compose)    +    Web 管理后台 (React + Vite)    +    普通用户 Web 端", { x: 2.5, y: 1.3, w: 9.5, h: 0.7, fontSize: 12, fontFace: "Calibri", color: C.gray });

  // Backend layer
  s.addText("↓  HTTP / SSE 流式  ↓", { x: 0.4, y: 2.15, w: 12, h: 0.4, fontSize: 11, color: C.gray, align: "center" });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 2.5, w: 12, h: 1.0, fill: { color: C.teal }, rectRadius: 0.08 });
  s.addText("☕ 后端层 (Java Spring Boot 3.2  :8888)", { x: 0.6, y: 2.6, w: 8, h: 0.35, fontSize: 14, fontFace: "Arial Black", color: C.white });
  s.addText("21 个 Controller · JWT 鉴权 · MyBatis-Plus ORM · RESTful API · SSE 流式推送", { x: 0.6, y: 3.0, w: 11, h: 0.35, fontSize: 11, fontFace: "Calibri", color: C.seafoam });

  // AI layer
  s.addText("↓  HTTP /ask  ↓", { x: 0.4, y: 3.55, w: 12, h: 0.4, fontSize: 11, color: C.gray, align: "center" });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 3.85, w: 12, h: 1.1, fill: { color: C.dark }, rectRadius: 0.08 });
  s.addText("🐍 AI 服务层 (Python FastAPI  :8000)", { x: 0.6, y: 3.95, w: 8, h: 0.35, fontSize: 14, fontFace: "Arial Black", color: C.mint });
  s.addText("RouterAgent 意图分类 → ShoppingAgent / KnowledgeQAAgent / AdminCopilotAgent / ChitChatAgent | LLM: 通义千问 qwen-plus | 向量库: FAISS | 文档解析: 语义分块", { x: 0.6, y: 4.35, w: 11.5, h: 0.45, fontSize: 10, fontFace: "Calibri", color: C.seafoam });

  // Data layer
  s.addText("↓  ↓  ↓", { x: 0.4, y: 5.0, w: 12, h: 0.35, fontSize: 11, color: C.gray, align: "center" });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 5.3, w: 5.5, h: 0.65, fill: { color: C.light }, rectRadius: 0.06 });
  s.addText("🗄️ MySQL 8.4 (shop_agent_db)", { x: 0.6, y: 5.4, w: 5, h: 0.45, fontSize: 12, fontFace: "Arial Black", color: C.teal });
  s.addShape(pptx.ShapeType.roundRect, { x: 6.9, y: 5.3, w: 5.5, h: 0.65, fill: { color: C.light }, rectRadius: 0.06 });
  s.addText("⚡ Redis 缓存", { x: 7.1, y: 5.4, w: 5, h: 0.45, fontSize: 12, fontFace: "Arial Black", color: C.teal });
}

// ===== Slide 4: AI 核心链路 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.teal } });
  s.addText("AI 核心链路详解", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.dark });

  // Flow diagram
  const steps = [
    { label: "用户输入", desc: "自然语言\n问题", color: C.teal },
    { label: "意图分类", desc: "LLM+关键词\nRouterAgent", color: C.seafoam },
    { label: "任务分发", desc: "Shopping\nKnowledgeQA\nChitChat\nAdminCopilot", color: C.mint },
    { label: "知识检索", desc: "FAISS向量库\n语义搜索", color: C.accent },
    { label: "LLM生成", desc: "通义千问\nqwen-plus", color: C.gray },
    { label: "流式返回", desc: "SSE推送\n逐字输出", color: C.dark }
  ];
  steps.forEach((st, i) => {
    const x = 0.2 + i * 2.2;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.3, w: 2.0, h: 1.6, fill: { color: st.color }, rectRadius: 0.1 });
    s.addText(st.label, { x, y: 1.4, w: 2.0, h: 0.5, fontSize: 14, fontFace: "Arial Black", color: C.white, align: "center" });
    s.addText(st.desc, { x, y: 1.9, w: 2.0, h: 0.9, fontSize: 10, fontFace: "Calibri", color: C.white, align: "center" });
    if (i < 5) {
      s.addText("→", { x: x + 2.0, y: 1.7, w: 0.2, h: 0.5, fontSize: 18, color: C.gray, align: "center" });
    }
  });

  // Agent details
  s.addText("核心 Agent 矩阵", { x: 0.6, y: 3.2, w: 8, h: 0.5, fontSize: 17, fontFace: "Arial Black", color: C.dark });

  const agents = [
    { name: "ShoppingAgent", desc: "商品推荐 · 多维度对比 · 个性化导购 · 购物车管理", icon: "🛒" },
    { name: "KnowledgeQAAgent", desc: "L1 直接检索 → L2 问题改写 → L3 多步推理（Orchestrator 编排）", icon: "📚" },
    { name: "ChitChatAgent", desc: "闲聊对话 · 情感识别 · 身份查询", icon: "💬" },
    { name: "AdminCopilotAgent", desc: "管理知识库 · 7项实时数据采集 · LLM 智能运营分析", icon: "⚙️" },
    { name: "InspectionAgent", desc: "重复文档检测 · 低质量Chunk · 过期知识 · 无人访问文档", icon: "🔍" }
  ];
  agents.forEach((a, i) => {
    const y = 3.8 + i * 0.48;
    s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y, w: 12, h: 0.42, fill: { color: i % 2 === 0 ? C.light : C.white }, rectRadius: 0.05 });
    s.addText(a.icon + "  " + a.name, { x: 0.6, y, w: 3.5, h: 0.42, fontSize: 12, fontFace: "Arial Black", color: C.teal });
    s.addText(a.desc, { x: 4.2, y, w: 8, h: 0.42, fontSize: 10, fontFace: "Calibri", color: C.gray });
  });
}

// ===== Slide 5: 前端展示 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.teal } });
  s.addText("前端展示 — 三端界面", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.dark });

  // Android
  s.addShape(pptx.ShapeType.roundRect, { x: 0.3, y: 1.2, w: 3.8, h: 4.2, fill: { color: C.light }, rectRadius: 0.12 });
  s.addText("📱 Android App", { x: 0.5, y: 1.3, w: 3.4, h: 0.5, fontSize: 16, fontFace: "Arial Black", color: C.teal, align: "center" });
  s.addText("Kotlin Compose 原生开发\n\n• AI 对话购物界面\n• 商品卡片流式展示\n• 图片识别推荐\n• 语音输入识别\n• 购物车 & 订单管理\n• 拍照识图导购", { x: 0.5, y: 1.9, w: 3.4, h: 3.2, fontSize: 11, fontFace: "Calibri", color: C.gray, align: "center", lineSpacingMultiple: 1.4 });

  // Web User
  s.addShape(pptx.ShapeType.roundRect, { x: 4.5, y: 1.2, w: 3.8, h: 4.2, fill: { color: C.light }, rectRadius: 0.12 });
  s.addText("🌐 Web 用户端", { x: 4.7, y: 1.3, w: 3.4, h: 0.5, fontSize: 16, fontFace: "Arial Black", color: C.teal, align: "center" });
  s.addText("React + Vite\n\n• 聊天式购物对话\n• 知识库浏览搜索\n• 会话管理\n• 商品浏览 & 搜索\n• 个人中心", { x: 4.7, y: 1.9, w: 3.4, h: 3.2, fontSize: 11, fontFace: "Calibri", color: C.gray, align: "center", lineSpacingMultiple: 1.4 });

  // Admin
  s.addShape(pptx.ShapeType.roundRect, { x: 8.7, y: 1.2, w: 3.8, h: 4.2, fill: { color: C.teal }, rectRadius: 0.12 });
  s.addText("⚙️ 管理后台", { x: 8.9, y: 1.3, w: 3.4, h: 0.5, fontSize: 16, fontFace: "Arial Black", color: C.white, align: "center" });
  s.addText("React + Vite\n\n• 数据大屏仪表盘\n• 商品/订单管理\n• 知识库管理\n• 管理助手 AI 对话\n• 知识巡检\n• Agent 执行监控\n• 用户/对话管理", { x: 8.9, y: 1.9, w: 3.4, h: 3.2, fontSize: 11, fontFace: "Calibri", color: C.seafoam, align: "center", lineSpacingMultiple: 1.4 });
}

// ===== Slide 6: 管理后台核心功能 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.teal } });
  s.addText("管理后台核心功能", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.dark });

  const features = [
    { title: "📊 数据大屏", desc: "4 核心指标卡片\n意图分布饼图(10色)\n7日趋势渐变折线图\n热门商品TOP5排名\n最近推荐记录\n未命中问题追踪", color: C.teal },
    { title: "⚙️ 管理助手", desc: "AI 对话式管理\n完整对话 CRUD\n历史记录持久化\n内置管理知识库\n实时数据库分析\n智能运营建议", color: C.seafoam },
    { title: "🔍 知识巡检", desc: "未命中问题聚类\nJaccard 相似度算法\n自动补库建议\n重复文档检测\n低质量 Chunk 扫描\n过期知识预警", color: C.mint },
    { title: "📋 数据管理", desc: "商品 CRUD 管理\n订单全流程追踪\n用户/对话管理\n知识库文档上传\n文档解析状态\nAgent 执行记录", color: C.accent }
  ];
  features.forEach((f, i) => {
    const x = 0.3 + i * 3.2;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.2, w: 3.0, h: 4.0, fill: { color: C.light }, rectRadius: 0.12 });
    s.addShape(pptx.ShapeType.roundRect, { x: x + 0.1, y: 1.3, w: 2.8, h: 0.55, fill: { color: f.color }, rectRadius: 0.08 });
    s.addText(f.title, { x: x + 0.1, y: 1.3, w: 2.8, h: 0.55, fontSize: 14, fontFace: "Arial Black", color: C.white, align: "center" });
    s.addText(f.desc, { x: x + 0.3, y: 2.1, w: 2.4, h: 2.9, fontSize: 11, fontFace: "Calibri", color: C.gray, lineSpacingMultiple: 1.55 });
  });
}

// ===== Slide 7: 数据与知识库体系 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.teal } });
  s.addText("数据与知识库体系", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.dark });

  // Database schema
  s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.2, w: 6.2, h: 2.1, fill: { color: C.light }, rectRadius: 0.1 });
  s.addText("🗄️ 数据库设计 (MySQL 8.4)", { x: 0.6, y: 1.3, w: 5, h: 0.45, fontSize: 15, fontFace: "Arial Black", color: C.teal });
  s.addText([
    { text: "用户体系：", options: { bold: true, fontSize: 10 } },
    { text: "user, admin\n", options: { fontSize: 10 } },
    { text: "电商体系：", options: { bold: true, fontSize: 10 } },
    { text: "product, category, order, cart, address\n", options: { fontSize: 10 } },
    { text: "对话体系：", options: { bold: true, fontSize: 10 } },
    { text: "conversation, message\n", options: { fontSize: 10 } },
    { text: "知识体系：", options: { bold: true, fontSize: 10 } },
    { text: "knowledge_doc, knowledge_chunk\n", options: { fontSize: 10 } },
    { text: "AI 体系：", options: { bold: true, fontSize: 10 } },
    { text: "agent_run, qa_log, qa_unanswered\n", options: { fontSize: 10 } },
    { text: "运营体系：", options: { bold: true, fontSize: 10 } },
    { text: "recommendation_log, notice", options: { fontSize: 10 } }
  ], { x: 0.8, y: 1.85, w: 5.6, h: 1.3, fontFace: "Calibri", color: C.gray, lineSpacingMultiple: 1.3 });

  // Knowledge flow
  s.addShape(pptx.ShapeType.roundRect, { x: 7.0, y: 1.2, w: 5.5, h: 2.1, fill: { color: C.teal }, rectRadius: 0.1 });
  s.addText("📚 知识库工作流", { x: 7.2, y: 1.3, w: 5, h: 0.45, fontSize: 15, fontFace: "Arial Black", color: C.white });
  s.addText("文档上传 → DocumentParser 语义分块 → FAISS 向量化 → 用户提问 → RetrievalAgent 检索 → LLM 生成回答 → 未命中记录 → 知识巡检 → 补库建议", { x: 7.2, y: 1.85, w: 5.1, h: 1.3, fontSize: 10, fontFace: "Calibri", color: C.seafoam, lineSpacingMultiple: 1.5 });

  // Stats
  s.addText("核心数据指标", { x: 0.6, y: 3.5, w: 8, h: 0.5, fontSize: 17, fontFace: "Arial Black", color: C.dark });

  const metrics = [
    { num: "21", label: "Java Controller" },
    { num: "10+", label: "AI Agent" },
    { num: "20+", label: "数据库表" },
    { num: "7", label: "意图类型" },
    { num: "3", label: "前端终端" }
  ];
  metrics.forEach((m, i) => {
    const x = 0.3 + i * 2.5;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 4.1, w: 2.2, h: 1.4, fill: { color: C.light }, rectRadius: 0.1 });
    s.addText(m.num, { x, y: 4.2, w: 2.2, h: 0.8, fontSize: 32, fontFace: "Arial Black", color: C.teal, align: "center" });
    s.addText(m.label, { x, y: 4.95, w: 2.2, h: 0.4, fontSize: 12, fontFace: "Calibri", color: C.gray, align: "center" });
  });
}

// ===== Slide 8: 技术亮点 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.teal };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.mint } });
  s.addText("技术亮点 & 创新点", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.white });

  const highlights = [
    { title: "多 Agent 协作架构", desc: "RouterAgent 中央路由 + 5个专业Agent分工协作。LLM意图分类 + 关键词fallback双路径，Orchestrator五层编排支持L3复杂推理。流式SSE推送，逐字输出提升体验。" },
    { title: "RAG 知识库闭环", desc: "文档上传 → 语义分块 → 向量检索 → LLM生成。未命中问题自动记录到 qa_unanswered 表，知识巡检定期检测重复/低质量/过期文档，形成知识库持续优化闭环。" },
    { title: "三端统一架构", desc: "Android(Kotlin Compose) + Web用户端(React) + 管理后台(React)，共享同一套Java后端和Python AI服务。JWT统一鉴权，scene字段区分用户/管理对话。" },
    { title: "真实数据驱动", desc: "管理助手基于 MySQL 实时采集 7项指标（文档/Chunk/问答/用户/未命中TOP5/最近文档），结合 LLM 生成专业运营分析报告，不是死板的模板输出。" }
  ];
  highlights.forEach((h, i) => {
    const y = 1.2 + i * 1.15;
    s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y, w: 12, h: 1.0, fill: { color: "1A3A4A" }, rectRadius: 0.08 });
    s.addText((i + 1).toString(), { x: 0.6, y: y + 0.1, w: 0.5, h: 0.5, fontSize: 22, fontFace: "Arial Black", color: C.accent, align: "center" });
    s.addText(h.title, { x: 1.3, y: y + 0.05, w: 4, h: 0.4, fontSize: 14, fontFace: "Arial Black", color: C.white });
    s.addText(h.desc, { x: 1.3, y: y + 0.45, w: 10.8, h: 0.5, fontSize: 10, fontFace: "Calibri", color: C.seafoam, lineSpacingMultiple: 1.3 });
  });
}

// ===== Slide 9: 分工与协作 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.teal } });
  s.addText("团队分工与协作", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.dark });

  // EvanYao826
  s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.2, w: 5.8, h: 4.0, fill: { color: C.light }, rectRadius: 0.12 });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.6, y: 1.3, w: 2.2, h: 0.45, fill: { color: C.teal }, rectRadius: 0.08 });
  s.addText("姚昆鹏", { x: 0.6, y: 1.3, w: 2.2, h: 0.45, fontSize: 13, fontFace: "Arial Black", color: C.white, align: "center" });
  s.addText("项目 Owner · 全栈主力开发", { x: 3.0, y: 1.3, w: 3, h: 0.45, fontSize: 11, fontFace: "Calibri", color: C.gray });
  s.addText([
    { text: "• 三端架构设计搭建\n", options: { fontSize: 11 } },
    { text: "• Java 全部基础 Controller (12个)\n", options: { fontSize: 11 } },
    { text: "• Python AI 核心链路\n", options: { fontSize: 11 } },
    { text: "   RouterAgent / ShoppingAgent\n", options: { fontSize: 11 } },
    { text: "   KnowledgeQAAgent (L1/L2/L3)\n", options: { fontSize: 11 } },
    { text: "   Orchestrator 编排框架\n", options: { fontSize: 11 } },
    { text: "• Android App 全量 (Kotlin Compose)\n", options: { fontSize: 11 } },
    { text: "• 前端框架 + 管理后台 UI 美化\n", options: { fontSize: 11 } },
    { text: "• 数据库 init.sql + Docker 部署\n", options: { fontSize: 11 } },
    { text: "• 设计文档 + 部署指南\n", options: { fontSize: 11 } },
    { text: "• Agent 执行监控功能", options: { fontSize: 11 } }
  ], { x: 0.6, y: 1.9, w: 5.4, h: 3.1, fontFace: "Calibri", color: C.dark, lineSpacingMultiple: 1.35 });

  // Gaogxx
  s.addShape(pptx.ShapeType.roundRect, { x: 6.6, y: 1.2, w: 5.8, h: 4.0, fill: { color: C.light }, rectRadius: 0.12 });
  s.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: 1.3, w: 2.2, h: 0.45, fill: { color: C.mint }, rectRadius: 0.08 });
  s.addText("Gaogxx", { x: 6.8, y: 1.3, w: 2.2, h: 0.45, fontSize: 13, fontFace: "Arial Black", color: C.white, align: "center" });
  s.addText("后端 + AI + 前端功能开发", { x: 9.2, y: 1.3, w: 3, h: 0.45, fontSize: 11, fontFace: "Calibri", color: C.gray });
  s.addText([
    { text: "• AdminChatController 完整对话系统\n", options: { fontSize: 11 } },
    { text: "• KnowledgeInspection 知识巡检后端\n", options: { fontSize: 11 } },
    { text: "• AdminCopilotAgent 重写 + 知识库\n", options: { fontSize: 11 } },
    { text: "• IntentClassifier 管理员路由修复\n", options: { fontSize: 11 } },
    { text: "• InspectionAgent Schema 修复\n", options: { fontSize: 11 } },
    { text: "• 仪表盘数据大屏重设计\n", options: { fontSize: 11 } },
    { text: "• 管理助手对话 UI 全功能\n", options: { fontSize: 11 } },
    { text: "• 知识巡检页面 (双 Tab)\n", options: { fontSize: 11 } },
    { text: "• UTF-8 全链路编码修复\n", options: { fontSize: 11 } },
    { text: "• 数据库列补充 + 种子数据\n", options: { fontSize: 11 } },
    { text: "• AI Agent 协作开发模式", options: { fontSize: 11 } }
  ], { x: 6.8, y: 1.9, w: 5.4, h: 3.1, fontFace: "Calibri", color: C.dark, lineSpacingMultiple: 1.35 });

  s.addText("协作模式：main ← dev(受保护) ← feat/* 功能分支  |  PR Review → Merge  |  仓库: github.com/EvanYao826/ShopAgent-X", { x: 0.6, y: 5.4, w: 12, h: 0.35, fontSize: 11, fontFace: "Calibri", color: C.gray, align: "center" });
}

// ===== Slide 10: 总结与展望 =====
{
  const s = pptx.addSlide();
  s.background = { color: C.teal };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: C.mint } });
  s.addText("总结与展望", { x: 0.6, y: 0.3, w: 10, h: 0.7, fontSize: 30, fontFace: "Arial Black", color: C.white });

  // Completed
  s.addText("✅ 已完成", { x: 0.6, y: 1.2, w: 5, h: 0.45, fontSize: 18, fontFace: "Arial Black", color: C.accent });
  const done = [
    "三端 AI 导购平台完整搭建",
    "多 Agent 协作 + SSE 流式推送",
    "RAG 知识库 + 向量检索 + 闭环优化",
    "管理后台 9 大功能模块",
    "Android App 全功能",
    "20+ 数据库表 + Docker 部署"
  ];
  done.forEach((d, i) => {
    s.addText("• " + d, { x: 0.8, y: 1.7 + i * 0.35, w: 5.5, h: 0.3, fontSize: 11, fontFace: "Calibri", color: C.seafoam });
  });

  // Future
  s.addText("🚀 未来展望", { x: 7.0, y: 1.2, w: 5, h: 0.45, fontSize: 18, fontFace: "Arial Black", color: C.accent });
  const future = [
    "接入更多 LLM（Claude/GPT-4）",
    "多模态：图片/语音输入导购",
    "用户画像 + 个性化推荐增强",
    "A/B 测试框架 + 效果评估",
    "商品知识图谱构建",
    "生产环境部署 + 性能优化"
  ];
  future.forEach((f, i) => {
    s.addText("• " + f, { x: 7.2, y: 1.7 + i * 0.35, w: 5.5, h: 0.3, fontSize: 11, fontFace: "Calibri", color: C.seafoam });
  });

  // Stats
  const stats = [
    { num: "110+", label: "Git Commits" },
    { num: "50+", label: "源文件" },
    { num: "3", label: "终端" },
    { num: "5", label: "AI Agent" }
  ];
  stats.forEach((st, i) => {
    const x = 0.6 + i * 3.2;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 4.0, w: 2.8, h: 1.3, fill: { color: "1A3A4A" }, rectRadius: 0.1 });
    s.addText(st.num, { x, y: 4.05, w: 2.8, h: 0.7, fontSize: 32, fontFace: "Arial Black", color: C.accent, align: "center" });
    s.addText(st.label, { x, y: 4.75, w: 2.8, h: 0.35, fontSize: 12, fontFace: "Calibri", color: C.seafoam, align: "center" });
  });

  s.addText("谢谢！欢迎提问", { x: 0.6, y: 5.5, w: 12, h: 0.5, fontSize: 18, fontFace: "Calibri", color: C.gray, align: "center" });
}

// Save
const outPath = "C:/Users/21124/Desktop/ShopAgent-X_答辩.pptx";
pptx.writeFile({ fileName: outPath }).then(() => console.log("Saved to: " + outPath));
