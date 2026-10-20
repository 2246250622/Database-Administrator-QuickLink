(function () {
  const navbarHTML = `
<header class="navbar-premium">
  <div class="nav-container">
    <div class="nav-left">
      <a href="home.html" class="nav-brand">
        <span class="brand-dot"></span>
        <span data-i18n="brand">DBA Portal</span>
      </a>
    </div>

    <div class="nav-center">
      <a href="home.html" class="nav-link" data-page="home.html" data-i18n="nav.dashboard">Dashboard</a>
      <a href="tickets.html" class="nav-link" data-page="tickets.html" data-i18n="nav.tickets">Tickets</a>
      <a href="warns.html" class="nav-link" data-page="warns.html" data-i18n="nav.warns">Warns</a>
      <a href="dbaas_list.html" class="nav-link" data-page="dbaas_list.html" data-i18n="nav.servers">Servers</a>
      <a href="inventory.html" class="nav-link" data-page="inventory.html" data-i18n="nav.inventory">Inventory</a>
    </div>

    <div class="nav-right">
      <button class="nav-btn" onclick='window.open("https://mail.21cn.net/", "_blank")' data-i18n="nav.webmail">Webmail</button>

      <select id="themeSelect" class="nav-select" title="Theme">
        <option value="system">System</option>
        <option value="light">Light</option>
        <option value="dark">Dark</option>
      </select>

      <select id="langSelect" class="nav-select" title="Language">
        <option value="en">ENG</option>
        <option value="zh-TW">繁</option>
        <option value="zh-CN">簡</option>
      </select>

      <div class="dropdown">
        <button class="nav-avatar" data-bs-toggle="dropdown">
          <img src="IMG_3755.jpg" alt="user" width="32" height="32">
        </button>
        <ul class="dropdown-menu dropdown-menu-end">
          <li><a class="dropdown-item" href="Admin.html" data-i18n="nav.admin">Admin</a></li>
          <li><hr class="dropdown-divider"></li>
          <li><a class="dropdown-item" href="#" onclick="logout(); return false;" data-i18n="nav.signout">Sign out</a></li>
        </ul>
      </div>
    </div>
  </div>
</header>
  `;

  const I18N = {
    en: {
      brand: "DBA Portal",
      "nav.dashboard": "Dashboard",
      "nav.tickets": "Tickets",
      "nav.warns": "Warns",
      "nav.servers": "Servers",
      "nav.inventory": "Inventory",
      "nav.webmail": "Webmail",
      "nav.admin": "Admin",
      "nav.signout": "Sign out",
      "dash.title": "Dashboard",
      "dash.subtitle": "Today's overview",
      "dash.open": "Open Incidents",
      "dash.critical": "Critical (Today)",
      "dash.warn": "Warnings (Today)",
      "dash.latest": "Latest Incidents",
      "dash.empty": "No incidents today",
      "tickets.title": "INC Tickets",
      "tickets.refresh": "Refresh",
      "tickets.blackouts": "Blackouts",
      "tickets.search": "Search INC, host, message...",
      "tickets.allStatus": "All status",
      "tickets.empty": "No matching tickets",
      "warns.title": "WARN Alerts",
      "warns.search": "Search INC, host, message...",
      "warns.empty": "No matching alerts",
      "servers.title": "Server List",
      "servers.subtitle": "Auto-loaded from GCIS DBaaS Summary2.xls",
      "servers.search": "Search VM / IP...",
      "inv.title": "Inventory",
      "inv.subtitle": "Auto-loaded from Inventory.xlsx",
      "inv.search": "Search VM / IP / Purpose...",
      "admin.title": "User Management",
      "admin.denied": "Only the owner can access this page."
    },
    "zh-TW": {
      brand: "DBA 入口",
      "nav.dashboard": "儀表板",
      "nav.tickets": "事件單",
      "nav.warns": "警告",
      "nav.servers": "伺服器",
      "nav.inventory": "資產清單",
      "nav.webmail": "網頁郵件",
      "nav.admin": "管理",
      "nav.signout": "登出",
      "dash.title": "儀表板",
      "dash.subtitle": "今日概覽",
      "dash.open": "未結事件",
      "dash.critical": "嚴重（今日）",
      "dash.warn": "警告（今日）",
      "dash.latest": "最新事件",
      "dash.empty": "今日無事件",
      "tickets.title": "事件單",
      "tickets.refresh": "重新載入",
      "tickets.blackouts": "屏蔽",
      "tickets.search": "搜尋事件 / 主機 / 訊息...",
      "tickets.allStatus": "全部狀態",
      "tickets.empty": "沒有符合的資料",
      "warns.title": "警告",
      "warns.search": "搜尋事件 / 主機 / 訊息...",
      "warns.empty": "沒有符合的資料",
      "servers.title": "伺服器清單",
      "servers.subtitle": "自動載入 GCIS DBaaS Summary2.xls",
      "servers.search": "搜尋 VM / IP...",
      "inv.title": "資產清單",
      "inv.subtitle": "自動載入 Inventory.xlsx",
      "inv.search": "搜尋 VM / IP / 用途...",
      "admin.title": "使用者管理",
      "admin.denied": "僅限擁有者可存取此頁面。"
    },
    "zh-CN": {
      brand: "DBA 入口",
      "nav.dashboard": "仪表板",
      "nav.tickets": "事件单",
      "nav.warns": "警告",
      "nav.servers": "服务器",
      "nav.inventory": "资产清单",
      "nav.webmail": "网页邮件",
      "nav.admin": "管理",
      "nav.signout": "登出",
      "dash.title": "仪表板",
      "dash.subtitle": "今日概览",
      "dash.open": "未结事件",
      "dash.critical": "严重（今日）",
      "dash.warn": "警告（今日）",
      "dash.latest": "最新事件",
      "dash.empty": "今日无事件",
      "tickets.title": "事件单",
      "tickets.refresh": "重新加载",
      "tickets.blackouts": "屏蔽",
      "tickets.search": "搜索事件 / 主机 / 消息...",
      "tickets.allStatus": "全部状态",
      "tickets.empty": "没有符合的资料",
      "warns.title": "警告",
      "warns.search": "搜索事件 / 主机 / 消息...",
      "warns.empty": "没有符合的资料",
      "servers.title": "服务器清单",
      "servers.subtitle": "自动加载 GCIS DBaaS Summary2.xls",
      "servers.search": "搜索 VM / IP...",
      "inv.title": "资产清单",
      "inv.subtitle": "自动加载 Inventory.xlsx",
      "inv.search": "搜索 VM / IP / 用途...",
      "admin.title": "用户管理",
      "admin.denied": "仅限所有者可访问此页面。"
    }
  };

  function getLang() {
    return localStorage.getItem("lang") || "en";
  }

  function applyLang(lang) {
    const dict = I18N[lang] || I18N.en;
    document.querySelectorAll("[data-i18n]").forEach(el => {
      const key = el.getAttribute("data-i18n");
      if (dict[key]) el.textContent = dict[key];
    });
    document.querySelectorAll("[data-i18n-placeholder]").forEach(el => {
      const key = el.getAttribute("data-i18n-placeholder");
      if (dict[key]) el.placeholder = dict[key];
    });
    document.documentElement.lang = lang === "zh-CN" ? "zh-CN" : lang === "zh-TW" ? "zh-HK" : "en";
    localStorage.setItem("lang", lang);
  }

  function getTheme() {
    return localStorage.getItem("theme") || "system";
  }

  function applyTheme(theme) {
    const root = document.documentElement;
    let effective = theme;
    if (theme === "system") {
      effective = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
    }
    root.setAttribute("data-theme", effective);
    localStorage.setItem("theme", theme);
  }

  function loadSupabase(cb) {
    if (window.supabase) return cb();
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2";
    s.onload = cb;
    document.head.appendChild(s);
  }

  function initAuth() {
    const SUPABASE_URL = "https://fyrzooerrrmkwrbjapoa.supabase.co";
    const SUPABASE_ANON_KEY = "sb_publishable_YvzQbyBUN4ex7WLgLV5HFA_n2ctfmu0";
    if (!window.supabaseClient) {
      window.supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    }
    const last = parseInt(localStorage.getItem("lastActivity") || "0", 10);
    if (Date.now() - last > 3600000) {
      logout();
      return;
    }
    localStorage.setItem("lastActivity", Date.now().toString());
    document.addEventListener("click", () => localStorage.setItem("lastActivity", Date.now().toString()));
  }

  window.logout = async function () {
    if (window.supabaseClient) await window.supabaseClient.auth.signOut();
    localStorage.removeItem("isLoggedIn");
    localStorage.removeItem("lastActivity");
    location.href = "index.html";
  };

  function injectNavbar() {
    const placeholder = document.getElementById("navbar-placeholder");
    if (!placeholder) return;
    placeholder.outerHTML = navbarHTML;

    const current = location.pathname.split("/").pop() || "home.html";
    document.querySelectorAll(".nav-link").forEach(link => {
      if (link.getAttribute("data-page") === current) link.classList.add("active");
    });

    const themeSelect = document.getElementById("themeSelect");
    const langSelect = document.getElementById("langSelect");
    if (themeSelect) {
      themeSelect.value = getTheme();
      themeSelect.addEventListener("change", () => applyTheme(themeSelect.value));
    }
    if (langSelect) {
      langSelect.value = getLang();
      langSelect.addEventListener("change", () => applyLang(langSelect.value));
    }

    applyTheme(getTheme());
    applyLang(getLang());

    window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => {
      if (getTheme() === "system") applyTheme("system");
    });
  }

  loadSupabase(initAuth);

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", injectNavbar);
  } else {
    injectNavbar();
  }
})();