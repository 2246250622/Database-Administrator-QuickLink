(function () {
  const navbarHTML = `
<header class="p-3 mb-3 border-bottom" style="background-color: #D0A2F7">
  <div class="container">
    <div class="d-flex flex-wrap align-items-center justify-content-center justify-content-lg-start">
      <a href="/" class="d-flex align-items-center mb-2 mb-lg-0 link-body-emphasis text-decoration-none">
        <svg class="bi me-2" width="40" height="32" role="img" aria-label="Bootstrap"><use xlink:href="#bootstrap"/></svg>
      </a>

      <ul class="nav col-12 col-lg-auto me-lg-auto mb-2 justify-content-center mb-md-0">
        <li><b><a href="home.html" class="nav-link px-2 link-body-emphasis">Database Administrator</a></b></li>
      </ul>

      <form class="col-12 col-lg-auto mb-3 mb-lg-0 me-lg-3" role="search">
        <div class="d-flex flex-wrap align-items-center gap-2">
          <button type="button" class="btn btn-outline-custom position-relative" onclick='window.open("https://mail.21cn.net/", "_blank")'>
            Webmail
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-envelope" viewBox="0 0 16 16">
              <path d="M0 4a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2zm2-1a1 1 0 0 0-1 1v.217l7 4.2 7-4.2V4a1 1 0 0 0-1-1zm13 2.383-4.708 2.825L15 11.105zm-.034 6.876-5.64-3.471L8 9.583l-1.326-.795-5.64 3.47A1 1 0 0 0 2 13h12a1 1 0 0 0 .966-.741M1 11.105l4.708-2.897L1 5.383z"/>
            </svg>
          </button>

          <div class="dropdown">
            <button class="btn btn-outline-custom dropdown-toggle" type="button" data-bs-toggle="dropdown">Alerts</button>
            <ul class="dropdown-menu">
              <li><a class="dropdown-item" href="tickets.html">Tickets</a></li>
              <li><a class="dropdown-item" href="warns.html">Warns</a></li>
            </ul>
          </div>

          <div class="dropdown">
            <button class="btn btn-outline-custom dropdown-toggle" type="button" data-bs-toggle="dropdown">DBaaS</button>
            <ul class="dropdown-menu">
              <li><a class="dropdown-item" data-bs-toggle="modal" data-bs-target="#staticBackdrop">Server List⛏️</a></li>
              <li><a class="dropdown-item" href="DBaaS_Summary.html">GCIS DBaaS Summary</a></li>
              <li><a class="dropdown-item" href="bookmark.html">Bookmarks</a></li>
            </ul>
          </div>

          <div class="dropdown">
            <button class="btn btn-outline-custom dropdown-toggle" type="button" data-bs-toggle="dropdown">Tools</button>
            <ul class="dropdown-menu">
              <li><a class="dropdown-item" href="analyzer.html">Troubleshooting</a></li>
              <li><a class="dropdown-item" href="Converter.html">Unit Converter⚡</a></li>
            </ul>
          </div>
        </div>

        <div class="modal fade" id="staticBackdrop" data-bs-backdrop="static" data-bs-keyboard="false" tabindex="-1" aria-labelledby="staticBackdropLabel1" aria-hidden="true">
          <div class="modal-dialog modal-dialog-centered">
            <div class="modal-content">
              <div class="modal-header">
                <h1 class="modal-title fs-5" id="staticBackdropLabel1">⛏️ Server List</h1>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
              </div>
              <div class="modal-body">
                <div class="m-4">
                  <div class="card text-center">
                    <div class="card-header">
                      <ul class="nav nav-tabs card-header-tabs">
                        <li class="nav-item"><a href="#home" class="nav-link active" data-bs-toggle="tab">Server List</a></li>
                        <li class="nav-item"><a href="#profile" class="nav-link" data-bs-toggle="tab">Inventory</a></li>
                        <li class="nav-item"><a href="#stake" class="nav-link" data-bs-toggle="tab">CA InFo</a></li>
                      </ul>
                    </div>
                    <div class="card-body">
                      <div class="tab-content">
                        <div class="tab-pane fade show active" id="home">
                          <div class="list-group">
                            <a href="dbaas_list.html" class="list-group-item list-group-item-action list-group-item-danger" target="_blank">Summary</a>
                          </div>
                        </div>
                        <div class="tab-pane fade" id="profile">
                          <div class="list-group">
                            <a href="inventory.html" class="list-group-item list-group-item-action list-group-item-info" target="_blank">Inventory</a>
                          </div>
                        </div>
                        <div class="tab-pane fade" id="stake">
                          <div class="list-group">
                          <a href="ca_info.html" class="list-group-item list-group-item-action list-group-item-success" target="_blank">CA Info</a>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">好的</button>
                <button type="button" class="btn btn-primary" data-bs-dismiss="modal">關閉</button>
              </div>
            </div>
          </div>
        </div>
      </form>

      <div class="dropdown text-end">
        <a href="#" class="d-block link-body-emphasis text-decoration-none dropdown-toggle" data-bs-toggle="dropdown" aria-expanded="false">
          <img src="IMG_3755.jpg" alt="mdo" width="32" height="32" class="rounded-circle">
        </a>
        <ul class="dropdown-menu text-small">
          <li><a class="dropdown-item" href="Admin.html">Admin</a></li>
          <li><hr class="dropdown-divider"></li>
          <li><a class="dropdown-item" href="#" onclick="logout(); return false;">Sign out</a></li>
        </ul>
      </div>
    </div>
  </div>
</header>
  `;

  const SUPABASE_URL = "https://fyrzooerrrmkwrbjapoa.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_YvzQbyBUN4ex7WLgLV5HFA_n2ctfmu0";
  const SESSION_TIMEOUT = 60 * 60 * 1000;

  function loadSupabase(callback) {
    if (window.supabase) {
      callback();
      return;
    }
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2";
    script.onload = callback;
    document.head.appendChild(script);
  }

  function initAuth() {
    window.supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    checkAuth();
    setupActivityTracker();
  }

  async function checkAuth() {
    const path = location.pathname;
    if (path.endsWith("index.html") || path.endsWith("/")) return;

    const { data } = await window.supabaseClient.auth.getSession();
    if (!data.session) {
      localStorage.removeItem("isLoggedIn");
      localStorage.removeItem("lastActivity");
      alert("Please log in first!");
      location.href = "index.html";
      return;
    }
    localStorage.setItem("isLoggedIn", "true");
    localStorage.setItem("lastActivity", Date.now().toString());
  }

  function setupActivityTracker() {
    function updateActivity() {
      localStorage.setItem("lastActivity", Date.now().toString());
    }
    ["click", "keydown", "mousemove", "scroll"].forEach(evt => {
      document.addEventListener(evt, updateActivity, { passive: true });
    });

    setInterval(async () => {
      const last = localStorage.getItem("lastActivity");
      if (last && Date.now() - parseInt(last) > SESSION_TIMEOUT) {
        await window.supabaseClient.auth.signOut();
        localStorage.removeItem("isLoggedIn");
        localStorage.removeItem("lastActivity");
        alert("Session expired. Please log in again.");
        location.href = "index.html";
      }
    }, 60 * 1000);
  }

  window.logout = async function () {
    if (window.supabaseClient) {
      await window.supabaseClient.auth.signOut();
    }
    localStorage.removeItem("isLoggedIn");
    localStorage.removeItem("lastActivity");
    location.href = "index.html";
  };

  function injectChatWidget() {
    if (document.getElementById('chatWidget')) return;

    const widget = document.createElement('div');
    widget.innerHTML = `
      <div id="chatWidget" style="position:fixed; bottom:20px; right:20px; z-index:9999;">
        <button id="chatToggle" class="btn btn-primary rounded-circle shadow position-relative" style="width:56px; height:56px; font-size:24px;">
          💬
          <span id="unreadBadge" class="position-absolute top-0 start-100 translate-middle badge rounded-pill bg-danger" style="display:none; font-size:0.65rem;">0</span>
        </button>
        <div id="chatPanel" class="card shadow" style="display:none; position:absolute; bottom:70px; right:0; width:340px; height:450px;">
          <div class="card-header d-flex justify-content-between align-items-center py-2">
            <strong>Chatroom</strong>
            <button class="btn-close btn-sm" id="chatClose"></button>
          </div>
          <div class="card-body p-2 d-flex flex-column" style="height: calc(100% - 46px);">
            <div id="chatMessages" class="flex-grow-1 overflow-auto mb-2" style="font-size:0.85rem; background:#f8fafc; border-radius:6px; padding:8px;"></div>
            <form id="chatForm" class="d-flex gap-1">
              <input type="text" id="chatInput" class="form-control form-control-sm" placeholder="Message..." autocomplete="off">
              <button type="submit" class="btn btn-sm btn-primary">Submit</button>
            </form>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(widget);

    const toggle = document.getElementById('chatToggle');
    const panel = document.getElementById('chatPanel');
    const closeBtn = document.getElementById('chatClose');
    const form = document.getElementById('chatForm');
    const input = document.getElementById('chatInput');
    const messagesDiv = document.getElementById('chatMessages');
    const unreadBadge = document.getElementById('unreadBadge');

    let myEmail = 'Anonymous';
    let loaded = false;
    let unreadCount = 0;

    function updateBadge() {
      if (unreadCount > 0 && panel.style.display === 'none') {
        unreadBadge.textContent = unreadCount > 99 ? '99+' : unreadCount;
        unreadBadge.style.display = 'block';
      } else {
        unreadBadge.style.display = 'none';
      }
    }

    toggle.addEventListener('click', async () => {
      const isHidden = panel.style.display === 'none';
      panel.style.display = isHidden ? 'block' : 'none';
      if (isHidden) {
        unreadCount = 0;
        updateBadge();
        if (!loaded) {
          messagesDiv.innerHTML = '<div class="text-muted small">載入中...</div>';
          await initChat();
          loaded = true;
        }
      }
    });

    closeBtn.addEventListener('click', () => panel.style.display = 'none');

    // ===== 新 INC 通知 =====
    let lastTicketCount = 0;

    async function checkNewTickets() {
      if (!window.supabaseClient) return;
      if (Notification.permission === 'default') {
        Notification.requestPermission();
      }
      const { data } = await window.supabaseClient
        .from('tickets')
        .select('id')
        .eq('status', 'OPEN');
      const count = data?.length || 0;
      if (lastTicketCount > 0 && count > lastTicketCount) {
        if (Notification.permission === 'granted') {
          new Notification('新 CRITICAL INC', {
            body: `有 ${count - lastTicketCount} 筆新的 OPEN ticket`,
            icon: 'IMG_3755.jpg'
          });
        }
      }
      lastTicketCount = count;
    }

    setInterval(checkNewTickets, 1000); // 每 1 秒檢查一次
    setTimeout(checkNewTickets, 1000);

    async function initChat() {
      try {
        const { data } = await window.supabaseClient.auth.getSession();
        if (data?.session) myEmail = data.session.user.email;

        const { data: msgs, error } = await window.supabaseClient
          .from('messages')
          .select('*')
          .order('created_at', { ascending: true })
          .limit(80);

        messagesDiv.innerHTML = '';
        if (error) {
          messagesDiv.innerHTML = `<div class="text-danger small">Load Failed: ${error.message}</div>`;
          return;
        }
        if (!msgs || msgs.length === 0) {
          messagesDiv.innerHTML = '<div class="text-muted small">No messages yet</div>';
        } else {
          msgs.forEach(msg => renderMsg(msg, false));
        }

        window.supabaseClient
          .channel('chat-widget')
          .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'messages' }, payload => {
            renderMsg(payload.new, true);
          })
          .subscribe();
      } catch (e) {
        messagesDiv.innerHTML = `<div class="text-danger small">錯誤：${e.message}</div>`;
      }
    }

    function renderMsg(msg, isNew) {
      const isMine = msg.user_email === myEmail;
      const time = new Date(msg.created_at).toLocaleTimeString('zh-HK', { hour: '2-digit', minute: '2-digit' });
      const div = document.createElement('div');
      div.className = 'mb-2' + (isMine ? ' text-end' : '');
      div.innerHTML = `
        <div class="small text-muted">${(msg.user_email || '').split('@')[0]} · ${time}</div>
        <span class="badge ${isMine ? 'bg-primary' : 'bg-light text-dark border'}">${escapeHtml(msg.content)}</span>
      `;
      messagesDiv.appendChild(div);
      messagesDiv.scrollTop = messagesDiv.scrollHeight;

      if (isNew && panel.style.display === 'none' && !isMine) {
        unreadCount++;
        updateBadge();
      }
    }

    function escapeHtml(t) {
      const d = document.createElement('div');
      d.textContent = t || '';
      return d.innerHTML;
    }

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      const content = input.value.trim();
      if (!content) return;
      input.value = '';
      await window.supabaseClient.from('messages').insert({ user_email: myEmail, content });
    });
  }

  function injectNavbar() {
    const placeholder = document.getElementById("navbar-placeholder");
    if (!placeholder) return;

    placeholder.outerHTML = navbarHTML;

    const currentPage = location.pathname.split("/").pop() || "home.html";
    document.querySelectorAll("[data-page]").forEach(btn => {
      if (btn.getAttribute("data-page") === currentPage) {
        btn.classList.add("active");
      }
      btn.addEventListener("click", () => {
        location.href = btn.getAttribute("data-page");
      });
    });

    function initModal() {
      if (typeof bootstrap === "undefined") {
        setTimeout(initModal, 50);
        return;
      }
      const modalEl = document.getElementById("staticBackdrop");
      if (!modalEl) return;
      const modal = new bootstrap.Modal(modalEl);
      const serverListBtn = document.querySelector('[data-bs-target="#staticBackdrop"]');
      if (serverListBtn) {
        serverListBtn.addEventListener("click", function (e) {
          e.preventDefault();
          modal.show();
        });
      }
    }
    initModal();
    injectChatWidget();
  }

  loadSupabase(initAuth);

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", injectNavbar);
  } else {
    injectNavbar();
  }
})();