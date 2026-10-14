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
        <button type="button" class="btn btn-outline-custom position-relative" onclick='window.open("https://mail.21cn.net/", "_blank")'>
          Webmail
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-envelope" viewBox="0 0 16 16">
            <path d="M0 4a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2zm2-1a1 1 0 0 0-1 1v.217l7 4.2 7-4.2V4a1 1 0 0 0-1-1zm13 2.383-4.708 2.825L15 11.105zm-.034 6.876-5.64-3.471L8 9.583l-1.326-.795-5.64 3.47A1 1 0 0 0 2 13h12a1 1 0 0 0 .966-.741M1 11.105l4.708-2.897L1 5.383z"/>
          </svg>
        </button>

        <button type="button" class="btn btn-outline-custom" data-bs-toggle="modal" data-bs-target="#staticBackdrop">Server List⛏️</button>
        <button type="button" class="btn btn-outline-custom" data-page="analyzer.html">Troubleshooting</button>
        <button type="button" class="btn btn-outline-custom" data-page="DBaaS_Summary.html">GCIS DBaaS Summary</button>
        <button type="button" class="btn btn-outline-custom" data-page="Converter.html">Unit Converter⚡</button>
   
        <!-- Modal -->
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
                        <li class="nav-item"><a href="#profile" class="nav-link" data-bs-toggle="tab">2</a></li>
                        <li class="nav-item"><a href="#stake" class="nav-link" data-bs-toggle="tab">3</a></li>
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
                            <a href="" class="list-group-item list-group-item-action list-group-item-success" target="_blank">Empty</a>
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

  // ===== Supabase + Session =====
  const SUPABASE_URL = "https://fyrzooerrrmkwrbjapoa.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_YvzQbyBUN4ex7WLgLV5HFA_n2ctfmu0";
  const SESSION_TIMEOUT = 60 * 60 * 1000; // 1 小時

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
  }

  // 啟動
  loadSupabase(initAuth);

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", injectNavbar);
  } else {
    injectNavbar();
  }
})();