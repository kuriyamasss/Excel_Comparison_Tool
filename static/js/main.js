// static/js/main.js (覆盖现有文件)
//
// 功能：
// - 在服务器端渲染的多步骤页面中实现真正的“上一步”导航（POST 回到 /prepare_fields）
// - reset 按钮回到根路径（/），避免前端仅 hide 导致空白
// - 检测浏览器刷新（reload），自动重置到根路径
// - 保持原有下载 / 关闭 / overlay 行为

function showOverlay(){ const ov = document.getElementById('overlay'); if(ov) ov.style.display='flex'; ov && ov.setAttribute('aria-hidden','false'); }
function hideOverlay(){ const ov = document.getElementById('overlay'); if(ov) ov.style.display='none'; ov && ov.setAttribute('aria-hidden','true'); }

// Helper: create and submit a POST form (used to go back to server-rendered step)
function postToUrl(path, params) {
  const form = document.createElement('form');
  form.method = 'post';
  form.action = path;
  for (const key in params) {
    const input = document.createElement('input');
    input.type = 'hidden';
    input.name = key;
    input.value = params[key] === undefined || params[key] === null ? '' : params[key];
    form.appendChild(input);
  }
  document.body.appendChild(form);
  form.submit();
}

document.addEventListener('DOMContentLoaded', function(){
  // If this page load is the result of a browser reload, redirect to root to reset state.
  try {
    const navEntries = performance.getEntriesByType && performance.getEntriesByType('navigation');
    if (navEntries && navEntries.length > 0) {
      if (navEntries[0].type === 'reload') {
        // force reset to root
        window.location.replace('/');
        return;
      }
    } else if (performance.navigation) {
      // fallback for older browsers
      if (performance.navigation.type === performance.navigation.TYPE_RELOAD) {
        window.location.replace('/');
        return;
      }
    }
  } catch (e) {
    // ignore if API unavailable
  }

  // Show overlay on any form submit (upload, prepare_fields, compare)
  document.querySelectorAll('form').forEach(function(f){
    f.addEventListener('submit', function(e){
      showOverlay();
    });
  });

  // STEP: "上一步" in step2 -> go to root (step1)
  const step2Back = document.getElementById('back1') || document.getElementById('step2-back');
  if (step2Back) {
    step2Back.addEventListener('click', function(e){
      e.preventDefault();
      // Simply navigate to root so server renders step1 fresh
      window.location.href = '/';
    });
  }

  // STEP: "上一步" in step3 -> submit old_id/new_id to /prepare_fields to render step2
  const step3Back = document.getElementById('back2') || document.getElementById('step3-back');
  if (step3Back) {
    step3Back.addEventListener('click', function(e){
      e.preventDefault();
      // collect needed hidden inputs present on the page
      // if inputs not found, fallback to root
      const old_id = (document.querySelector('input[name=old_id]') || {}).value;
      const new_id = (document.querySelector('input[name=new_id]') || {}).value;
      // preserve sheet selections and header settings if present
      const sheet_old = (document.querySelector('input[name=sheet_old]') || {}).value;
      const sheet_new = (document.querySelector('input[name=sheet_new]') || {}).value;
      const header_mode = (document.querySelector('input[name=header_mode]') || {}).value;
      const header_row_index = (document.querySelector('input[name=header_row_index]') || {}).value;

      if (old_id && new_id) {
        showOverlay();
        // Post back to /prepare_fields with old_id/new_id and previous selections
        postToUrl('/prepare_fields', {
          old_id: old_id,
          new_id: new_id,
          sheet_old: sheet_old,
          sheet_new: sheet_new,
          header_mode: header_mode,
          header_row_index: header_row_index
        });
      } else {
        // fallback: reload root
        window.location.href = '/';
      }
    });
  }

  // language selection (cookie + reload)
  const sel = document.getElementById('lang');
  if (sel) {
    sel.addEventListener('change', function(){
      const val = sel.value;
      document.cookie = 'lang=' + val + '; path=/; max-age=' + (365*24*60*60);
      const u = new URL(window.location.href);
      u.searchParams.set('lang', val);
      window.location.href = u.toString();
    });
  }

  // download via fetch and finally hide overlay
  const dl = document.getElementById('downloadBtn');
  if (dl) {
    dl.addEventListener('click', async function(e){
      e.preventDefault();
      const url = dl.getAttribute('data-url') || dl.getAttribute('data-href') || dl.dataset.url;
      const name = dl.getAttribute('data-name') || dl.dataset.name || 'compare_result.xlsx';
      try {
        showOverlay();
        const res = await fetch(url, { method: 'GET' });
        if (!res.ok) throw new Error('Network response was not ok');
        const blob = await res.blob();
        const blobUrl = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(blobUrl);
      } catch (err) {
        alert('下载失败：' + err.message);
      } finally {
        // After download, we want a fresh start: navigate to root so server renders step1
        hideOverlay();
      }
    });
  }

  // reset button: navigate to root to ensure server returns to initial state
  const resetBtn = document.getElementById('resetBtn');
  if (resetBtn) {
    resetBtn.addEventListener('click', function(e){
      e.preventDefault();
      // navigate to root; the server will render the initial upload page
      window.location.href = '/';
    });
  }

  // close program: call backend /shutdown then close window (as before)
  const closeBtn = document.getElementById('closeBtn');
  if (closeBtn) {
    closeBtn.addEventListener('click', async function(e){
      e.preventDefault();
      showOverlay();
      try {
        await fetch('/shutdown', { method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' } });
      } catch (err) {
        // ignore
      } finally {
        // try to close tab/window; if blocked, navigate to blank
        try { window.close(); } catch (e) {}
        try { window.location.href = 'about:blank'; } catch (e) {}
      }
    });
  }

});
