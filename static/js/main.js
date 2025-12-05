// main.js: 负责前端交互（overlay, download, reset, language cookie）
function showOverlay(){ const ov = document.getElementById('overlay'); if(ov) ov.style.display='flex'; ov && ov.setAttribute('aria-hidden','false'); }
function hideOverlay(){ const ov = document.getElementById('overlay'); if(ov) ov.style.display='none'; ov && ov.setAttribute('aria-hidden','true'); }

document.addEventListener('DOMContentLoaded', function(){
  // 在任何表单提交时显示 overlay
  document.querySelectorAll('form').forEach(function(f){
    f.addEventListener('submit', function(e){
      showOverlay();
    });
  });

  // 语言选择：写 cookie 并重载页面（把 lang 参数附加到 URL）
  const sel = document.getElementById('lang');
  if(sel){
    sel.addEventListener('change', function(){
      const val = sel.value;
      document.cookie = 'lang=' + val + '; path=/; max-age=' + (365*24*60*60);
      const u = new URL(window.location.href);
      u.searchParams.set('lang', val);
      window.location.href = u.toString();
    });
  }

  // 下载逻辑：使用 fetch 获取 blob，完成后隐藏 overlay
  const dl = document.getElementById('downloadBtn');
  if(dl){
    dl.addEventListener('click', async function(e){
      e.preventDefault();
      const url = dl.getAttribute('data-url');
      const name = dl.getAttribute('data-name') || 'compare_result.xlsx';
      try{
        showOverlay();
        const res = await fetch(url, { method: 'GET' });
        if(!res.ok) throw new Error('Network response was not ok');
        const blob = await res.blob();
        const blobUrl = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(blobUrl);
      }catch(err){
        alert('下载失败：' + err.message);
      }finally{
        hideOverlay();
      }
    });
  }

  // reset button: 回首页，保留语言 cookie
  const resetBtn = document.getElementById('resetBtn');
  if(resetBtn){
    resetBtn.addEventListener('click', function(e){
      e.preventDefault();
      hideOverlay();
      window.location.href = '/';
    });
  }
});
