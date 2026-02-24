(function() {
  'use strict';

  // === 로그인/로그아웃 버튼 (메인 페이지) ===
  var loginBtn = document.getElementById('loginBtn');
  if (loginBtn) {
    loginBtn.addEventListener('click', async function() {
      loginBtn.disabled = true;
      loginBtn.textContent = '브라우저 열는 중...';
      try {
        await fetch('/api/auth/login', { method: 'POST' });
        loginBtn.textContent = '로그인 대기 중...';
        var pollTimer = setInterval(async function() {
          try {
            var r = await fetch('/api/auth/status');
            var d = await r.json();
            if (d.logged_in) {
              clearInterval(pollTimer);
              location.reload();
            }
          } catch(e) {}
        }, 2000);
        setTimeout(function() {
          clearInterval(pollTimer);
          loginBtn.disabled = false;
          loginBtn.textContent = 'ChatGPT로 로그인';
        }, 130000);
      } catch(e) {
        loginBtn.disabled = false;
        loginBtn.textContent = 'ChatGPT로 로그인';
      }
    });
  }

  var logoutBtn = document.getElementById('logoutBtn');
  if (logoutBtn) {
    logoutBtn.addEventListener('click', async function() {
      await fetch('/api/auth/logout', { method: 'POST' });
      location.reload();
    });
  }

  // === 파일 선택 버튼 ===
  var filePickBtn = document.getElementById('filePickBtn');
  var refFilesInput = document.getElementById('refFiles');
  var fileNamesEl = document.getElementById('fileNames');
  if (filePickBtn && refFilesInput) {
    filePickBtn.addEventListener('click', function() { refFilesInput.click(); });
    refFilesInput.addEventListener('change', function() {
      var names = Array.from(refFilesInput.files).map(function(f) { return f.name; });
      fileNamesEl.textContent = names.length ? names.join(', ') : '선택된 파일 없음';
    });
  }

  // === 메인 페이지: 전자책 생성 폼 ===
  var generateForm = document.getElementById('generateForm');
  if (generateForm) {
    generateForm.addEventListener('submit', async function(e) {
      e.preventDefault();

      var topic = document.getElementById('topicInput').value.trim();
      if (!topic) { alert('주제를 입력해주세요.'); return; }

      var includeImages = document.getElementById('includeImages').checked;

      // 선택된 출력 형식 수집
      var formatCheckboxes = document.querySelectorAll('input[name="outputFormat"]:checked');
      var outputFormats = Array.from(formatCheckboxes).map(function(cb) { return cb.value; });
      if (outputFormats.length === 0) outputFormats = ['pdf'];

      // FormData로 파일 + 텍스트 데이터 전송
      var fd = new FormData();
      fd.append('topic', topic);
      fd.append('include_images', includeImages ? '1' : '0');
      fd.append('output_formats', JSON.stringify(outputFormats));

      // 첨부 파일
      var refFilesEl = document.getElementById('refFiles');
      if (refFilesEl && refFilesEl.files.length > 0) {
        Array.from(refFilesEl.files).forEach(function(f) { fd.append('ref_files', f); });
      }

      // 참고 링크
      var refLinksEl = document.getElementById('refLinks');
      if (refLinksEl) fd.append('ref_links', refLinksEl.value.trim());

      var overlay = document.getElementById('progressOverlay');
      if (overlay) overlay.style.display = 'flex';

      try {
        var res = await fetch('/api/generate', {
          method: 'POST',
          body: fd,
        });
        var data = await res.json();

        if (!data.success) {
          if (overlay) overlay.style.display = 'none';
          alert('생성 실패: ' + (data.error || '알 수 없는 오류'));
          return;
        }

        pollProgress(data.task_id);

      } catch (err) {
        if (overlay) overlay.style.display = 'none';
        alert('오류: ' + err.message);
      }
    });
  }

  // === 테스트 모드 버튼 ===
  var testBtn = document.getElementById('testGenerateBtn');
  if (testBtn) {
    testBtn.addEventListener('click', async function() {
      testBtn.disabled = true;
      testBtn.textContent = '⏳ 파일 생성 중...';
      var overlay = document.getElementById('progressOverlay');
      var msgEl = document.getElementById('progressMessage');
      if (overlay) overlay.style.display = 'flex';
      if (msgEl) msgEl.textContent = '더미 데이터로 PDF·DOCX·PPTX·HWP 생성 중...';

      try {
        var res = await fetch('/api/test_generate', { method: 'POST' });
        var data = await res.json();
        if (data.success) {
          window.location.href = '/result/' + data.task_id;
        } else {
          if (overlay) overlay.style.display = 'none';
          alert('테스트 생성 실패: ' + (data.error || '알 수 없는 오류'));
          testBtn.disabled = false;
          testBtn.textContent = '🧪 테스트 모드 (AI 없이 즉시 생성)';
        }
      } catch(e) {
        if (overlay) overlay.style.display = 'none';
        alert('오류: ' + e.message);
        testBtn.disabled = false;
        testBtn.textContent = '🧪 테스트 모드 (AI 없이 즉시 생성)';
      }
    });
  }

  function pollProgress(taskId) {
    var timer = setInterval(async function() {
      try {
        var res = await fetch('/api/progress/' + taskId);
        var json = await res.json();
        if (!json.success) return;

        var d = json.data;
        var msgEl  = document.getElementById('progressMessage');
        var barEl  = document.getElementById('progressBar');
        var detailEl = document.getElementById('progressDetail');

        if (msgEl) msgEl.textContent = d.message || '';
        if (d.total_steps > 0 && barEl) {
          var pct = Math.round((d.step / d.total_steps) * 100);
          barEl.style.width = pct + '%';
        }
        if (detailEl) detailEl.textContent = d.step + ' / ' + d.total_steps + ' 단계';

        if (d.status === 'completed') {
          clearInterval(timer);
          window.location.href = '/result/' + taskId;
        } else if (d.status === 'error') {
          clearInterval(timer);
          var overlay = document.getElementById('progressOverlay');
          if (overlay) overlay.style.display = 'none';
          alert('생성 실패: ' + d.message);
        }
      } catch (e) { /* 네트워크 오류 무시 */ }
    }, 2000);
  }

})();
