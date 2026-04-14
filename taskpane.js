const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyPI4RfVgUMYgWE0xYTU8eEoBoKc7h3_C18DBv9M6_0VSGLNocYZsFGXlJZf3WiwsqQ/exec';

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    trackEmailOpen();
  }
});

function trackEmailOpen() {
  const item = Office.context.mailbox.item;

  const recipient = Office.context.mailbox.userProfile.emailAddress || '';
  const sender    = item.from ? item.from.emailAddress : '';
  const subject   = item.subject || '(no subject)';

  setStatus('loading', 'Logging open event...', '');

  fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ recipient, sender, subject })
  })
  .then(function(res) { return res.json(); })
  .then(function(data) {
    if (data.status === 'success') {
      setStatus('success', 'Open logged', sender + ' — ' + subject);
    } else if (data.status === 'duplicate') {
      setStatus('duplicate', 'Already logged', 'Unique open — not duplicated');
    } else {
      setStatus('error', 'Error: ' + data.message, '');
    }
  })
  .catch(function(err) {
    setStatus('error', 'Network error', err.toString());
  });
}

function setStatus(type, msg, detail) {
  const dot    = document.getElementById('statusDot');
  const status = document.getElementById('statusMsg');
  const detailEl = document.getElementById('detailMsg');

  dot.className = 'dot ' + type;
  status.textContent = msg;
  detailEl.textContent = detail;
}