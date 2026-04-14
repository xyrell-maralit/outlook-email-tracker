const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyPI4RfVgUMYgWE0xYTU8eEoBoKc7h3_C18DBv9M6_0VSGLNocYZsFGXlJZf3WiwsqQ/exec';

function sendToSheets(recipient, sender, subject) {
  fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ recipient, sender, subject })
  })
  .then(function() {
    setStatus('success', 'Open logged', sender + ' — ' + subject);
  })
  .catch(function(err) {
    setStatus('error', 'Network error', err.toString());
  });
}

function setStatus(type, msg, detail) {
  var dot = document.getElementById('statusDot');
  var status = document.getElementById('statusMsg');
  var detailEl = document.getElementById('detailMsg');
  if (dot) dot.className = 'dot ' + type;
  if (status) status.textContent = msg;
  if (detailEl) detailEl.textContent = detail;
}

if (typeof Office !== 'undefined') {
  Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
      var item = Office.context.mailbox.item;
      var recipient = Office.context.mailbox.userProfile.emailAddress || '';
      var sender = item.from ? item.from.emailAddress : '';
      var subject = item.subject || '(no subject)';
      setStatus('loading', 'Logging open event...', '');
      sendToSheets(recipient, sender, subject);
    } else {
      setStatus('error', 'Not in Outlook context', '');
    }
  });
} else {
  setStatus('error', 'Office.js not loaded', 'Open this inside Outlook');
}
