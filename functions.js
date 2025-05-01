Office.onReady(() => {
  checkAssignmentStatus();
});

// 📝 HTML-Entities dekodieren
function decodeHTMLEntities(html) {
  const txt = document.createElement('textarea');
  txt.innerHTML = html;
  return txt.value;
}

// 📝 HTML zu Plain Text mit LF-Zeilenumbrüchen (LF statt CRLF)
function htmlToPlainTextWithLineBreaks(text) {
  // <br> wird zu LF
  let result = text.replace(/<br\s*\/?>/gi, '\n');
  // Blockende-Tags werden zu LF
  result = result.replace(/<\/(p|div|ul|ol|li|table|tr|h[1-6])>/gi, '\n');
  // Alle übrigen Tags entfernen
  result = result.replace(/<[^>]+>/g, '');
  return result;
}

// 📏 Vorverarbeitung analog Java: HTML dekodieren, in Plain Text wandeln, whitespace-normalisieren, trim auf 4900 Zeichen
function preprocessBody(htmlContent) {
  // HTML-Entities dekodieren
  let decoded = decodeHTMLEntities(htmlContent);
  // Non-breaking spaces ersetzen
  decoded = decoded.replace(/\u00A0/g, ' ');
  // In Plain Text mit LF-Zeilenumbrüchen
  let plain = htmlToPlainTextWithLineBreaks(decoded);
  // Auf max. 4900 Zeichen kürzen (Java substring-Logik)
  if (plain.length > 4900) {
    plain = plain.substring(0, 4900);
  }
  console.log('processed length:', plain.length);
  console.log('processed content:', JSON.stringify(plain));
  return plain;
}

// 📛 SHA256 Hash-Funktion (UTF-8)
function sha256(str) {
  const encoder = new TextEncoder();
  return crypto.subtle.digest('SHA-256', encoder.encode(str))
    .then(hash => Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join(''));
}

// 🕒 Datumformat für MySQL: YYYY-MM-DD HH:MM:SS (lokale Zeit)
function formatDateTimeForMySQL(date) {
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ` +
         `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}`;
}

// 🔍 Hozzárendelési státusz lekérdezése
async function checkAssignmentStatus() {
  const item = Office.context.mailbox.item;
  try {
    const result = await new Promise(resolve => item.body.getAsync('html', resolve));
    if (result.status !== Office.AsyncResultStatus.Succeeded) throw new Error(result.error.message);

    const processed = preprocessBody(result.value);
    const bodyHash = await sha256(processed);
    console.log('Computed SHA256:', bodyHash);

    const payload = {
      subject: item.subject,
      receivedDateTime: formatDateTimeForMySQL(item.dateTimeCreated),
      from_address: item.from.emailAddress,
      to_address: item.to.length > 0 ? item.to[0].emailAddress : '',
      bodyHash
    };

    const response = await fetch('https://bogir.hu/V2/api/emails/emails_assignment_check.php', {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + btoa('Admin_2024$$:S3cure+P@ssw0rd2024!'),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
    if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);

    const data = await response.json();
    const statusDiv = document.getElementById('status');
    const btn = document.querySelector('button');
    if (data.status === 'assigned') {
      statusDiv.innerText = `Hozzárendelve: ${data.felelos}`;
      btn.disabled = true;
    } else if (data.status === 'unassigned') {
      statusDiv.innerText = 'Még nincs hozzárendelve.';
      btn.disabled = false;
    } else {
      statusDiv.innerText = 'E-mail nincs rögzítve az adatbázisban.';
      btn.disabled = true;
    }
  } catch (err) {
    console.error('checkAssignmentStatus error:', err.name, err.message);
    document.getElementById('status').innerText = `Hiba a betöltés során: ${err.message}`;
  }
}

// 🛡️ Hozzárendelés funkció
async function assignEmail() {
  const item = Office.context.mailbox.item;
  const statusDiv = document.getElementById('status');
  const btn = document.querySelector('button');
  btn.disabled = true;
  statusDiv.innerText = 'Hozzárendelés folyamatban...';

  try {
    const result = await new Promise(resolve => item.body.getAsync('html', resolve));
    if (result.status !== Office.AsyncResultStatus.Succeeded) throw new Error(result.error.message);

    const processed = preprocessBody(result.value);
    const bodyHash = await sha256(processed);
    console.log('Computed SHA256 (assign):', bodyHash);

    const payload = {
      subject: item.subject,
      receivedDateTime: formatDateTimeForMySQL(item.dateTimeCreated),
      from_address: item.from.emailAddress,
      to_address: item.to.length > 0 ? item.to[0].emailAddress : '',
      bodyHash,
      assignee: Office.context.mailbox.userProfile.emailAddress
    };

    const response = await fetch('https://bogir.hu/V2/api/emails/emails_assignment.php', {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + btoa('Admin_2024$$:S3cure+P@ssw0rd2024!'),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
    if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);

    const data = await response.json();
    if (data.status === 'success') {
      statusDiv.innerText = `Sikeres hozzárendelés: ${data.felelos}`;
    } else {
      throw new Error(data.message || 'Ismeretlen hiba');
    }
  } catch (err) {
    console.error('assignEmail error:', err.name, err.message);
    statusDiv.innerText = `Hiba a hozzárendelés során: ${err.message}`;
    btn.disabled = false;
  }
}
