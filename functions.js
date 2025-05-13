async function ensureMasterCategory() {
  // 1) Hole dir ein Graph-Access-Token für Outlook:
  const tokenResponse = await new Promise(resolve =>
    Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true }, resolve)
  );
  if (tokenResponse.status !== Office.AsyncResultStatus.Succeeded) {
    throw new Error('Token fehlgeschlagen: ' + tokenResponse.error.message);
  }
  const token = tokenResponse.value;

  // 2) Prüfe vorhandene Kategorien
  let resp = await fetch('https://graph.microsoft.com/v1.0/me/outlook/masterCategories', {
    headers: { 'Authorization': `Bearer ${token}` }
  });
  let { value: cats } = await resp.json();
  const exists = cats.some(c => c.displayName === 'Hozzárendelve');

  if (!exists) {
    // 3) Falls nicht vorhanden, erstelle sie mit Grün
    await fetch('https://graph.microsoft.com/v1.0/me/outlook/masterCategories', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        displayName: 'Hozzárendelve',
        // Graph-Farben: presetColor enum; Grün ist "preset0" oder "green"
        color: 'green'
      })
    });
    console.log('Master-Kategorie "Hozzárendelve" angelegt.');
  }
}


Office.onReady(() => {
    ensureMasterCategory().catch(console.error);
  checkAssignmentStatus();
});

function safelyHideElement(el) {
  if (!el) return;

  // Fokus entfernen, wenn ein Kind fokussiert ist
  if (el.contains(document.activeElement)) {
    document.activeElement.blur();
  }

  // Jetzt ist es sicher, aria-hidden zu setzen
  el.setAttribute('aria-hidden', 'true');
}

function decodeHTMLEntities(html) {
  const txt = document.createElement('textarea');
  txt.innerHTML = html;
  return txt.value;
}

function htmlToPlainTextWithLineBreaks(text) {
  let result = text.replace(/<br\s*\/?>/gi, '\n');
  result = result.replace(/<\/(p|div|ul|ol|li|table|tr|h[1-6])>/gi, '\n');
  result = result.replace(/<[^>]+>/g, '');
  return result;
}



function preprocessBody(htmlContent) {
  let decoded = decodeHTMLEntities(htmlContent);
  decoded = decoded.replace(/\u00A0/g, ' ');
  let plain = htmlToPlainTextWithLineBreaks(decoded);
  if (plain.length > 4900) {
    plain = plain.substring(0, 4900);
  }
  console.log('processed length:', plain.length);
  console.log('processed content:', JSON.stringify(plain));
  return plain;
}

function formatDateTimeForMySQL(date) {
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ` +
         `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}`;
}

async function checkAssignmentStatus() {
  const item = Office.context.mailbox.item;
  try {
    const result = await new Promise(resolve => item.body.getAsync('html', resolve));
    if (result.status !== Office.AsyncResultStatus.Succeeded) throw new Error(result.error.message);

    const processed = preprocessBody(result.value);

    const payload = {
      subject: item.subject,
      receivedDateTime: formatDateTimeForMySQL(item.dateTimeCreated),
      from_address: item.from.emailAddress,
      to_address: item.to.length > 0 ? item.to[0].emailAddress : ''
    };

console.log('payload=',payload);

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
    if (data.status.status === 'assigned') {
      statusDiv.innerText = `Hozzárendelve: ${data.status.felelos}`;
      btn.disabled = true;


const CATEGORY_NAME = 'Hozzárendelve';

// Beim Zuweisen
item.categories.addAsync(
  [CATEGORY_NAME],
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(`Kategorie '${CATEGORY_NAME}' wurde gesetzt.`);
    } else {
      console.error('Fehler beim Setzen der Kategorie:', asyncResult.error);
    }
  }
);


    } else if (data.status.status === 'unassigned') {
      statusDiv.innerText = 'Még nincs hozzárendelve.';
      btn.disabled = false;
    } else {
      statusDiv.innerText = 'E-mail nincs rögzítve az adatbázisban.';
      btn.disabled = true;
    }
  } catch (err) {
    console.error('checkAssignmentStatus error:', err.name, err.message);
    const statusDiv = document.getElementById('status');
    statusDiv.innerText = `Hiba a betöltés során: ${err.message}`;
  
    const statusWrapper = document.getElementById('statusWrapper');
    safelyHideElement(statusWrapper); // Jetzt barrierefrei ausblenden
  }
}

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

    const payload = {
      subject: item.subject,
      receivedDateTime: formatDateTimeForMySQL(item.dateTimeCreated),
      from_address: item.from.emailAddress,
      to_address: item.to.length > 0 ? item.to[0].emailAddress : '',
      assignee: Office.context.mailbox.userProfile.emailAddress
    };

console.log('payload',payload);

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
    if (data.status && data.status.status === 'success') {
      statusDiv.innerText = `Sikeres hozzárendelés: ${data.status.felelos}`;
    } else {
      throw new Error(data.status?.message || 'Ismeretlen hiba');
    }
  } catch (err) {
    console.error('assignEmail error:', err.name, err.message);
    statusDiv.innerText = `Hiba a hozzárendelés során: ${err.message}`;
    btn.disabled = false;
  }
}
