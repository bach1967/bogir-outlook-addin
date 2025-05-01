Office.onReady(() => {
  checkAssignmentStatus();
});

// 📝 HTML zu Plain Text mit Zeilenumbrüchen
function htmlToPlainTextWithLineBreaks(html) {
  let text = html.replace(/<br\s*\/?/gi, '\n');
  text = text.replace(/<\/(p|div|ul|ol|li|table|tr|h[1-6])>/gi, '\n');
  text = text.replace(/<[^>]+>/g, '');
  return text;
}

// 📏 Vorverarbeitung analog Java: HTML->Plain, trim auf 4900 Zeichen
function preprocessBody(htmlContent) {
  const plain = htmlToPlainTextWithLineBreaks(htmlContent);
  return plain.length > 4900 ? plain.substring(0, 4900) : plain;
}

// 🕒 Datumformat für MySQL: YYYY-MM-DD HH:MM:SS (lokale Zeit)
function formatDateTimeForMySQL(date) {
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
}

// 🔍 Hozzárendelési státusz lekérdezése
async function checkAssignmentStatus() {
  const item = Office.context.mailbox.item;
  try {
    const result = await new Promise((resolve) => item.body.getAsync("html", resolve));
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      throw new Error(result.error.message);
    }
    const htmlContent = result.value;
    const processed = preprocessBody(htmlContent);
    console.log('Processed body for hash:', processed);

    const bodyHash = await sha256(processed);
    console.log('Computed SHA256:', bodyHash);

    const receivedLocal = formatDateTimeForMySQL(item.dateTimeCreated);
    const url = "https://bogir.hu/V2/api/emails/emails_assignment_check.php";
    const payload = {
      subject: item.subject,
      receivedDateTime: receivedLocal,
      from_address: item.from.emailAddress,
      to_address: item.to.length > 0 ? item.to[0].emailAddress : "",
      bodyHash: bodyHash
    };
    const options = {
      method: "POST",
      headers: {
        "Authorization": "Basic " + btoa("Admin_2024$$:S3cure+P@ssw0rd2024!"),
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    };
    // 🔧 Debug: URL und Optionen vor fetch
    console.log('Fetching URL:', url);
    console.log('Fetch options:', options);

    const response = await fetch(url, options);
    console.log('Fetch response:', response);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status} ${response.statusText}`);
    }
    const data = await response.json();
    console.log('Response JSON:', data);

    const statusDiv = document.getElementById("status");
    const assignBtn = document.querySelector("button");
    if (data.status === "assigned") {
      statusDiv.innerText = `Hozzárendelve: ${data.felelos}`;
      assignBtn.disabled = true;
    } else if (data.status === "unassigned") {
      statusDiv.innerText = "Még nincs hozzárendelve.";
      assignBtn.disabled = false;
    } else {
      statusDiv.innerText = "E-mail nincs rögzítve az adatbázisban.";
      assignBtn.disabled = true;
    }
  } catch (err) {
    console.error('checkAssignmentStatus error:', err);
    document.getElementById("status").innerText = `Hiba a betöltés során: ${err.message}`;
  }
}

// 📛 SHA256 Hash-Funktion (UTF-8)
function sha256(str) {
  const encoder = new TextEncoder();
  return crypto.subtle.digest('SHA-256', encoder.encode(str))
    .then(hash => Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join(''));
}

// 🛡️ Hozzárendelés funkció mit gleicher Vorverarbeitung
async function assignEmail() {
  const item = Office.context.mailbox.item;
  const statusDiv = document.getElementById("status");
  const assignBtn = document.querySelector("button");
  assignBtn.disabled = true;
  statusDiv.innerText = "Hozzárendelés folyamatban...";
  try {
    const result = await new Promise((resolve) => item.body.getAsync("html", resolve));
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      throw new Error(result.error.message);
    }
    const processed = preprocessBody(result.value);
    console.log('Processed body for assign:', processed);

    const bodyHash = await sha256(processed);
    console.log('Computed SHA256 for assign:', bodyHash);

    const receivedLocal = formatDateTimeForMySQL(item.dateTimeCreated);
    const url = "https://bogir.hu/V2/api/emails/emails_assignment.php";
    const payload = {
      subject: item.subject,
      receivedDateTime: receivedLocal,
      from_address: item.from.emailAddress,
      to_address: item.to.length > 0 ? item.to[0].emailAddress : "",
      bodyHash: bodyHash,
      assignee: Office.context.mailbox.userProfile.emailAddress
    };
    const options = {
      method: "POST",
      headers: {
        "Authorization": "Basic " + btoa("Admin_2024$$:S3cure+P@ssw0rd2024!"),
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    };
    console.log('Assign fetch URL:', url);
    console.log('Assign fetch options:', options);

    const response = await fetch(url, options);
    console.log('Assign response:', response);
    if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);
    const data = await response.json();
    console.log('Assign response JSON:', data);
    if (data.status === "success") {
      statusDiv.innerText = `Sikeres hozzárendelés: ${data.felelos}`;
    } else {
      throw new Error(data.message || 'Ismeretlen hiba');
    }
  } catch (err) {
    console.error('assignEmail error:', err);
    statusDiv.innerText = `Hiba a hozzárendelés során: ${err.message}`;
    assignBtn.disabled = false;
  }
}
