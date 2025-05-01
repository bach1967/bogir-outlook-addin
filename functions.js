Office.onReady(() => {
  checkAssignmentStatus();
});

// 📝 HTML zu Plain Text mit Zeilenumbrüchen
function htmlToPlainTextWithLineBreaks(html) {
  let text = html.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<\/(p|div|ul|ol|li|table|tr|h[1-6])>/gi, '\n');
  text = text.replace(/<[^>]+>/g, '');
  return text;
}

// 📏 Vorverarbeitung analog Java: HTML->Plain, trim auf 4900 Zeichen
function preprocessBody(htmlContent) {
  const plain = htmlToPlainTextWithLineBreaks(htmlContent);
  const truncated = plain.length > 4900 ? plain.substring(0, 4900) : plain;
  return truncated;
}

// 🕒 Datumformat für MySQL: YYYY-MM-DD HH:MM:SS (lokale Zeit)
function formatDateTimeForMySQL(date) {
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
}

// 🔍 Hozzárendelési státusz lekérdezése
function checkAssignmentStatus() {
  const item = Office.context.mailbox.item;

  item.body.getAsync("html", (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error('Error getting body HTML:', result.error);
      document.getElementById("status").innerText = `Hiba a body lekérdezésénél: ${result.error.message}`;
      return;
    }
    const htmlContent = result.value;
    const processed = preprocessBody(htmlContent);

    // Debug: log der vor dem Hashing verwendeten Zeichenkette
    console.log('Processed body for hash (length=' + processed.length + '):', processed);

    sha256(processed).then(bodyHash => {
      const receivedLocal = formatDateTimeForMySQL(item.dateTimeCreated);
      const payload = {
        subject: item.subject,
        receivedDateTime: receivedLocal,
        from_address: item.from.emailAddress,
        to_address: item.to.length > 0 ? item.to[0].emailAddress : "",
        bodyHash: bodyHash
      };

      console.log('checkAssignmentStatus payload:', payload);

      fetch("https://bogir.hu/V2/api/emails/emails_assignment_check.php", {
        method: "POST",
        headers: {
          "Authorization": "Basic " + btoa("Admin_2024$$:S3cure+P@ssw0rd2024!"),
          "Content-Type": "application/json"
        },
        body: JSON.stringify(payload)
      })
      .then(response => response.json())
      .then(data => {
        console.log('Response data:', data);
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
      })
      .catch(err => {
        console.error('Fetch error:', err);
        document.getElementById("status").innerText = "Hiba a betöltés során.";
      });
    });
  });
}

// 📛 SHA256 Hash-Funktion (UTF-8)
function sha256(str) {
  const encoder = new TextEncoder();
  const data = encoder.encode(str);
  return crypto.subtle.digest('SHA-256', data).then(hash => {
    const hex = Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('');
    console.log('Computed SHA256:', hex);
    return hex;
  });
}

// 🛡️ Hozzárendelés funkció mit gleicher Vorverarbeitung
function assignEmail() {
  const item = Office.context.mailbox.item;
  const statusDiv = document.getElementById("status");
  const assignBtn = document.querySelector("button");

  assignBtn.disabled = true;
  statusDiv.innerText = "Hozzárendelés folyamatban...";

  item.body.getAsync("html", (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error('Error getting body HTML for assign:', result.error);
      statusDiv.innerText = `Hiba a body lekérdezésénél: ${result.error.message}`;
      assignBtn.disabled = false;
      return;
    }
    const htmlContent = result.value;
    const processed = preprocessBody(htmlContent);

    console.log('Processed body for assign (length=' + processed.length + '):', processed);

    sha256(processed).then(bodyHash => {
      const receivedLocal = formatDateTimeForMySQL(item.dateTimeCreated);
      const payload = {
        subject: item.subject,
        receivedDateTime: receivedLocal,
        from_address: item.from.emailAddress,
        to_address: item.to.length > 0 ? item.to[0].emailAddress : "",
        bodyHash: bodyHash,
        assignee: Office.context.mailbox.userProfile.emailAddress
      };

      console.log('assignEmail payload:', payload);

      fetch("https://bogir.hu/V2/api/emails/emails_assignment.php", {
        method: "POST",
        headers: {
          "Authorization": "Basic " + btoa("Admin_2024$$:S3cure+P@ssw0rd2024!"),
          "Content-Type": "application/json"
        },
        body: JSON.stringify(payload)
      })
      .then(response => response.json())
      .then(data => {
        console.log('Assign response data:', data);
        if (data.status === "success") {
          statusDiv.innerText = `Sikeres hozzárendelés: ${data.felelos}`;
        } else {
          statusDiv.innerText = `Hiba a hozzárendelés során: ${data.message || 'Ismeretlen hiba'}`;
          assignBtn.disabled = false;
        }
      })
      .catch(err => {
        console.error('Assign fetch error:', err);
        statusDiv.innerText = "Hiba a hozzárendelés során.";
        assignBtn.disabled = false;
      });
    });
  });
}
