
Office.onReady(() => {
  checkAssignmentStatus();
});


function formatDateTimeForMySQL(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  const seconds = String(d.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}


// 🔍 Hozzárendelési státusz lekérdezése
function checkAssignmentStatus() {
  const item = Office.context.mailbox.item;

  item.body.getAsync("text", (result) => {
    const bodyContent = result.value;
    sha256(bodyContent).then(bodyHash => {

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
        const statusDiv = document.getElementById("status");
        if (data.status === "assigned") {
          statusDiv.innerText = `Hozzárendelve: ${data.felelos}`;
        } else if (data.status === "unassigned") {
          statusDiv.innerText = "Még nincs hozzárendelve.";
        } else {
          statusDiv.innerText = "E-mail nincs rögzítve az adatbázisban.";
        }
      })
      .catch(err => {
        document.getElementById("status").innerText = "Hiba a betöltés során.";
      });
    });
  });
}

// 📛 SHA256 funkció
function sha256(str) {
  const encoder = new TextEncoder();
  const data = encoder.encode(str);
  return crypto.subtle.digest('SHA-256', data).then(hash => {
    return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('');
  });
}

// 🛡️ Hozzárendelés funkció (később implementáljuk)
function assignEmail() {
  alert("Hozzárendelés funkció később kerül megvalósításra.");
}
