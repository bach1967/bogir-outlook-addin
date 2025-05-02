Office.onReady(() => {
  checkAssignmentStatus();
});



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
