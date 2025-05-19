let guests = [];

function sanitizePhone(phone) {
  if (!phone) return ""; // Return an empty string if phone is null or undefined
  return String(phone).replace(/[^0-9]/g, ""); // Remove non-numeric characters
}

// Convert Excel serial date to JavaScript Date
function formatDate(excelDate) {
  const jsDate = new Date(1900, 0, 1); // Start from Excel's base date (1st Jan 1900)
  jsDate.setDate(jsDate.getDate() + excelDate - 2); // Adjust to Excel's starting point
  const options = {
    year: "numeric",
    month: "long",
    day: "numeric",
    weekday: "long",
  };
  return new Intl.DateTimeFormat("en-US", options).format(jsDate);
}

// Convert Excel time value to a time string
function formatTime(excelTime) {
  const minutes = Math.floor(excelTime * 24 * 60);
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  return `${hours}:${mins < 10 ? "0" : ""}${mins}`;
}

function loadFromStorage() {
  const saved = localStorage.getItem("guests");
  if (saved) {
    guests = JSON.parse(saved);
    displayGuests();
    document.getElementById("status").textContent = "Loaded saved data.";
  }
}

function saveToStorage() {
  localStorage.setItem("guests", JSON.stringify(guests));
}

function clearData() {
  localStorage.removeItem("guests");
  guests = [];
  document.getElementById("guestTable").innerHTML = "";
  document.getElementById("status").textContent = "Data cleared.";
}

function displayGuests() {
  const table = document.getElementById("guestTable");
  table.innerHTML = "";

  const headers = Object.keys(guests[0]);
  headers.push("Action");

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headers.forEach((key) => {
    const th = document.createElement("th");
    th.textContent = key;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  guests.forEach((guest, index) => {
    const row = document.createElement("tr");
    if (guest.status === "done") row.classList.add("done");
    if (guest.status === "no-whatsapp") row.classList.add("no-whatsapp");

    headers.forEach((key) => {
      const td = document.createElement("td");
      if (key === "Action") {
        const btn = document.createElement("button");
        btn.textContent = "Send Message";
        btn.onclick = () => sendMessage(guest, index);
        td.appendChild(btn);
      } else if (key === "Date" && guest[key]) {
        td.textContent = formatDate(guest[key]);
      } else if (key === "Time" && guest[key]) {
        td.textContent = formatTime(guest[key]);
      } else {
        td.textContent = guest[key] ?? "";
      }
      row.appendChild(td);
    });

    tbody.appendChild(row);
  });

  table.appendChild(tbody);
}

function sendMessage(guest, index) {
  const phone = sanitizePhone(guest.Phone);
  if (!phone || phone.length < 10) {
    guests[index].status = "no-whatsapp";
    saveToStorage();
    displayGuests();
    return;
  }

  const message =
    `Hello ${guest.Name}, Thanks for choosing <a href="www.egypttravelist.com">Egypttravelist</a> as a transfer provider.%0A%0A` +
    `Your transfer from ${guest.From} to ${guest.To} on ${formatDate(
      guest.Date
    )} at ${formatTime(guest.Time)} is confirmed.%0A%0A` +
    `Your driver will be waiting for you outside the exit door holding a sign with your name OR a sign of Egypt Travelist.%0A` +
    `If you have any questions, write me anytime. If you can't find the driver, contact us immediately.%0A%0A` +
    `We also do all tours and trips inside Egypt.%0AHave a safe flight!\\nwww.egypttravelist.com\\nThank you!`;

  const url = `https://web.whatsapp.com/send?phone=${phone}&text=${message}`;
  window.open(url, "_blank");

  guests[index].status = "done";
  saveToStorage();
  displayGuests();
}

document.getElementById("excelInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    guests = jsonData.map((g) => ({ ...g, status: "pending" }));
    saveToStorage();
    displayGuests();
    document.getElementById("status").textContent = "Excel data loaded.";
  };

  reader.readAsArrayBuffer(file);
});

document.getElementById("clearData").addEventListener("click", clearData);

window.onload = loadFromStorage;
