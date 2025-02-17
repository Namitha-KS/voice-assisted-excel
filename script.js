// Column Setup Page
const columnsContainer = document.getElementById('columns-container');
const addColumnButton = document.getElementById('add-column');
const enterDataButton = document.getElementById('enter-data');
const dataEntryPage = document.getElementById('data-entry-page');
const columnSetupPage = document.getElementById('column-setup-page');

let columns = [];

addColumnButton.addEventListener('click', () => {
  const columnDiv = document.createElement('div');
  columnDiv.className = 'column';
  columnDiv.innerHTML = `
    <input type="text" placeholder="Column Name" class="column-name">
    <select class="column-type">
      <option value="text">Text</option>
      <option value="number">Number</option>
      <option value="total">Total</option>
    </select>
  `;
  columnsContainer.appendChild(columnDiv);
});

enterDataButton.addEventListener('click', () => {
  columns = Array.from(document.querySelectorAll('.column')).map(col => ({
    name: col.querySelector('.column-name').value,
    type: col.querySelector('.column-type').value
  }));
  if (columns.length > 0) {
    columnSetupPage.style.display = 'none';
    dataEntryPage.style.display = 'block';
    initializeTable();
  } else {
    alert('Please add at least one column.');
  }
});

// Data Entry Page
const tableHeaders = document.getElementById('table-headers');
const tableBody = document.getElementById('table-body');
const addRowButton = document.getElementById('add-row');
const startVoiceButton = document.getElementById('start-voice');
const saveButton = document.getElementById('save');
const downloadButton = document.getElementById('download');

let currentRow = null;
let currentCellIndex = 0;
let isVoiceActive = false;
const recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
recognition.continuous = false;
recognition.interimResults = false;

function initializeTable() {
  tableHeaders.innerHTML = '';
  tableBody.innerHTML = '';
  columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col.name;
    tableHeaders.appendChild(th);
  });
}

addRowButton.addEventListener('click', () => {
  const row = document.createElement('tr');
  columns.forEach((col, index) => {
    const cell = document.createElement('td');
    cell.contentEditable = true;
    if (col.type === 'total') {
      cell.contentEditable = false;
      cell.dataset.type = 'total';
    }
    row.appendChild(cell);
  });
  tableBody.appendChild(row);
});

startVoiceButton.addEventListener('click', () => {
  if (!isVoiceActive) {
    isVoiceActive = true;
    currentRow = tableBody.lastElementChild;
    currentCellIndex = 0;
    startVoiceInput();
  }
});

recognition.onresult = (event) => {
  const transcript = event.results[0][0].transcript;
  const currentCell = currentRow.children[currentCellIndex];
  if (columns[currentCellIndex].type === 'number') {
    currentCell.textContent = parseFloat(transcript) || 0;
  } else {
    currentCell.textContent = transcript;
  }
  currentCellIndex++;
  if (currentCellIndex < columns.length) {
    startVoiceInput();
  } else {
    isVoiceActive = false;
    calculateTotals();
  }
};

function startVoiceInput() {
  if (columns[currentCellIndex].type === 'total') {
    currentCellIndex++;
    if (currentCellIndex < columns.length) {
      startVoiceInput();
    } else {
      isVoiceActive = false;
      calculateTotals();
    }
    return;
  }
  recognition.start();
  alert(`Speak for: ${columns[currentCellIndex].name}`);
}

function calculateTotals() {
  tableBody.querySelectorAll('tr').forEach(row => {
    let total = 0;
    row.querySelectorAll('td').forEach((cell, index) => {
      if (columns[index].type === 'number') {
        total += parseFloat(cell.textContent) || 0;
      }
    });
    row.querySelector('[data-type="total"]').textContent = total;
  });
}

saveButton.addEventListener('click', () => {
  isVoiceActive = false;
  recognition.stop();
});

downloadButton.addEventListener('click', () => {
  const wb = XLSX.utils.table_to_book(document.getElementById('data-table'), { sheet: "Sheet 1" });
  XLSX.writeFile(wb, 'table-data.xlsx');
});