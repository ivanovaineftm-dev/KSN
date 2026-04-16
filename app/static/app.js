const form = document.getElementById('upload-form');
const processBtn = document.getElementById('process-btn');
const downloadBtn = document.getElementById('download-btn');
const clearBtn = document.getElementById('clear-btn');
const statusNode = document.getElementById('status');
const dashboard = document.getElementById('dashboard');

const mainFileInput = document.getElementById('main-file-input');
const locationsFileInput = document.getElementById('locations-file-input');
const baristaFileInput = document.getElementById('barista-file-input');

const totalRowsNode = document.getElementById('total-rows');
const errorRowsNode = document.getElementById('error-rows');
const validRowsNode = document.getElementById('valid-rows');
const tableBody = document.getElementById('analytics-table-body');

const ALLOWED_EXTENSIONS = ['.xlsx', '.xls'];

let currentFileId = null;
let departmentsData = [];
let sortConfig = { key: 'quality', direction: 'desc' };
let summaryChart = null;
let departmentChart = null;

function setStatus(message, type = 'default', isLoading = false) {
  statusNode.className = `status ${type === 'error' ? 'error' : ''} ${type === 'success' ? 'success' : ''}`.trim();
  statusNode.innerHTML = isLoading ? `<span class="loader"></span>${message}` : message;
}

function getFileNameNodeByInputId(inputId) {
  const map = {
    'main-file-input': 'main-file-name',
    'locations-file-input': 'locations-file-name',
    'barista-file-input': 'barista-file-name',
  };
  return document.getElementById(map[inputId]);
}

function isValidExcelFile(file) {
  if (!file) return false;
  const loweredName = file.name.toLowerCase();
  return ALLOWED_EXTENSIONS.some((ext) => loweredName.endsWith(ext));
}

function setInputFile(input, file) {
  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(file);
  input.files = dataTransfer.files;
  updateFileName(input);
}

function updateFileName(input) {
  const nameNode = getFileNameNodeByInputId(input.id);
  const file = input.files?.[0];
  nameNode.textContent = file ? file.name : 'Файл не выбран';
}

function attachDropZone(dropZone) {
  const inputId = dropZone.dataset.input;
  const input = document.getElementById(inputId);

  input.addEventListener('change', () => updateFileName(input));

  ['dragenter', 'dragover'].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    dropZone.addEventListener(eventName, () => {
      dropZone.classList.remove('dragover');
    });
  });

  dropZone.addEventListener('drop', (event) => {
    event.preventDefault();
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;

    if (!isValidExcelFile(file)) {
      setStatus('Допустимы только файлы .xlsx и .xls.', 'error');
      return;
    }

    setInputFile(input, file);
    setStatus('');
  });
}

function renderTableRows() {
  const sortedData = [...departmentsData].sort((a, b) => {
    const valueA = a[sortConfig.key];
    const valueB = b[sortConfig.key];

    if (typeof valueA === 'string') {
      const cmp = valueA.localeCompare(String(valueB), 'ru');
      return sortConfig.direction === 'asc' ? cmp : -cmp;
    }

    const cmp = Number(valueA) - Number(valueB);
    return sortConfig.direction === 'asc' ? cmp : -cmp;
  });

  tableBody.innerHTML = '';
  sortedData.forEach((item) => {
    const tr = document.createElement('tr');
    if (item.quality < 100) {
      tr.className = 'error-row';
    }

    tr.innerHTML = `
      <td>${item.department}</td>
      <td>${item.quality}%</td>
      <td>${item.total_rows}</td>
      <td>${item.valid_rows}</td>
    `;
    tableBody.appendChild(tr);
  });
}

function destroyCharts() {
  if (summaryChart) {
    summaryChart.destroy();
    summaryChart = null;
  }
  if (departmentChart) {
    departmentChart.destroy();
    departmentChart = null;
  }
}

function renderCharts(payload) {
  destroyCharts();

  summaryChart = new Chart(document.getElementById('summary-chart'), {
    type: 'pie',
    data: {
      labels: ['Ошибки', 'Корректные строки'],
      datasets: [{
        data: [payload.errors, payload.valid_rows],
        backgroundColor: ['rgba(220, 38, 38, 0.6)', 'rgba(37, 99, 235, 0.7)'],
        borderWidth: 1,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
    },
  });

  departmentChart = new Chart(document.getElementById('department-chart'), {
    type: 'bar',
    data: {
      labels: departmentsData.map((item) => item.department),
      datasets: [{
        label: 'КСН, %',
        data: departmentsData.map((item) => item.quality),
        backgroundColor: departmentsData.map((item) => (item.quality < 100 ? 'rgba(220, 38, 38, 0.45)' : 'rgba(37, 99, 235, 0.8)')),
        borderWidth: 1,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          title: { display: true, text: '%' },
        },
        x: {
          title: { display: true, text: 'Подразделения' },
        },
      },
    },
  });
}

async function processFiles(mainFile, locationsFile, baristaFile) {
  const formData = new FormData();
  formData.append('main_file', mainFile);
  formData.append('locations_file', locationsFile);
  formData.append('barista_file', baristaFile);

  const uploadResponse = await fetch('/upload/', {
    method: 'POST',
    body: formData,
  });
  if (!uploadResponse.ok) {
    const errPayload = await uploadResponse.json().catch(() => ({}));
    throw new Error(errPayload.detail || 'Ошибка загрузки файлов.');
  }

  const uploadPayload = await uploadResponse.json();
  currentFileId = uploadPayload.file_id;

  const processResponse = await fetch('/process/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ file_id: currentFileId }),
  });

  if (!processResponse.ok) {
    const errPayload = await processResponse.json().catch(() => ({}));
    throw new Error(errPayload.detail || 'Ошибка при обработке файла.');
  }

  return processResponse.json();
}

function resetState() {
  [mainFileInput, locationsFileInput, baristaFileInput].forEach((input) => {
    input.value = '';
    updateFileName(input);
  });

  dashboard.classList.add('hidden');
  tableBody.innerHTML = '';
  totalRowsNode.textContent = '0';
  errorRowsNode.textContent = '0';
  validRowsNode.textContent = '0';
  departmentsData = [];
  currentFileId = null;
  downloadBtn.disabled = true;
  setStatus('');
  destroyCharts();
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  const mainFile = mainFileInput.files?.[0];
  const locationsFile = locationsFileInput.files?.[0];
  const baristaFile = baristaFileInput.files?.[0];

  if (!mainFile || !locationsFile || !baristaFile) {
    setStatus('Загрузите все 3 файла перед обработкой.', 'error');
    return;
  }

  if (![mainFile, locationsFile, baristaFile].every((file) => isValidExcelFile(file))) {
    setStatus('Допустимы только файлы .xlsx и .xls.', 'error');
    return;
  }

  processBtn.disabled = true;
  downloadBtn.disabled = true;
  setStatus('Выполняется загрузка и обработка...', 'default', true);

  try {
    const payload = await processFiles(mainFile, locationsFile, baristaFile);
    departmentsData = payload.departments || [];

    totalRowsNode.textContent = String(payload.total_rows || 0);
    errorRowsNode.textContent = String(payload.errors || 0);
    validRowsNode.textContent = String(payload.valid_rows || 0);

    renderCharts(payload);
    renderTableRows();

    dashboard.classList.remove('hidden');
    downloadBtn.disabled = false;
    setStatus('Обработка завершена успешно.', 'success');
  } catch (error) {
    dashboard.classList.add('hidden');
    setStatus(error.message || 'Произошла ошибка.', 'error');
  } finally {
    processBtn.disabled = false;
  }
});

downloadBtn.addEventListener('click', () => {
  if (!currentFileId) {
    setStatus('Нет файла для скачивания. Сначала выполните обработку.', 'error');
    return;
  }
  window.location.href = `/download/${currentFileId}`;
});

clearBtn.addEventListener('click', resetState);

document.querySelectorAll('.drop-zone').forEach(attachDropZone);
document.querySelectorAll('th[data-sort]').forEach((th) => {
  th.addEventListener('click', () => {
    const key = th.dataset.sort;
    if (sortConfig.key === key) {
      sortConfig.direction = sortConfig.direction === 'asc' ? 'desc' : 'asc';
    } else {
      sortConfig = { key, direction: key === 'department' ? 'asc' : 'desc' };
    }
    renderTableRows();
  });
});
