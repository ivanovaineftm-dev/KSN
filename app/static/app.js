const form = document.getElementById('upload-form');
const mainFileInput = document.getElementById('main-file-input');
const locationsFileInput = document.getElementById('locations-file-input');
const submitBtn = document.getElementById('submit-btn');
const statusNode = document.getElementById('status');
const analyticsSection = document.getElementById('analytics-section');
const analyticsEmpty = document.getElementById('analytics-empty');
const analyticsChart = document.getElementById('analytics-chart');

function setStatus(message, type = '') {
  statusNode.textContent = message;
  statusNode.className = `status ${type}`.trim();
}

function renderAnalytics(analytics) {
  analyticsChart.innerHTML = '';

  if (!analytics.length) {
    analyticsSection.hidden = false;
    analyticsEmpty.hidden = false;
    return;
  }

  analyticsSection.hidden = false;
  analyticsEmpty.hidden = true;

  analytics.forEach((item) => {
    const barItem = document.createElement('div');
    barItem.className = 'bar-item';

    const value = document.createElement('div');
    value.className = 'bar-value';
    value.textContent = `${item.quality}%`;

    const barWrap = document.createElement('div');
    barWrap.className = 'bar-wrap';

    const bar = document.createElement('div');
    bar.className = 'bar';
    bar.style.height = `${Math.max(0, Math.min(item.quality, 100))}%`;
    bar.title = `${item.department}: ${item.valid_rows}/${item.total_rows} (${item.quality}%)`;

    const label = document.createElement('div');
    label.className = 'bar-label';
    label.textContent = item.department;

    barWrap.appendChild(bar);
    barItem.appendChild(value);
    barItem.appendChild(barWrap);
    barItem.appendChild(label);
    analyticsChart.appendChild(barItem);
  });
}

async function downloadProcessedFile(downloadUrl, filename) {
  const response = await fetch(downloadUrl);
  if (!response.ok) {
    throw new Error('Не удалось скачать обработанный файл.');
  }

  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  const mainFile = mainFileInput.files?.[0];
  const locationsFile = locationsFileInput.files?.[0];

  if (!mainFile) {
    setStatus('Выберите основной файл перед загрузкой.', 'error');
    return;
  }

  submitBtn.disabled = true;
  setStatus('Файлы загружаются и обрабатываются...');

  try {
    const formData = new FormData();
    formData.append('main_file', mainFile);
    if (locationsFile) {
      formData.append('locations_file', locationsFile);
    }

    const response = await fetch('/upload', {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      const detail = payload.detail || 'Ошибка обработки файла.';
      throw new Error(detail);
    }

    const payload = await response.json();
    renderAnalytics(payload.analytics || []);
    await downloadProcessedFile(payload.download_url, payload.filename || `processed_${mainFile.name}`);

    setStatus('Готово! Обработанный файл скачан, аналитика построена.', 'ok');
  } catch (error) {
    setStatus(error.message || 'Произошла ошибка.', 'error');
  } finally {
    submitBtn.disabled = false;
  }
});
