const form = document.getElementById('upload-form');
const fileInput = document.getElementById('file-input');
const submitBtn = document.getElementById('submit-btn');
const statusNode = document.getElementById('status');

function setStatus(message, type = '') {
  statusNode.textContent = message;
  statusNode.className = `status ${type}`.trim();
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  const file = fileInput.files?.[0];
  if (!file) {
    setStatus('Выберите файл перед загрузкой.', 'error');
    return;
  }

  submitBtn.disabled = true;
  setStatus('Файл загружается и обрабатывается...');

  try {
    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch('/upload', {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      const detail = payload.detail || 'Ошибка обработки файла.';
      throw new Error(detail);
    }

    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    const fallbackName = `processed_${file.name}`;
    const filenameMatch = disposition.match(/filename="?([^\"]+)"?/i);
    const filename = filenameMatch?.[1] || fallbackName;

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);

    setStatus('Готово! Обработанный файл скачан.', 'ok');
  } catch (error) {
    setStatus(error.message || 'Произошла ошибка.', 'error');
  } finally {
    submitBtn.disabled = false;
  }
});
