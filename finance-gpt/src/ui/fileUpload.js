import { processUploadedFiles } from '../excel/operations.js';

function setupFileUpload() {
  const fileInput = document.getElementById('fileInput');
  const uploadArea = document.querySelector('.upload-area');
  const fileList = document.getElementById('fileList');
  const processButton = document.getElementById('processFiles');
  const uploadedFiles = new Set();

  // Handle drag and drop events
  uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
  });

  uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
  });

  uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    handleFiles(e.dataTransfer.files);
  });

  // Handle file input change
  fileInput.addEventListener('change', (e) => {
    handleFiles(e.target.files);
  });

  // Handle process button click
  processButton.addEventListener('click', () => processUploadedFiles(uploadedFiles, processButton));

  function handleFiles(files) {
    Array.from(files).forEach(file => {
      if (uploadedFiles.has(file.name)) {
        return; // Skip duplicate files
      }

      uploadedFiles.add(file.name);
      const fileItem = document.createElement('div');
      fileItem.className = 'file-item';
      fileItem.innerHTML = `
        <span class="file-item-name">${file.name}</span>
        <button class="file-item-remove" data-filename="${file.name}">Remove</button>
      `;

      fileItem.querySelector('.file-item-remove').addEventListener('click', () => {
        uploadedFiles.delete(file.name);
        fileItem.remove();
        processButton.style.display = uploadedFiles.size > 0 ? 'block' : 'none';
      });

      fileList.appendChild(fileItem);
      processButton.style.display = 'block';
    });
  }
}

export { setupFileUpload }; 