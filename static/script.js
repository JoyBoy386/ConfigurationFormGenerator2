document.addEventListener('DOMContentLoaded', function () {
    const fileInput = document.getElementById('fileInput');
    const uploadBtn = document.getElementById('uploadBtn');
    const uploadArea = document.getElementById('uploadArea');
    const uploadText = document.querySelector('.upload-text');

    // Handle file selection via button or drag/drop
    function handleFiles(files) {
        if (files.length === 0) return;

        const file = files[0];
        const validTypes = ['text/html', 'application/xhtml+xml'];
        const validExtensions = ['.html', '.htm'];

        const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
        const isValid = validTypes.includes(file.type) || validExtensions.includes(fileExtension);

        if (!isValid) {
            alert('Please upload a valid HTML file (.html or .htm).');
            fileInput.value = ''; // clear invalid selection
            uploadBtn.disabled = true;
            uploadText.innerHTML = 'Drop Assigned Computer Name here<br><small>or click to browse</small>';
            return;
        }

        uploadBtn.disabled = false;
        uploadText.innerHTML = `<strong>${file.name}</strong> selected<br><small>Ready to extract</small>`;
    }

    // File input change
    fileInput.addEventListener('change', function (e) {
        handleFiles(e.target.files);
    });

    // Drag & Drop
    ['dragenter', 'dragover'].forEach(eventName => {
        uploadArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.style.borderColor = '#1c7ed6';
            uploadArea.style.backgroundColor = '#f0f9ff';
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.style.borderColor = '#d1d5da';
            uploadArea.style.backgroundColor = '';
        });
    });

    // Handle drop
    uploadArea.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        fileInput.files = files; // sync with hidden input
        handleFiles(files);
    });

    // Reset button state on init
    uploadBtn.disabled = true;
});