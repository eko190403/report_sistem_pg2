document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('upload-form');
    const masterInput = document.getElementById('master_file');
    const reportInput = document.getElementById('report_file');
    const masterName = document.getElementById('master-file-name');
    const reportName = document.getElementById('report-file-name');
    const submitBtn = document.getElementById('submit-btn');
    
    // UI States
    const loadingState = document.getElementById('loading-state');
    const successState = document.getElementById('success-state');
    const errorState = document.getElementById('error-state');
    const errorMessage = document.getElementById('error-message');

    // Drag and drop visual feedback
    const setupDragAndDrop = (dropAreaId, inputElement, nameElement) => {
        const dropArea = document.getElementById(dropAreaId);
        
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, preventDefaults, false);
        });
        
        ['dragenter', 'dragover'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.add('dragover'), false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.remove('dragover'), false);
        });
        
        dropArea.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            inputElement.files = files;
            updateFileName(inputElement, nameElement);
        }, false);
    };

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function updateFileName(input, displayElement) {
        if (input.files.length > 0) {
            displayElement.textContent = input.files[0].name;
            displayElement.style.color = 'var(--text)';
            displayElement.style.background = 'rgba(255, 255, 255, 0.1)';
        } else {
            displayElement.textContent = 'Belum ada file dipilih';
            displayElement.style.color = 'var(--primary)';
            displayElement.style.background = 'rgba(0, 242, 96, 0.1)';
        }
    }

    setupDragAndDrop('master-drop-area', masterInput, masterName);
    setupDragAndDrop('report-drop-area', reportInput, reportName);

    masterInput.addEventListener('change', () => updateFileName(masterInput, masterName));
    reportInput.addEventListener('change', () => updateFileName(reportInput, reportName));

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        if (!masterInput.files.length || !reportInput.files.length) {
            alert('Silakan pilih kedua file terlebih dahulu!');
            return;
        }

        const formData = new FormData(form);

        // Hide form, show loading
        form.style.display = 'none';
        loadingState.style.display = 'block';

        try {
            const response = await fetch('/process', {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                // Get filename from header if possible, or use default
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = 'Processed_Report.xlsx';
                if (contentDisposition && contentDisposition.indexOf('attachment') !== -1) {
                    const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    const matches = filenameRegex.exec(contentDisposition);
                    if (matches != null && matches[1]) { 
                        filename = matches[1].replace(/['"]/g, '');
                    }
                }

                // Create blob and trigger download
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Show success
                loadingState.style.display = 'none';
                successState.style.display = 'block';
            } else {
                const errorText = await response.text();
                throw new Error(errorText || 'Gagal memproses file.');
            }
        } catch (error) {
            console.error(error);
            loadingState.style.display = 'none';
            errorState.style.display = 'block';
            errorMessage.textContent = error.message;
        }
    });
});

function resetForm() {
    const form = document.getElementById('upload-form');
    
    // Reset file inputs
    form.reset();
    document.getElementById('master-file-name').textContent = 'Belum ada file dipilih';
    document.getElementById('report-file-name').textContent = 'Belum ada file dipilih';
    
    // Reset UI states
    document.getElementById('success-state').style.display = 'none';
    document.getElementById('error-state').style.display = 'none';
    form.style.display = 'block';
}
