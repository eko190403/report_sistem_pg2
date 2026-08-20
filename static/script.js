document.addEventListener('DOMContentLoaded', () => {
    const absensiForm = document.getElementById('upload-form');
    const hkoForm = document.getElementById('hko-form');
    
    const masterInput = document.getElementById('master_file');
    const reportInput = document.getElementById('report_file');
    const hkoInput = document.getElementById('hko_file');
    
    const masterName = document.getElementById('master-file-name');
    const reportName = document.getElementById('report-file-name');
    const hkoName = document.getElementById('hko-file-name');
    
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
    setupDragAndDrop('hko-drop-area', hkoInput, hkoName);

    masterInput.addEventListener('change', () => updateFileName(masterInput, masterName));
    reportInput.addEventListener('change', () => updateFileName(reportInput, reportName));
    hkoInput.addEventListener('change', () => updateFileName(hkoInput, hkoName));

    async function handleFormSubmit(e, formElement, inputs, url) {
        e.preventDefault();
        
        for (let input of inputs) {
            if (!input.files.length) {
                alert('Silakan pilih file terlebih dahulu!');
                return;
            }
        }

        const formData = new FormData(formElement);

        formElement.style.display = 'none';
        loadingState.style.display = 'block';

        try {
            const response = await fetch(url, {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = 'Processed.xlsx';
                if (contentDisposition && contentDisposition.indexOf('attachment') !== -1) {
                    const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    const matches = filenameRegex.exec(contentDisposition);
                    if (matches != null && matches[1]) { 
                        filename = matches[1].replace(/['"]/g, '');
                    }
                }

                const blob = await response.blob();
                const objUrl = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = objUrl;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(objUrl);
                
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
    }

    absensiForm.addEventListener('submit', (e) => handleFormSubmit(e, absensiForm, [masterInput, reportInput], '/process'));
    hkoForm.addEventListener('submit', (e) => handleFormSubmit(e, hkoForm, [hkoInput], '/process_hko'));
});

function resetForm() {
    const absensiForm = document.getElementById('upload-form');
    const hkoForm = document.getElementById('hko-form');
    
    absensiForm.reset();
    hkoForm.reset();
    
    document.getElementById('master-file-name').textContent = 'Belum ada file dipilih';
    document.getElementById('report-file-name').textContent = 'Belum ada file dipilih';
    document.getElementById('hko-file-name').textContent = 'Belum ada file dipilih';
    
    document.getElementById('success-state').style.display = 'none';
    document.getElementById('error-state').style.display = 'none';
    
    // Restore the correct form based on active tab
    if (document.getElementById('tab-absensi').classList.contains('active')) {
        absensiForm.style.display = 'block';
    } else {
        hkoForm.style.display = 'block';
    }
}

function switchTab(tabId) {
    document.getElementById('tab-absensi').classList.remove('active');
    document.getElementById('tab-hko').classList.remove('active');
    
    document.getElementById('upload-form').style.display = 'none';
    document.getElementById('hko-form').style.display = 'none';
    document.getElementById('success-state').style.display = 'none';
    document.getElementById('error-state').style.display = 'none';
    
    if (tabId === 'absensi') {
        document.getElementById('tab-absensi').classList.add('active');
        document.getElementById('upload-form').style.display = 'block';
    } else {
        document.getElementById('tab-hko').classList.add('active');
        document.getElementById('hko-form').style.display = 'block';
    }
}
