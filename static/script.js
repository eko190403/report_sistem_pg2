document.addEventListener('DOMContentLoaded', () => {
    const absensiForm = document.getElementById('upload-form');
    const hkoForm = document.getElementById('hko-form');
    
    const masterInput = document.getElementById('master_file');
    const reportInput = document.getElementById('report_file');
    const hkoInput = document.getElementById('hko_file');
    
    const masterName = document.getElementById('master-file-name');
    const reportName = document.getElementById('report-file-name');
    const hkoName = document.getElementById('hko-file-name');
    
    const loaderOverlay = document.getElementById('loader-overlay');
    const progressBar = document.getElementById('progress-bar');
    const progressText = document.getElementById('progress-text');
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
            if (input.files.length > 6) {
                alert('Maksimal hanya boleh memilih 6 file!');
                input.value = '';
                displayElement.textContent = 'Belum ada file dipilih';
                displayElement.style.color = 'var(--primary)';
                displayElement.style.background = 'rgba(0, 242, 96, 0.1)';
                return;
            }
            if (input.files.length === 1) {
                displayElement.textContent = input.files[0].name;
            } else {
                displayElement.textContent = `${input.files.length} file terpilih`;
            }
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
        
        // Reset progress bar
        loaderOverlay.classList.add('active');
        progressBar.style.width = '0%';
        progressText.textContent = '0%';
        
        // Kalkulasi estimasi waktu (4 detik per file dengan engine calamine)
        let numFiles = 0;
        for (let input of inputs) {
            numFiles += input.files.length;
        }
        const estimatedTimeMs = Math.max(3000, numFiles * 4000);
        const updateInterval = 100; // update setiap 100ms
        const totalSteps = estimatedTimeMs / updateInterval;
        const increment = 95 / totalSteps;
        
        let progress = 0;
        const progressInterval = setInterval(() => {
            if (progress < 95) {
                progress += increment;
                let currentProgress = Math.min(95, progress + (Math.random() * 1.5));
                progressBar.style.width = currentProgress + '%';
                progressText.textContent = Math.floor(currentProgress) + '%';
            }
        }, updateInterval);

        try {
            const response = await fetch(url, {
                method: 'POST',
                body: formData
            });
            
            clearInterval(progressInterval);

            if (response.ok) {
                // Selesaikan progress bar
                progressBar.style.width = '100%';
                progressText.textContent = '100%';
                
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
                
                setTimeout(() => {
                    loaderOverlay.classList.remove('active');
                    successState.style.display = 'block';
                }, 500);
                
            } else {
                const errorText = await response.text();
                throw new Error(errorText || 'Gagal memproses file.');
            }
        } catch (error) {
            console.error(error);
            clearInterval(progressInterval);
            loaderOverlay.classList.remove('active');
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
