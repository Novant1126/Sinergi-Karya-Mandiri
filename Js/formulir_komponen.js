document.addEventListener('DOMContentLoaded', function () {
    const form = document.querySelector('.form-komponen');
    initializeCamera(form);

    document.querySelector('#addKomponenBtn').addEventListener('click', () => {
        const komponen = document.querySelector('[name="id_komponen[]"]').value;
        const komponenText = document.querySelector('[name="id_komponen[]"] option:checked').text;
        const temuan = document.querySelector('[name="id_temuan[]"]').value;
        const temuanText = document.querySelector('[name="id_temuan[]"] option:checked').text;
        const solusi = document.querySelector('[name="id_solusi[]"]').value;
        const solusiText = document.querySelector('[name="id_solusi[]"] option:checked').text;
        const prioritas = document.querySelector('[name="prioritas[]"]').value;
        const prioritasText = document.querySelector('[name="prioritas[]"] option:checked').text;
        const keterangan = document.querySelector('[name="keterangan[]"]').value;
        const fotoInput = document.querySelector('[name="foto_bukti[]"]');
        const fotoFile = fotoInput.files[0];
    
        // Validasi input
        if (!komponen || !temuan || !solusi || !prioritas || !fotoFile) {
            alert('Harap lengkapi semua field yang bertanda * sebelum menambahkan.');
            return;
        }
    
        const reader = new FileReader();
        reader.onload = function (e) {
            const tableBody = document.querySelector('#komponenTable tbody');
            const newRow = document.createElement('tr');
            newRow.innerHTML = `
                <td>${komponenText}<input type="hidden" name="id_komponen[]" value="${komponen}"></td>
                <td>${temuanText}<input type="hidden" name="id_temuan[]" value="${temuan}"></td>
                <td>${solusiText}<input type="hidden" name="id_solusi[]" value="${solusi}"></td>
                <td>${prioritasText}<input type="hidden" name="prioritas[]" value="${prioritas}"></td>
                <td>${keterangan}<input type="hidden" name="keterangan[]" value="${keterangan}"></td>
                <td><img src="${e.target.result}" class="img-thumbnail" style="max-width: 100px;"><input type="hidden" name="foto_bukti_base64[]" value="${e.target.result}"></td>
                <td><button type="button" class="btn btn-danger btn-delete-row">Hapus</button></td>
            `;
            tableBody.appendChild(newRow);
    
            // Event untuk menghapus baris
            newRow.querySelector('.btn-delete-row').addEventListener('click', () => {
                newRow.remove();
            });
    
            // Reset form setelah menambahkan ke tabel
            form.reset();
        };
        reader.readAsDataURL(fotoFile);
    });
    

    // Inisialisasi kamera
    function initializeCamera(form) {
        const video = form.querySelector('.video-komponen');
        const canvas = form.querySelector('.canvas-komponen');
        const btnOpenCamera = form.querySelector('.btn-open-camera');
        const btnCaptureImage = form.querySelector('.btn-capture-image');
        const fotoInput = form.querySelector('.bukti_foto');

        let stream = null;

        btnOpenCamera.addEventListener('click', () => {
            navigator.mediaDevices.getUserMedia({ video: true })
                .then(mediaStream => {
                    stream = mediaStream;
                    video.srcObject = stream;
                    video.classList.remove('d-none');
                    btnCaptureImage.classList.remove('d-none');
                    video.play();
                })
                .catch(() => alert('Tidak dapat mengakses kamera.'));
        });

        btnCaptureImage.addEventListener('click', () => {
            canvas.width = video.videoWidth;
            canvas.height = video.videoHeight;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

            const dataURL = canvas.toDataURL('image/png');

            const blob = dataURLToBlob(dataURL);
            const file = new File([blob], `foto_${Date.now()}.png`, { type: 'image/png' });

            const dataTransfer = new DataTransfer();
            dataTransfer.items.add(file);
            fotoInput.files = dataTransfer.files;

            const fileNameDisplay = form.querySelector('.file-name-display');
            if (!fileNameDisplay) {
                const newFileNameDisplay = document.createElement('p');
                newFileNameDisplay.classList.add('file-name-display');
                newFileNameDisplay.textContent = `File Tersimpan: ${file.name}`;
                fotoInput.parentNode.appendChild(newFileNameDisplay);
            } else {
                fileNameDisplay.textContent = `File Tersimpan: ${file.name}`;
            }

            video.classList.add('d-none');
            btnCaptureImage.classList.add('d-none');

            if (stream) {
                stream.getTracks().forEach(track => track.stop());
            }
        });
    }

    function dataURLToBlob(dataURL) {
        const [header, base64Data] = dataURL.split(',');
        const mime = header.match(/:(.*?);/)[1];
        const binaryData = atob(base64Data);
        const byteArray = new Uint8Array(binaryData.length);

        for (let i = 0; i < binaryData.length; i++) {
            byteArray[i] = binaryData.charCodeAt(i);
        }

        return new Blob([byteArray], { type: mime });
    }
});
