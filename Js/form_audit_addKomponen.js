document.getElementById('addKomponenBtn').addEventListener('click', addKomponen);
document.getElementById('hapusKomponenBtn').addEventListener('click', hapusKomponen);

function initializeCamera(container) {
    const video = container.querySelector('#videoKomponen');
    const canvas = container.querySelector('#canvasKomponen');
    const img = container.querySelector('#capturedImageKomponen');
    const openCameraButton = container.querySelector('#openCameraKomponen');
    const captureImageButton = container.querySelector('#captureImageKomponen');
    const saveImageButton = container.querySelector('#saveImageKomponen');
    const inputFile = container.querySelector('input[type="file"]');

    let stream;
    openCameraButton.addEventListener('click', async () => {
        try {
            stream = await navigator.mediaDevices.getUserMedia({ video: true });
            video.srcObject = stream;
            video.classList.remove('d-none');
            video.play();

            img.classList.add('d-none'); // Sembunyikan gambar jika kamera dibuka kembali
            canvas.classList.add('d-none'); // Sembunyikan canvas

            openCameraButton.classList.add('d-none');
            captureImageButton.classList.remove('d-none');
        } catch (error) {
            console.error('Kamera tidak dapat diakses:', error.message);
        }
    });

    captureImageButton.addEventListener('click', () => {
        const context = canvas.getContext('2d');
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        context.drawImage(video, 0, 0, canvas.width, canvas.height);

        img.src = canvas.toDataURL('image/png');
        img.classList.remove('d-none');
        video.classList.add('d-none'); // Sembunyikan video setelah gambar diambil

        captureImageButton.classList.add('d-none');
        saveImageButton.classList.remove('d-none');
    });

    saveImageButton.addEventListener('click', () => {
        const context = canvas.getContext('2d');
        context.drawImage(video, 0, 0, canvas.width, canvas.height);

        canvas.toBlob(blob => {
            const fileName = `gambar_${new Date().toISOString().slice(0, 19).replace(/[-:T]/g, '')}.png`;
            const file = new File([blob], fileName, { type: 'image/png' });

            const dataTransfer = new DataTransfer();
            dataTransfer.items.add(file);
            inputFile.files = dataTransfer.files;
        }, 'image/png');
    });

    container.querySelector('#hapusKomponenBtn').addEventListener('click', () => {
        if (stream) {
            stream.getTracks().forEach(track => track.stop());
        }
    });
}

function addKomponen(e) {
    e.preventDefault();

    const komponenContainer = document.getElementById('komponenContainer');
    const komponenRow = komponenContainer.querySelector('.komponenRow').cloneNode(true);

    const inputs = komponenRow.querySelectorAll('select, input, textarea');
    inputs.forEach(input => {
        input.value = '';
        if (input.type === 'file') {
            input.value = '';
        }
    });

    const selects = komponenRow.querySelectorAll('select');
    selects.forEach(select => {
        $(select).val(null).trigger('change');
    });

    komponenContainer.appendChild(komponenRow);

    // Inisialisasi kamera untuk komponen baru setelah ditambahkan ke DOM
    initializeCamera(komponenRow);
}

function hapusKomponen(e) {
    e.preventDefault();

    const komponenContainer = document.getElementById('komponenContainer');
    const komponenRows = komponenContainer.querySelectorAll('.komponenRow');

    if (komponenRows.length > 1) {
        komponenContainer.removeChild(komponenRows[komponenRows.length - 1]);
    } else {
        alert('Tidak ada komponen untuk dihapus');
    }
}

// Inisialisasi kamera pada komponen pertama kali
initializeCamera(document.querySelector('.komponenRow'));
