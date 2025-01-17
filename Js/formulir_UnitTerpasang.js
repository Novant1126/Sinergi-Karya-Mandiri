
document.addEventListener('DOMContentLoaded', function () {
    const cardsContainer = document.getElementById('cardsContainerUnitTerpasang'); // Kontainer card
    const addCardBtn = document.getElementById('addCardBtnUnitTerpasang'); // Tombol tambah card

    // Event listener untuk tombol tambah card
    addCardBtn.addEventListener('click', function () {
        // Buat elemen card baru secara dinamis
        const newCard = document.createElement('div');
        newCard.className = 'col-12 col-md-4';
        newCard.innerHTML = `
    <div class="card shadow mt-3">
        <div class="card-header text-center bg-primary text-white fs-7">
            Ambil Gambar Unit Terpasang
        </div>
        <div class="card-body">
            <input type="file" class="form-control shadow-sm mb-1 foto-unit" name="foto_instalasi[]" accept="image/*">
            <div class="camera-container">
                <video class="w-100 rounded mb-1" autoplay></video>
                <canvas class="d-none border border-secondary rounded mb-3"></canvas>
                <img class="d-none img-thumbnail mb-3 captured-image" alt="Captured Image" style="width: 100%; height: auto; object-fit: cover;">
            </div>
            <div class="text-center mb-3">
                <button type="button" class="btn btn-primary btn-sm fs-7 mb-1 open-camera"><i class="fas fa-camera"></i></button>
                <button type="button" class="btn btn-warning btn-sm fs-7 d-none mb-1 take-picture"><i class="fas fa-camera-retro"></i></button>
                <button type="button" class="btn btn-success btn-sm fs-7 d-none mb-1 save-image"><i class="fas fa-save"></i></button>
                <button type="button" class="btn btn-secondary btn-sm fs-7 d-none switch-camera mb-1"><i class="fas fa-exchange-alt"></i></button>
                <button type="button" class="btn btn-danger btn-sm fs-7 mb-1 delete-card"><i class="fas fa-trash"></i></button> <!-- Tombol Hapus Card -->
            </div>
            <input type="text" class="form-control shadow-sm mb-2" name="nama_instalasi[]" placeholder="Nama Unit Terpasang">
            <input type="text" class="form-control shadow-sm mb-1" name="deskripsi_instalasi[]" placeholder="Deskripsi Unit Terpasang">
        </div>
    </div>
`;

        // Tambahkan card baru ke dalam kontainer
        cardsContainer.appendChild(newCard);

        // Inisialisasi fitur kamera pada card baru
        initializeCameraFeatures(newCard);

        // Tambahkan event listener untuk tombol Hapus Card
        const deleteCardBtn = newCard.querySelector('.delete-card');
        deleteCardBtn.addEventListener('click', function () {
            cardsContainer.removeChild(newCard); // Menghapus card
        });
    });

    // Fungsi untuk inisialisasi kamera pada card
    function initializeCameraFeatures(card) {
        const openCameraBtn = card.querySelector('.open-camera');
        const takePictureBtn = card.querySelector('.take-picture');
        const saveImageBtn = card.querySelector('.save-image');
        const switchCameraBtn = card.querySelector('.switch-camera');
        const videoElement = card.querySelector('video');
        const canvasElement = card.querySelector('canvas');
        const capturedImageElement = card.querySelector('.captured-image');
        let cameraStream = null;
        let currentFacingMode = "environment"; // Default kamera belakang

        // Fungsi membuka kamera
        openCameraBtn.addEventListener('click', function () {
            navigator.mediaDevices.getUserMedia({
                video: { facingMode: currentFacingMode }
            })
                .then(function (stream) {
                    cameraStream = stream;
                    videoElement.srcObject = stream;
                    videoElement.play();
                    videoElement.classList.remove('d-none');
                    takePictureBtn.classList.remove('d-none');
                    switchCameraBtn.classList.remove('d-none');
                    openCameraBtn.classList.add('d-none');
                })
                .catch(function (error) {
                    alert('Tidak dapat mengakses kamera: ' + error.message);
                });
        });

        // Fungsi mengambil gambar dari video
        takePictureBtn.addEventListener('click', function () {
            if (!cameraStream) {
                alert('Kamera belum aktif!');
                return;
            }

            canvasElement.width = videoElement.videoWidth;
            canvasElement.height = videoElement.videoHeight;

            const ctx = canvasElement.getContext('2d');
            ctx.drawImage(videoElement, 0, 0, canvasElement.width, canvasElement.height);

            // Tampilkan gambar yang diambil
            const dataURL = canvasElement.toDataURL('image/png');
            capturedImageElement.src = dataURL;
            capturedImageElement.classList.remove('d-none');
            videoElement.classList.add('d-none');

            // Hentikan kamera setelah mengambil gambar
            if (cameraStream) {
                const tracks = cameraStream.getTracks();
                tracks.forEach(track => track.stop());
                cameraStream = null;
            }

            saveImageBtn.classList.remove('d-none');
        });

        // Fungsi menyimpan gambar ke dalam input hidden
        saveImageBtn.addEventListener('click', function () {
            if (!capturedImageElement.src) {
                alert('Ambil gambar terlebih dahulu!');
                return;
            }

            // Cek apakah input file gambar sudah ada di form, jika tidak buat input baru
            let inputFile = card.querySelector('.foto-unit');
            if (!inputFile) {
                // Membuat input file baru untuk menyimpan gambar yang diambil
                inputFile = document.createElement('input');
                inputFile.type = 'file';  // Tipe file
                inputFile.name = 'foto_instalasi[]';  // Nama input sesuai kebutuhan
                inputFile.accept = 'image/*';  // Menerima hanya gambar

                // Menambahkan input file ke dalam form atau card
                card.querySelector('.card-body').appendChild(inputFile);
            }

            // Mengambil tanggal dan waktu saat ini
            const now = new Date();
            const datetime = now.toISOString().replace(/[-:T.]/g, ''); // Format: YYYYMMDDHHMMSS

            // Konversi gambar yang diambil menjadi file Blob untuk diupload melalui form
            const dataURL = capturedImageElement.src;
            const blob = dataURItoBlob(dataURL); // Konversi data URL ke Blob
            const file = new File([blob], `Instalasi_${datetime}.png`, { type: 'image/png' });

            // Set input file dengan file yang baru dibuat
            const dataTransfer = new DataTransfer();
            dataTransfer.items.add(file);
            inputFile.files = dataTransfer.files;

            alert('Gambar berhasil disimpan ke input file.');
        });

        // Fungsi untuk mengonversi Data URL menjadi Blob
        function dataURItoBlob(dataURI) {
            const byteString = atob(dataURI.split(',')[1]);
            const mimeString = dataURI.split(',')[0].split(':')[1].split(';')[0];
            const arrayBuffer = new ArrayBuffer(byteString.length);
            const uint8Array = new Uint8Array(arrayBuffer);
            for (let i = 0; i < byteString.length; i++) {
                uint8Array[i] = byteString.charCodeAt(i);
            }
            return new Blob([uint8Array], { type: mimeString });
        }

        // Fungsi beralih kamera
        switchCameraBtn.addEventListener('click', function () {
            currentFacingMode = (currentFacingMode === "environment") ? "user" : "environment";
            if (cameraStream) {
                const tracks = cameraStream.getTracks();
                tracks.forEach(track => track.stop());
            }

            navigator.mediaDevices.getUserMedia({ video: { facingMode: currentFacingMode } })
                .then(function (stream) {
                    cameraStream = stream;
                    videoElement.srcObject = stream;
                    videoElement.play();
                })
                .catch(function (error) {
                    alert('Tidak dapat mengakses kamera: ' + error.message);
                });
        });
    }
});
