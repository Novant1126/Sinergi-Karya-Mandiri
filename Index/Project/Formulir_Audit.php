<?php
// Koneksi ke database
include '../../Koneksi/Koneksi.php';
session_start();
// Periksa apakah id_tower ada di URL
if (isset($_GET['id_tower'])) {
    $_SESSION['id_tower'] = $_GET['id_tower'];
    $id_tower = $_GET['id_tower'];

    // Query untuk mengambil id_gedung berdasarkan id_tower
    $query = "SELECT id_gedung FROM audit_tower WHERE id_tower = ?";
    $stmt = $conn->prepare($query);
    $stmt->bind_param('i', $id_tower); // Mengikat parameter dengan tipe integer
    $stmt->execute();
    $result = $stmt->get_result();

    // Periksa apakah data ditemukan
    if ($result->num_rows > 0) {
        $row = $result->fetch_assoc();
        $id_gedung = $row['id_gedung'];  // Menyimpan id_gedung
    } else {
        echo "Tower tidak ditemukan.";
        exit;
    }
} else {
    echo "ID Tower tidak ditemukan.";
    exit;
}


// Query untuk mengambil data komponen dari tabel master_komponen
$sql_komponen = "SELECT id_komponen, nama_komponen FROM komponen";
$result_komponen = $conn->query($sql_komponen);

// Ambil data dari tabel solusi_komponen
$solusi_query = "SELECT * FROM solusi_komponen";
$solusi_result = $conn->query($solusi_query);

// Ambil data dari tabel temuan_komponen
$temuan_query = "SELECT * FROM temuan_komponen";
$temuan_result = $conn->query($temuan_query);
?>

<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Form Audit</title>
  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons/font/bootstrap-icons.css" rel="stylesheet">
  <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">

  <link rel="stylesheet" href="css/mobile-formulir_audit.css">

</head>
<!-- Proses/proses_formulir_anudit_komponen.php -->

<body class="bg-gray-100">
  <div class="container-fluid mx-auto mt-10 p-5">
    <h2 class="text-center fw-bold mb-1 text-dark fs-4 fs-md-3 fs-lg-2">Form General Check Lift - PT Sinergi Karya
      Mandiri</h2>
    <p class="text-center mt-0 mb-4 text-muted small">Selamat Menjalankan Tugas Auditnya</p>

    <form id="auditForm" method="POST" action="Proses/proses_formulir_audit_gabungan.php" enctype="multipart/form-data">
      <input type="hidden" name="id_tower" value="<?php echo htmlspecialchars($id_tower); ?>">
      <input type="hidden" name="id_gedung" value="<?php echo htmlspecialchars($id_gedung); ?>">

      <!-- Bagian Lift -->
      <div class="mb-4">
        <div class="row g-3">
          <div class="col-12 col-md-4">
            <label for="liftNo" class="form-label">Nomor Lift</label>
            <input type="text" name="lift_no[]" class="form-control" required>
          </div>
          <div class="col-12 col-md-4">
            <label for="liftBrand" class="form-label">Merek Lift</label>
            <input type="text" name="lift_brand[]" class="form-control">
          </div>
          <div class="col-12 col-md-4">
            <label for="liftType" class="form-label">Type Lift</label>
            <input type="text" name="lift_type[]" class="form-control">
          </div>
        </div>
        <div class="row mt-2" id="cardsContainerUnitTerpasang">
          <!-- Kontainer untuk card -->
        </div>
        <div class="mt-3">
          <button type="button" id="addCardBtnUnitTerpasang" class="btn btn-secondary">Tambah
            Komponen</button>
        </div>
      </div>
      <div class="text-center mt-4">
        <button type="submit" id="saveButton" class="btn btn-primary">Simpan</button>
      </div>
      <!-- Modal untuk pesan sukses
            <div class="modal fade" id="successModal" tabindex="-1" aria-labelledby="successModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="successModalLabel">Sukses</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    Data berhasil disimpan!
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Tutup</button>
                </div>
                </div>
            </div>
            </div> -->

    </form>
    <!-- <script>
        // Menangani pengiriman form menggunakan AJAX
            document.getElementById('auditForm').addEventListener('submit', function(e) {
                e.preventDefault(); // Mencegah form submit biasa

                var formData = new FormData(this); // Mengambil data form

                var xhr = new XMLHttpRequest(); // Membuat objek XMLHttpRequest
                xhr.open('POST', 'Proses/proses_formulir_audit_gabungan.php', true); // Menentukan URL PHP

                // Menangani respons dari server
                xhr.onload = function() {
                    if (xhr.status == 200) {
                        // Jika sukses, tampilkan modal sukses
                        var myModal = new bootstrap.Modal(document.getElementById('successModal'), {});
                        myModal.show(); // Tampilkan modal
                        // Reset form setelah sukses
                        document.getElementById('auditForm').reset();
                    } else {
                        alert('Terjadi kesalahan saat mengirim data.');
                    }
                };

                // Kirim data ke server
                xhr.send(formData);
            });
        </script> -->

    <script>
    document.addEventListener('DOMContentLoaded', function() {
      const cardsContainer = document.getElementById('cardsContainerUnitTerpasang'); // Kontainer card
      const addCardBtn = document.getElementById('addCardBtnUnitTerpasang'); // Tombol tambah card

      // Event listener untuk tombol tambah card
      addCardBtn.addEventListener('click', function() {
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
        deleteCardBtn.addEventListener('click', function() {
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
        openCameraBtn.addEventListener('click', function() {
          navigator.mediaDevices.getUserMedia({
              video: {
                facingMode: currentFacingMode
              }
            })
            .then(function(stream) {
              cameraStream = stream;
              videoElement.srcObject = stream;
              videoElement.play();
              videoElement.classList.remove('d-none');
              takePictureBtn.classList.remove('d-none');
              switchCameraBtn.classList.remove('d-none');
              openCameraBtn.classList.add('d-none');
            })
            .catch(function(error) {
              alert('Tidak dapat mengakses kamera: ' + error.message);
            });
        });

        // Fungsi mengambil gambar dari video
        takePictureBtn.addEventListener('click', function() {
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
        saveImageBtn.addEventListener('click', function() {
          if (!capturedImageElement.src) {
            alert('Ambil gambar terlebih dahulu!');
            return;
          }

          // Cek apakah input file gambar sudah ada di form, jika tidak buat input baru
          let inputFile = card.querySelector('.foto-unit');
          if (!inputFile) {
            // Membuat input file baru untuk menyimpan gambar yang diambil
            inputFile = document.createElement('input');
            inputFile.type = 'file'; // Tipe file
            inputFile.name = 'foto_instalasi[]'; // Nama input sesuai kebutuhan
            inputFile.accept = 'image/*'; // Menerima hanya gambar

            // Menambahkan input file ke dalam form atau card
            card.querySelector('.card-body').appendChild(inputFile);
          }

          // Mengambil tanggal dan waktu saat ini
          const now = new Date();
          const datetime = now.toISOString().replace(/[-:T.]/g, ''); // Format: YYYYMMDDHHMMSS

          // Konversi gambar yang diambil menjadi file Blob untuk diupload melalui form
          const dataURL = capturedImageElement.src;
          const blob = dataURItoBlob(dataURL); // Konversi data URL ke Blob
          const file = new File([blob], `Instalasi_${datetime}.png`, {
            type: 'image/png'
          });

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
          return new Blob([uint8Array], {
            type: mimeString
          });
        }

        // Fungsi beralih kamera
        switchCameraBtn.addEventListener('click', function() {
          currentFacingMode = (currentFacingMode === "environment") ? "user" : "environment";
          if (cameraStream) {
            const tracks = cameraStream.getTracks();
            tracks.forEach(track => track.stop());
          }

          navigator.mediaDevices.getUserMedia({
              video: {
                facingMode: currentFacingMode
              }
            })
            .then(function(stream) {
              cameraStream = stream;
              videoElement.srcObject = stream;
              videoElement.play();
            })
            .catch(function(error) {
              alert('Tidak dapat mengakses kamera: ' + error.message);
            });
        });
      }
    });
    </script>







    <!-- <script src="../../Js/formulir_UnitTerpasang.js"></script> -->
    <!-- <script src="../../Js/formulir_komponen.js"></script> -->

    <!-- Scripts -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.4/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script>
    $(document).ready(function() {
      // Inisialisasi Select2 tanpa fitur hapus pilihan (clear)
      $('.select2').select2({
        theme: 'bootstrap-5',
        placeholder: function() {
          return $(this).data('placeholder');
        },
        allowClear: false // Nonaktifkan tombol silang untuk menghapus pilihan
      });
    });
    </script>
</body>

</html>