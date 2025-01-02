<?php
// Koneksi ke database
include '../../Koneksi/Koneksi.php';

session_start();

$id_gedung = (int) $_SESSION['id_gedung'] ?? 0;
$id_tower = (int) $_SESSION['id_tower'] ?? 0;
$id_lift = (int) $_GET['id_lift'] ?? 0;

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

    <form id="komponenForm" action="proses_formulir_komponen.php" method="POST" enctype="multipart/form-data">
      <input type="hidden" id="id_lift" name="id_lift" value="<?php echo $id_lift; ?>">
      <input type="hidden" id="id_tower" name="id_tower" value="<?php echo $id_tower; ?>">
      <input type="hidden" id="id_gedung" name="id_gedung" value="<?php echo $id_gedung; ?>">
      <div class="mt-3">
        <div class="form-komponen">
          <div class="card">
            <div class="card-header text-center">Form Komponen</div>
            <div class="card-body">
              <div class="row">
                <!-- Komponen -->
                <div class="col-12 mb-3">
                  <label for="komponen" class="form-label text-gray-700"><span class="required">*</span>Check
                    Item</label>
                  <select id="komponen" name="id_komponen" class="form-select select2 shadow-sm mb-4" required>
                    <option value="">Pilih Komponen</option>
                    <?php
                    if ($result_komponen->num_rows > 0) {
                      while ($row = $result_komponen->fetch_assoc()) {
                        echo '<option value="' . htmlspecialchars($row['id_komponen']) . '">' . htmlspecialchars($row['nama_komponen']) . '</option>';
                      }
                    } else {
                      echo '<option value="">Tidak ada komponen tersedia</option>';
                    }
                    ?>
                  </select>
                </div>

                <!-- Temuan -->
                <div class="col-12 mb-3">
                  <label class="form-label text-gray-700"><span class="required">*</span>Temuan</label>
                  <select id="temuan" name="id_temuan" class="form-select select2 shadow-sm mb-4" required>
                    <option value="">Pilih Temuan</option>
                    <?php
                    while ($row = $temuan_result->fetch_assoc()) {
                      echo "<option value='" . htmlspecialchars($row['id_temuan']) . "'>" . htmlspecialchars($row['nama_temuan']) . "</option>";
                    }
                    ?>
                  </select>
                </div>

                <!-- Solusi -->
                <div class="col-12 mb-3">
                  <label class="form-label text-gray-700"><span class="required">*</span>Solusi</label>
                  <select id="solusi" name="id_solusi" class="form-select select2 shadow-sm mb-4" required>
                    <option value="">Pilih Solusi</option>
                    <?php
                    while ($row = $solusi_result->fetch_assoc()) {
                      echo "<option value='" . htmlspecialchars($row['id_solusi']) . "'>" . htmlspecialchars($row['nama_solusi']) . "</option>";
                    }
                    ?>
                  </select>
                </div>

                <!-- Prioritas -->
                <div class="col-12 mb-3">
                  <label for="prioritas" class="form-label text-gray-700"><span
                      class="required">*</span>Prioritas</label>
                  <select id="prioritas" name="prioritas" class="form-select shadow-sm mb-4" required>
                    <option value="">Pilih Prioritas</option>
                    <option value="1">Prioritas 1</option>
                    <option value="2">Prioritas 2</option>
                    <option value="3">Prioritas 3</option>
                  </select>
                </div>

                <!-- Lantai -->
                <div class="col-12 mb-3">
                  <label for="keterangan" class="form-label text-gray-700">Lantai</label>
                  <textarea id="keterangan" name="keterangan" class="form-control shadow-sm mt-3"></textarea>
                </div>
                <!-- File/Foto -->
                <div class="col-12 mb-3">
                  <label class="form-label">Bukti Foto</label>
                  <input type="file" class="form-control" name="foto_bukti" id="foto_bukti" accept="image/*">
                </div>
                <div class="camera-section col-12">
                  <video class="video-komponen d-none" autoplay></video>
                  <canvas class="canvas-komponen d-none"></canvas>
                  <img class="captured-image-komponen d-none img-thumbnail" alt="Captured Image">
                  <div class="mt-3 mb-3 d-flex justify-content-between">
                    <button type="button" class="btn btn-primary btn-open-camera">Buka Kamera</button>
                    <button type="button" class="btn btn-success btn-capture-image d-none">Ambil Gambar</button>
                  </div>
                </div>

                <!-- Tambah ke Tabel -->
                <div class="card-footer text-center">
                  <button type="button" id="addToTable" class="btn btn-info">Tambah ke Tabel</button>
                </div>

              </div>
            </div>
          </div>
        </div>
      </div>
    </form>

    <!-- Tabel Sementara -->
    <table class="table table-bordered mt-4" id="temporaryTable">
      <thead>
        <tr>
          <th>Komponen</th>
          <th>Temuan</th>
          <th>Solusi</th>
          <th>Prioritas</th>
          <th>Lantai</th>
          <th>Foto</th>
          <th>Aksi</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>

    <!-- Simpan Data ke Server -->
    <div class="text-center mt-3">
      <button id="saveData" class="btn btn-success">Simpan Semua</button>
    </div>

    <script>
    document.addEventListener("DOMContentLoaded", () => {
      const tableBody = document.querySelector("#temporaryTable tbody");
      const addToTableBtn = document.querySelector("#addToTable");
      const saveDataBtn = document.querySelector("#saveData");

      let dataAll = [];

      addToTableBtn.addEventListener("click", () => {

        const id_lift = document.querySelector("#id_lift").value;
        const id_gedung = document.querySelector("#id_gedung").value;
        const id_tower = document.querySelector("#id_tower").value;
        const komponen = document.querySelector("#komponen").value;
        const komponenText = document.querySelector("#komponen option:checked").textContent;
        const temuan = document.querySelector("#temuan").value;
        const temuanText = document.querySelector("#temuan option:checked").textContent;
        const solusi = document.querySelector("#solusi").value;
        const solusiText = document.querySelector("#solusi option:checked").textContent;
        const prioritas = document.querySelector("#prioritas").value;
        const prioritasText = document.querySelector("#prioritas option:checked").textContent;
        const keterangan = document.querySelector("#keterangan").value;
        const fotoInput = document.querySelector("#foto_bukti");
        const fotoFile = fotoInput.files[0];


        const data = {
          id_lift,
          id_gedung,
          id_tower,
          komponen,
          temuan,
          solusi,
          prioritasText,
          keterangan,
          foto: fotoFile
        };

        if (!komponen || !temuan || !solusi || !prioritas) {
          alert("Mohon lengkapi semua field.");
          return;
        }
        dataAll.push(data);
        const row = `
                <tr>
                    <td>${komponenText}<input type="hidden" name="id_komponen" value="${komponen}"></td>
                    <td>${temuanText}<input type="hidden" name="id_temuan" value="${temuan}"></td>
                    <td>${solusiText}<input type="hidden" name="id_solusi" value="${solusi}"></td>
                    <td>${prioritasText}<input type="hidden" name="prioritas" value="${prioritas}"></td>
                    <td>${keterangan}<input type="hidden" name="keterangan" value="${keterangan}"></td>
                    <td>${fotoFile ? fotoFile.name : "Tidak ada foto"}<input type="hidden" name="foto_bukti" value="${fotoFile ? fotoFile.name : ''}"></td>
                    <td><button class="btn btn-danger btn-sm deleteRow">Hapus</button></td>
                </tr>`;
        tableBody.insertAdjacentHTML("beforeend", row);
        document.querySelector("#komponenForm").reset();
      });

      saveDataBtn.addEventListener("click", () => {
        // const rows = [...tableBody.querySelectorAll("tr")];
        if (dataAll.length === 0) {
          alert("Tidak ada data untuk disimpan.");
          return;
        }

        const formData = new FormData();
        dataAll.forEach((row, index) => {
          formData.append(`data[${index}][id_lift]`, row.id_lift);
          formData.append(`data[${index}][id_gedung]`, row.id_gedung);
          formData.append(`data[${index}][id_tower]`, row.id_tower);
          formData.append(`data[${index}][komponen]`, row.komponen);
          formData.append(`data[${index}][temuan]`, row.temuan);
          formData.append(`data[${index}][solusi]`, row.solusi);
          formData.append(`data[${index}][prioritasText]`, row.prioritasText);
          formData.append(`data[${index}][keterangan]`, row.keterangan)

          if (row.foto) {
            formData.append(`data[${index}][foto]`, row.foto)
          }
        })
        // Cek isi formData dengan iterasi
        // for (let [key, value] of formData.entries()) {
        //   console.log(key, value);
        // }
        fetch("Proses/proses_formulir_komponen.php", {
            method: "POST",
            body: formData,
          })
          .then((response) => response.json())
          .then((result) => {
            alert(result.message);
            if (result.status === "success") {
              tableBody.innerHTML = "";
            }

          })
          .catch((error) => {
            console.error(error);
            alert("Terjadi kesalahan.");
          });
      });
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