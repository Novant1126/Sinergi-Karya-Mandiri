<?php
// Koneksi ke database
include '../../Koneksi/Koneksi.php';

// Periksa apakah id_tower ada di URL
if (isset($_GET['id_tower'])) {
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

    <!-- Tailwind CSS -->
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@3.3.2/dist/tailwind.min.css" rel="stylesheet">

    <!-- Select2 CSS -->
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet">
    <link rel="stylesheet" href="css/mobile-formulir_audit.css">

</head>
<!-- Proses/proses_formulir_anudit_komponen.php -->
<body class="bg-gray-100">
    <div class="container-fluid mx-auto mt-10 p-5">
        <h2 class="text-center fw-bold mb-1 text-dark fs-4 fs-md-3 fs-lg-2">Form General Check Lift - PT Sinergi Karya
            Mandiri</h2>
        <p class="text-center mt-0 mb-4 text-muted small">Selamat Menjalankan Tugas Auditnya</p>

        <form id="auditForm" method="POST" action="Proses/proses_formulir_audit_gabungan.php" enctype="multipart/form-data">
        <input type="hidden" name="id_tower" value="<?php echo htmlspecialchars($id_tower); ?>"> <!-- Mengirim id_tower ke proses PHP -->
        <input type="hidden" name="id_gedung" value="<?php echo htmlspecialchars($id_gedung); ?>"> <!-- Mengirim id_gedung ke proses PHP -->
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
            <div class="col-12 col-md-4 template-card" style="display: none;">
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
                            <button type="button" class="btn btn-primary btn-sm fs-7 mb-1 open-camera">Buka Kamera</button>
                            <button type="button" class="btn btn-warning btn-sm fs-7 d-none mb-1 take-picture">Ambil Gambar</button>
                            <button type="button" class="btn btn-success btn-sm fs-7 d-none mb-1 save-image">Simpan Gambar</button>
                            <button type="button" class="btn btn-secondary btn-sm fs-7 d-none switch-camera mb-1">Beralih Kamera</button>
                        </div>
                        <input type="text" class="form-control shadow-sm mb-2" name="nama_instalasi[]" placeholder="Nama Unit Terpasang">
                        <input type="text" class="form-control shadow-sm mb-1" name="deskripsi_instalasi[]" placeholder="Deskripsi Unit Terpasang">
                    </div>
                </div>
            </div>
        </div>
        <div class="mt-3">
            <button type="button" id="addCardBtnUnitTerpasang" class="btn btn-secondary">Tambah Komponen</button>
        </div>
    </div>

    <!-- Bagian Komponen -->
    <div class="mt-3">
        <div class="form-komponen">
            <div class="card">
                <div class="card-header text-center">Form Komponen</div>
                <div class="card-body">
                    <div class="row">
                        <div class="col-12 mb-3">
                            <label for="komponen" class="form-label text-gray-700"><span class="required">*</span>Chek Item</label>
                            <select class="form-select select2 shadow-sm mb-4" required>
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
                        <div class="col-12 mb-3">
                            <label class="form-label text-gray-700"><span class="required">*</span>Temuan</label>
                            <select class="form-select select2 shadow-sm mb-4" required>
                                <option value="">Pilih Temuan</option>
                                <?php
                                while ($row = $temuan_result->fetch_assoc()) {
                                    echo "<option value='" . htmlspecialchars($row['id_temuan']) . "'>" . htmlspecialchars($row['nama_temuan']) . "</option>";
                                }
                                ?>
                            </select>
                        </div>
                        <div class="col-12 mb-3">
                            <label class="form-label text-gray-700"><span class="required">*</span>Solusi</label>
                            <select class="form-select select2 shadow-sm mb-4" required>
                                <option value="">Pilih Solusi</option>
                                <?php
                                while ($row = $solusi_result->fetch_assoc()) {
                                    echo "<option value='" . htmlspecialchars($row['id_solusi']) . "'>" . htmlspecialchars($row['nama_solusi']) . "</option>";
                                }
                                ?>
                            </select>
                        </div>
                        <div class="col-12 mb-3">
                            <label for="prioritas" class="form-label text-gray-700"><span class="required">*</span>Prioritas</label>
                            <select class="form-select shadow-sm mb-4" required>
                                <option value="">Pilih Prioritas</option>
                                <option value="1">Prioritas 1</option>
                                <option value="2">Prioritas 2</option>
                                <option value="3">Prioritas 3</option>
                            </select>
                        </div>
                        <div class="col-12 mb-3">
                            <label for="keterangan" class="form-label text-gray-700">Lantai</label>
                            <textarea class="form-control shadow-sm mt-3"></textarea>
                        </div>
                        <div class="col-12 mb-3">
                            <label class="form-label">Bukti Foto</label>
                            <input type="file" class="form-control bukti_foto" name="foto_bukti[]" accept="image/*">
                        </div>
                        <div class="camera-section col-12">
                            <video class="video-komponen d-none" autoplay></video>
                            <canvas class="canvas-komponen d-none"></canvas>
                            <img class="captured-image-komponen d-none img-thumbnail" alt="Captured Image">
                            <div class="mt-3 d-flex justify-content-between">
                                <button type="button" class="btn btn-primary btn-open-camera">Buka Kamera</button>
                                <button type="button" class="btn btn-primary btn-open-camera">Balik gambar</button>
                                <button type="button" class="btn btn-success btn-capture-image d-none">Ambil Gambar</button>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="card-footer text-center">
                    <button type="button" id="addKomponenBtn" class="btn btn-info">Tambah Komponen</button>
                </div>
            </div>
        </div>

        <!-- Tabel -->
        <div class="table-responsive mt-4">
            <table id="komponenTable" class="table">
                <thead>
                    <tr>
                        <th>Komponen</th>
                        <th>Temuan</th>
                        <th>Solusi</th>
                        <th>Prioritas</th>
                        <th>Keterangan</th>
                        <th>Foto Bukti</th>
                        <th>Aksi</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>
    </div>

    <div class="text-center mt-4">
        <button type="submit" id="saveButton" class="btn btn-primary">Simpan</button>
    </div>
</form>

<script src="../../Js/formulir_UnitTerpasang.js"></script>
<script src="../../Js/formulir_komponen.js"></script>










        <!-- Scripts -->
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.4/jquery.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/js/bootstrap.bundle.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
        <script>
            $(document).ready(function () {
                // Inisialisasi Select2 tanpa fitur hapus pilihan (clear)
                $('.select2').select2({
                    theme: 'bootstrap-5',
                    placeholder: function () {
                        return $(this).data('placeholder');
                    },
                    allowClear: false // Nonaktifkan tombol silang untuk menghapus pilihan
                });
            });
        </script>
</body>

</html>