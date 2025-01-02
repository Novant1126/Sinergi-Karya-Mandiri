<?php
include '../../Koneksi/Koneksi.php';

// Pastikan id_gedung ada di URL
if (isset($_GET['id_gedung'])) {
    $id_gedung = $_GET['id_gedung'];

    // Query untuk mengambil nama gedung dan project_code berdasarkan id_gedung
    $gedungQuery = "SELECT nama_gedung, project_code FROM gedung WHERE id_gedung = ?";
    $stmtGedung = $conn->prepare($gedungQuery);
    $stmtGedung->bind_param('i', $id_gedung);
    $stmtGedung->execute();
    $resultGedung = $stmtGedung->get_result();

    // Periksa jika nama gedung dan project_code ditemukan
    if ($resultGedung->num_rows > 0) {
        $rowGedung = $resultGedung->fetch_assoc();
        $namaGedung = $rowGedung['nama_gedung'];
        $projectCode = $rowGedung['project_code'];
    } else {
        $namaGedung = "Nama Gedung Tidak Ditemukan"; // Jika nama gedung tidak ditemukan
        $projectCode = "Project Code Tidak Ditemukan"; // Jika project_code tidak ditemukan
    }

    // Tentukan jumlah data per halaman
    $limit = 10;

    // Menghitung jumlah total data
    $countQuery = "SELECT COUNT(*) as total FROM audit_tower WHERE id_gedung = ?";
    $stmtCount = $conn->prepare($countQuery);
    $stmtCount->bind_param('i', $id_gedung);
    $stmtCount->execute();
    $resultCount = $stmtCount->get_result();
    $rowCount = $resultCount->fetch_assoc();
    $totalData = $rowCount['total'];

    // Hitung total halaman
    $totalPages = ceil($totalData / $limit);

    // Tentukan halaman saat ini
    $page = isset($_GET['page']) ? (int) $_GET['page'] : 1;
    if ($page < 1) {
        $page = 1;
    } elseif ($page > $totalPages) {
        $page = $totalPages;
    }

    // Hitung posisi data yang akan ditampilkan
    $offset = ($page - 1) * $limit;

    // Query untuk mengambil data tower berdasarkan id_gedung dan paging
    $query = "
    SELECT 
        t.id_tower,
        t.nama_tower,
        t.pic,
        t.jumlah_lantai
    FROM 
        audit_tower t
    WHERE 
        t.id_gedung = ?
    LIMIT ?, ? 
    ";

    // Persiapkan dan jalankan query untuk data tower
    $stmt = $conn->prepare($query);
    $stmt->bind_param('iii', $id_gedung, $offset, $limit);
    $stmt->execute();
    $result = $stmt->get_result();
} else {
    echo "ID Gedung tidak ditemukan.";
}
?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hasil Audit Tower</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.0.0/dist/tailwind.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons/font/bootstrap-icons.css" rel="stylesheet">



    <style>
        @media (max-width: 576px) {
            table {
                font-size: 12px;
                /* Ukuran font lebih kecil */
            }

            th,
            td {
                white-space: nowrap;
                /* Hindari teks melampaui sel */
                overflow: hidden;
                /* Sembunyikan teks yang terlalu panjang */
                text-overflow: ellipsis;
                /* Tambahkan elipsis */
            }
        }
    </style>

</head>

<body class="bg-light">
    <nav class="bg-blue-600 shadow-md">
        <div class="container mx-auto px-4 py-3 flex justify-between items-center">
            <!-- Logo dan Menu Desktop -->
            <div class="flex items-center">
                <!-- Logo -->
                <div class="mr-6">
                    <img src="https://tailwindui.com/plus/img/logos/mark.svg?color=indigo&shade=600" alt="Your Company"
                        class="h-8 w-auto" />
                </div>

                <!-- Menu Desktop -->
                <div class="hidden md:flex space-x-4">
                    <a href="../../View/Master/Master_Komponen.php" class="text-white hover:text-gray-200">Master
                        Komponen</a>
                    <a href="../../View/Master/Master_Solusi.php" class="text-white hover:text-gray-200">Master Solusi
                        Temuan</a>
                    <a href="../../View/Audit/Hasil_Audit.php" class="text-white hover:text-gray-200">Hasil Audit</a>
                    <a href="../../View/Audit/FORM_AUDIT.php" class="text-white hover:text-gray-200">Formulir Audit</a>
                </div>
            </div>

            <!-- Hamburger Menu (Mobile) -->
            <div class="md:hidden">
                <button class="text-white focus:outline-none" onclick="toggleMobileMenu()" aria-label="Open main menu">
                    <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24"
                        stroke-width="1.5" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round"
                            d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5"></path>
                    </svg>
                </button>
            </div>
        </div>
    </nav>
    <!-- Mobile Menu -->
    <div id="mobileMenu" class="hidden md:hidden bg-white text-gray-800 px-4 py-3 space-y-2">
        <a href="#" class="block px-4 py-2 hover:bg-gray-100">Dashboard</a>
        <div class="border-t pt-2">
            <p class="text-gray-600 font-semibold">Master</p>
            <a href="../../View/Master/Master_Komponen.php" class="block px-4 py-2 hover:bg-gray-100">Master
                komponen</a>
            <a href="../../View/Master/Master_Solusi.php" class="block px-4 py-2 hover:bg-gray-100">Master Solusi dan
                Temuan</a>
        </div>
        <div class="border-t pt-2">
            <p class="text-gray-600 font-semibold">Audit</p>
            <a href="../../View/Audit/FORM_AUDIT.php" class="block px-4 py-2 hover:bg-gray-100">Fomulir Audit Lift</a>
            <a href="../../View/Audit/Hasil_Audit.php" class="block px-4 py-2 hover:bg-gray-100">Hasil Audit lift</a>
            <a href="#" class="block px-4 py-2 hover:bg-gray-100">Fomulir Audit</a>
            <a href="#" class="block px-4 py-2 hover:bg-gray-100">Hasil Audit</a>
        </div>
        <div class="border-t pt-2">
            <p class="text-gray-600 font-semibold">Calendar</p>
            <a href="#" class="block px-4 py-2 hover:bg-gray-100">Upcoming Events</a>
            <a href="#" class="block px-4 py-2 hover:bg-gray-100">Past Events</a>
            <a href="#" class="block px-4 py-2 hover:bg-gray-100">Event Settings</a>
        </div>
    </div>
    <script src="../Js.js"></script>


    <div class="container-fluid p-5">
        <h1 class="text-center mb-4 fs-4 fw-bold">Hasil Audit Tower</h1>
        <?php if (isset($id_gedung)): ?>
            <h3 class="text-center mb-4 fs-8 fs-sm-7 fs-md-4">Data Tower untuk Gedung: <?= htmlspecialchars($namaGedung) ?>,
                Dengan Code: <?= htmlspecialchars($projectCode) ?></h3>

            <div class="btn-group btn-sm mb-3" role="group" aria-label="Export Buttons">
                <!-- Tombol Export Excel -->
                <button id="exportExcel" class="btn btn-success btn-sm" title="Export to Excel"><i
                        class="fas fa-file-excel"></i></button>
                <!-- Tombol Export Word -->
                <button id="exportWord" class="btn btn-primary btn-sm" title="Export to Word"><i
                        class="fas fa-file-word"></i> </button>
                <!-- Tombol Export PPTX -->
                <button id="PPTX" class="btn btn-warning btn-sm" title="Export to PowerPoint"><i
                        class="fas fa-file-powerpoint"></i></button>
            </div>
            <button type="button" class="btn btn-primary mb-3 btn-sm" data-bs-toggle="modal"
                data-bs-target="#tambahModalTower"><i class="bi bi-plus-circle"></i> Tambah</button>


            <script>
                document.getElementById("exportExcel").addEventListener("click", function () {
                    alert("Export to Excel clicked!");
                    // Tambahkan logika untuk ekspor Excel
                });

                document.getElementById("exportWord").addEventListener("click", function () {
                    alert("Export to Word clicked!");
                    // Tambahkan logika untuk ekspor Word
                });

                document.getElementById("PPTX").addEventListener("click", function () {
                    alert("Export to PPTX clicked!");
                    // Tambahkan logika untuk ekspor PowerPoint
                });
            </script>
            <script src="PPTX.js"></script> <!-- Tambahkan file logika eksternal -->

            <div class="table-responsive">
                <table class="table table-sm table-bordered text-center">
                    <thead class="table-info">
                        <tr>
                            <th scope="col" class="py-2 px-2 small">Aksi</th>
                            <th scope="col" class="py-2 px-2 small">No.</th>
                            <th scope="col" class="py-2 px-2 small">Nama Tower</th>
                            <th scope="col" class="py-2 px-2 small">Inspektor</th>
                            <th scope="col" class="py-2 px-2 small">Jumlah Lantai</th>
                            <th scope="col" class="py-2 px-2 small">Lanjut Audit</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php
                        if ($result->num_rows > 0) {
                            $counter = 1; // Variabel untuk menghitung nomor urut
                            while ($row = $result->fetch_assoc()) {
                                echo '<tr id="row-' . htmlspecialchars($row['id_tower']) . '" class="align-middle">';

                                // Tombol Detail
                                echo '<td class="py-2 px-2 text-center">';
                                echo '<a href="Hasil_Audit.php?id_tower=' . htmlspecialchars($row['id_tower']) . '" class="btn btn-info btn-sm" data-id="' . htmlspecialchars($row['id_tower']) . '" title="Detail">';
                                echo '<i class="bi bi-eye"></i>';
                                echo '</a>';
                                echo "<button class='btn btn-danger btn-sm ms-1' data-bs-toggle='modal' data-bs-target='#hapusModal' onclick='setHapusId(" . $row['id_tower'] . ")'>";
                                echo "<i class='bi bi-trash'></i>"; // Ikon hapus
                                echo "</button>";
                                echo '</td>';

                                // Nomor urut otomatis
                                echo '<td class="py-2 px-2 small">' . $counter++ . '</td>';

                                // Data lain (Nama Tower, PIC, Jumlah Lantai)
                                echo '<td class="py-2 px-2 small">' . htmlspecialchars($row['nama_tower']) . '</td>';
                                echo '<td class="py-2 px-2 small">' . htmlspecialchars($row['pic']) . '</td>';
                                echo '<td class="py-2 px-2 small">' . htmlspecialchars($row['jumlah_lantai']) . '</td>';

                                // Tombol untuk Halaman Selanjutnya
                                echo '<td class="py-2 px-2 text-center">';
                                echo '<a href="Formulir_Audit.php?id_tower=' . htmlspecialchars($row['id_tower']) . '" class="btn btn-success btn-sm">';
                                echo '<i class="bi bi-arrow-right-circle"></i> Isi Data Lain';
                                echo '</a>';
                                echo '</td>';

                                echo '</tr>';
                            }
                        } else {
                            echo '<tr>';
                            echo '<td colspan="7" class="text-center py-2 px-2 small">Tidak ada data tersedia.</td>';
                            echo '</tr>';
                        }
                        ?>
                    </tbody>
                </table>
            </div>
        <?php endif; ?>
        <!-- Pagination -->
        <nav class="mt-2">
            <ul class="pagination pagination-sm justify-content-center">
                <!-- Tombol Previous -->
                <li class="page-item <?= $page == 1 ? 'disabled' : '' ?>">
                    <a class="page-link" href="?id_tower=<?= $id_tower ?>&page=<?= $page - 1 ?>">Previous</a>
                </li>

                <!-- Nomor Halaman -->
                <?php for ($i = 1; $i <= $totalPages; $i++): ?>
                    <li class="page-item <?= $i == $page ? 'active' : '' ?>">
                        <a class="page-link" href="?id_tower=<?= $id_tower ?>&page=<?= $i ?>"><?= $i ?></a>
                    </li>
                <?php endfor; ?>

                <!-- Tombol Next -->
                <li class="page-item <?= $page == $totalPages ? 'disabled' : '' ?>">
                    <a class="page-link" href="?id_tower=<?= $id_tower ?>&page=<?= $page + 1 ?>">Next</a>
                </li>
            </ul>
        </nav>


        <!-- Modal Tambah Tower -->
        <div class="modal fade" id="tambahModalTower" tabindex="-1" aria-labelledby="tambahModalTowerLabel"
            aria-hidden="true">
            <div class="modal-dialog modal-lg modal-dialog-centered">
                <!-- Gunakan modal-lg untuk PC dan modal-dialog-centered untuk menengah -->
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="tambahModalTowerLabel">Tambah Data Tower</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <!-- Form untuk tambah data tower -->
                        <form action="Proses/proses_tower.php" method="POST">
                            <!-- Input Nama Tower -->
                            <div class="mb-3">
                                <label for="nama_tower" class="form-label">Nama Tower</label>
                                <input type="text" class="form-control" id="nama_tower" name="nama_tower" required>
                            </div>

                            <!-- Input PIC -->
                            <div class="mb-3">
                                <label for="pic" class="form-label">PIC</label>
                                <input type="text" class="form-control" id="pic" name="pic" required>
                            </div>

                            <!-- Input Jumlah Lantai -->
                            <div class="mb-3">
                                <label for="jumlah_lantai" class="form-label">Jumlah Lantai</label>
                                <input type="number" class="form-control" id="jumlah_lantai" name="jumlah_lantai"
                                    required>
                            </div>

                            <!-- Menambahkan ID Gedung yang dikirim melalui URL -->
                            <input type="hidden" name="id_gedung" value="<?= htmlspecialchars($_GET['id_gedung']) ?>">

                            <div class="modal-footer">
                                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Batal</button>
                                <button type="submit" class="btn btn-primary">Simpan</button>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
        </div>
        <!-- Modal Informasi Hapus -->
        <div class="modal fade" id="hapusModal" tabindex="-1" aria-labelledby="hapusModalLabel" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content">
                    <div class="modal-body text-center">
                        <i class="bi bi-exclamation-triangle-fill text-danger fs-1"></i>
                        <p class="mt-3 fw-bold">Apakah Anda yakin ingin menghapus data tower ini?</p>
                        <form action="Proses/proses_tower.php" method="POST">
                        <input type="hidden" name="id_gedung" value="<?= htmlspecialchars($_GET['id_gedung']) ?>">
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id_tower" id="hapus-id_tower">
                            <button type="button" class="btn btn-secondary mt-3" data-bs-dismiss="modal">Batal</button>
                            <button type="submit" class="btn btn-danger mt-3">Hapus</button>
                        </form>
                    </div>
                </div>
            </div>
        </div>





    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script>
        $(document).ready(function () {
            // Event klik pada tombol detail
            $(document).on('click', '.btn-detail', function (e) {
                e.preventDefault();
                const id_tower = $(this).data('id');
                // Mengirimkan id_tower sebagai parameter di URL
                window.location.href = `Hasil_Audit_Komponen.php?id_tower=${id_tower}`;
            });
        });

    
        function setHapusId(id) {
            document.getElementById('hapus-id_tower').value = id;
        }
    </script>
    
</body>

</html>