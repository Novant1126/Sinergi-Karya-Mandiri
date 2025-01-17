<?php
include '../../Koneksi/Koneksi.php';

// Tentukan jumlah data per halaman
$rowsPerPage = 10; // Atur jumlah data per halaman sesuai kebutuhan Anda

// Ambil halaman saat ini dari parameter URL (default = 1)
$page = isset($_GET['page']) ? (int) $_GET['page'] : 1;
$page = max($page, 1); // Pastikan halaman minimal adalah 1

// Hitung offset
$offset = ($page - 1) * $rowsPerPage;

// Ambil data dari database dengan paginasi
$query = "SELECT * FROM komponen LIMIT $rowsPerPage OFFSET $offset";
$result = $conn->query($query);

// Hitung total data
$totalQuery = "SELECT COUNT(*) AS total_rows FROM komponen";
$totalResult = $conn->query($totalQuery);

if ($totalResult) {
    $row = $totalResult->fetch_assoc();
    $totalRows = $row['total_rows'];
    $totalPages = ceil($totalRows / $rowsPerPage);
} else {
    // Jika query gagal
    $totalRows = 0;
    $totalPages = 0;
}

// Tentukan batas halaman yang ditampilkan
$maxPagesToShow = 5; // Batas jumlah halaman yang akan ditampilkan dalam navigasi
$startPage = max(1, $page - floor($maxPagesToShow / 2));
$endPage = min($totalPages, $startPage + $maxPagesToShow - 1);

?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Master Komponen</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.0.0/dist/tailwind.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css" rel="stylesheet">
</head>

<body class="bg-light">

<nav class="bg-blue-600 shadow-md">
    <div class="container mx-auto px-4 py-3 flex justify-between items-center">
        <!-- Logo dan Menu Desktop -->
        <div class="flex items-center">
            <!-- Logo -->
            <div class="mr-6">
                <img 
                    src="https://tailwindui.com/plus/img/logos/mark.svg?color=indigo&shade=600" 
                    alt="Your Company" 
                    class="h-8 w-auto"
                />
            </div>

            <!-- Menu Desktop -->
            <div class="hidden md:flex space-x-4">
                <a href="../../Index/Master/Master_Komponen.php" class="text-white hover:text-gray-200">Master komponen</a>
                <a href="../../Index/Master/Master_Solusi.php" class="text-white hover:text-gray-200">Master Solusi dan Temuan</a>
                <a href="../../Index/Project/Identitas_Gedung.php" class="text-white hover:text-gray-200">Hasil Audit</a>
                <a href="../../Index/Project/Formulir_Audit.php" class="text-white hover:text-gray-200">Formulir Audit</a>
            </div>
        </div>

        <!-- Hamburger Menu (Mobile) -->
        <div class="md:hidden">
            <button 
                class="text-white focus:outline-none" 
                onclick="toggleMobileMenu()"
                aria-label="Open main menu"
            >
                <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5"></path>
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
        <a href="../../Index/Master/Master_Komponen.php" class="block px-4 py-2 hover:bg-gray-100">Master komponen</a>
        <a href="../../Index/Master/Master_Solusi.php" class="block px-4 py-2 hover:bg-gray-100">Master Solusi dan Temuan</a>
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
    <script src="../../Js/Js.js"></script>
    <div class="row p-5">
        <!-- Kolom untuk Master Komponen Lift -->
        <div class="col-12 col-md-6 mb-4">
                <div class="d-flex justify-content-between align-items-center mb-4">
                    <h1 class="fs-4 fw-bold">Master Komponen Lift</h1>
                    <!-- Tombol Tambah -->
                    <button type="button" class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#addModal">Tambah Komponen</button>
                </div>

                <!-- Tabel Data -->
                <div class="table-responsive overflow-auto rounded-lg shadow">
                    <table class="min-w-full bg-white border border-gray-300 rounded-lg text-sm">
                        <thead class="bg-gray-200 text-gray-700">
                            <tr>
                                <th class="py-3 px-3 text-left border border-gray-300">No</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Code Komponen</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Nama Komponen</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Keterangan</th>
                                <th class="py-3 px-3 text-center border border-gray-300">Aksi</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php
                            $no = 1; // Inisialisasi nomor
                            while ($row = $result->fetch_assoc()): ?>
                                <tr class="hover:bg-gray-100 border-b border-gray-200">
                                    <td class="py-3 px-3 text-left"><?= $no++ ?></td>
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['code_komponen']) ?></td>
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['nama_komponen']) ?></td>
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['keterangan']) ?></td>
                                    <td class="py-3 px-3 text-center d-flex justify-content-between">
                                        <!-- Tombol Edit dengan ikon -->
                                        <button class="btn btn-warning btn-sm" data-bs-toggle="modal"
                                                data-bs-target="#editModal" data-id="<?= $row['id_komponen'] ?>"
                                                data-code="<?= $row['code_komponen'] ?>"
                                                data-name="<?= $row['nama_komponen'] ?>"
                                                data-keterangan="<?= $row['keterangan'] ?>">
                                            <i class="fas fa-edit"></i> <!-- Ikon Edit -->
                                        </button>
                                        
                                        <!-- Tombol Hapus dengan ikon -->
                                        <a href="proses_master_komponen.php?delete=<?= $row['id_komponen'] ?>"
                                        class="btn btn-danger btn-sm ms-1"
                                        onclick="return confirm('Yakin ingin menghapus data ini?')">
                                            <i class="fas fa-trash-alt"></i> <!-- Ikon Hapus -->
                                        </a>
                                    </td>
                                </tr>
                            <?php endwhile; ?>
                        </tbody>
                    </table>
                </div>
                <!-- Paginasi -->
                <nav class="mt-4">
                    <ul class="pagination justify-content-center pagination-sm">
                        <li class="page-item <?= $page == 1 ? 'disabled' : '' ?>">
                            <a class="page-link" href="?page=<?= $page - 1 ?>">Previous</a>
                        </li>

                        <!-- Tampilkan hanya halaman yang relevan -->
                        <?php for ($i = $startPage; $i <= $endPage; $i++): ?>
                            <li class="page-item <?= $i == $page ? 'active' : '' ?>">
                                <a class="page-link" href="?page=<?= $i ?>"><?= $i ?></a>
                            </li>
                        <?php endfor; ?>

                        <li class="page-item <?= $page == $totalPages ? 'disabled' : '' ?>">
                            <a class="page-link" href="?page=<?= $page + 1 ?>">Next</a>
                        </li>
                    </ul>
                </nav>
        </div>

        <!-- Kolom untuk Master Komponen Eskalator -->
        <div class="col-12 col-md-6 mb-4">
            <div class="d-flex justify-content-between align-items-center mb-4">
                <h1 class="fs-4 fw-bold">Master Komponen Eskalator</h1>
                <!-- Tombol Tambah -->
                <button type="button" class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#addModalEskalator">Tambah Komponen</button>
            </div>
                <!-- Tabel Data -->
                <div class="table-responsive overflow-auto rounded-lg shadow">
                    <table class="min-w-full bg-white border border-gray-300 rounded-lg text-sm">
                        <thead class="bg-gray-200 text-gray-700">
                            <tr>
                                <th class="py-3 px-3 text-left border border-gray-300">No</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Code Komponen</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Nama Komponen</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Keterangan</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Aksi</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php
                            $no = 1;
                            $result = $conn->query("SELECT * FROM komponen_eskalator");

                            while ($row = $result->fetch_assoc()): ?>
                                <tr class="hover:bg-gray-100 border-b border-gray-200">
                                    <td class="py-3 px-3 text-left"><?= $no++ ?></td>
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['code_eskalator']) ?></td>
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['nama_eskalator']) ?></td>
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['keterangan_eskalator']) ?></td>
                                    <td class="py-3 px-3 text-center d-flex justify-content-between">
                                        <!-- Tombol Edit dengan ikon -->
                                        <button class="btn btn-warning btn-sm" data-bs-toggle="modal"
                                                data-bs-target="#editModalEskalator" data-id="<?= $row['id_eskalator'] ?>"
                                                data-code="<?= $row['code_eskalator'] ?>"
                                                data-name="<?= $row['nama_eskalator'] ?>"
                                                data-keterangan="<?= $row['keterangan_eskalator'] ?>">
                                            <i class="fas fa-edit"></i> <!-- Ikon Edit -->
                                        </button>

                                        <!-- Tombol Hapus dengan ikon, menambahkan margin kiri untuk jarak -->
                                        <a href="proses_master_komponen.php?delete=<?= $row['id_eskalator'] ?>"
                                        class="btn btn-danger btn-sm ms-1"
                                        onclick="return confirm('Yakin ingin menghapus data ini?')">
                                            <i class="fas fa-trash-alt"></i> <!-- Ikon Hapus -->
                                        </a>
                                    </td>

                                </tr>
                            <?php endwhile; ?>
                        </tbody>
                    </table>
                </div>

                <!-- Paginasi -->
                <nav class="mt-4">
                    <ul class="pagination justify-content-center pagination-sm">
                        <li class="page-item <?= $page == 1 ? 'disabled' : '' ?>">
                            <a class="page-link" href="?page=<?= $page - 1 ?>">Previous</a>
                        </li>
                        <?php for ($i = 1; $i <= $totalPages; $i++): ?>
                            <li class="page-item <?= $i == $page ? 'active' : '' ?>">
                                <a class="page-link" href="?page=<?= $i ?>"><?= $i ?></a>
                            </li>
                        <?php endfor; ?>
                        <li class="page-item <?= $page == $totalPages ? 'disabled' : '' ?>">
                            <a class="page-link" href="?page=<?= $page + 1 ?>">Next</a>
                        </li>
                    </ul>
                </nav>
            </div>
    </div>


        <!-- Modal Tambah -->
        <div class="modal fade" id="addModal" tabindex="-1" aria-labelledby="addModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="addModalLabel">Tambah Komponen</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <form action="proses_master_komponen.php?type=komponen" method="POST">
                            <input type="hidden" name="add" value="1">
                            <div class="mb-3">
                                <label for="code_komponen" class="form-label">Code Komponen</label>
                                <input type="text" name="code_komponen" id="code_komponen" class="form-control"
                                    required>
                            </div>
                            <div class="mb-3">
                                <label for="nama_komponen" class="form-label">Nama Komponen</label>
                                <input type="text" name="nama_komponen" id="nama_komponen" class="form-control"
                                    required>
                            </div>
                            <div class="mb-3">
                                <label for="keterangan" class="form-label">Keterangan</label>
                                <input name="keterangan" id="keterangan" class="form-control" rows="3"></input>
                            </div>
                            <button type="submit" class="btn btn-primary">Simpan</button>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <!-- Modal Edit -->
        <div class="modal fade" id="editModal" tabindex="-1" aria-labelledby="editModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="editModalLabel">Edit Komponen</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <form action="proses_master_komponen.php?type=komponen" method="POST">
                            <input type="hidden" name="id_komponen" id="edit_id">
                            <div class="mb-3">
                                <label for="edit_code_komponen" class="form-label">Code Komponen</label>
                                <input type="text" class="form-control" id="edit_code_komponen" name="code_komponen"
                                    required>
                            </div>
                            <div class="mb-3">
                                <label for="edit_nama_komponen" class="form-label">Nama Komponen</label>
                                <input type="text" class="form-control" id="edit_nama_komponen" name="nama_komponen"
                                    required>
                            </div>
                            <div class="mb-3">
                                <label for="edit_keterangan" class="form-label">Keterangan</label>
                                <textarea class="form-control" id="edit_keterangan" name="keterangan"
                                    rows="3"></textarea>
                            </div>
                            <button type="submit" name="edit" class="btn btn-primary">Simpan Perubahan</button>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <!-- Modal Tambah -->
        <div class="modal fade" id="addModalEskalator" tabindex="-1" aria-labelledby="addModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <form action="proses_master_komponen.php" method="post">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h5 class="modal-title" id="addModalLabel">Tambah Komponen</h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <div class="mb-3">
                                <label for="code_eskalator" class="form-label">Code Komponen</label>
                                <input type="text" class="form-control" id="code_eskalator" name="code_eskalator"
                                    required>
                            </div>
                            <div class="mb-3">
                                <label for="nama_eskalator" class="form-label">Nama Komponen</label>
                                <input type="text" class="form-control" id="nama_eskalator" name="nama_eskalator"
                                    required>
                            </div>
                            <div class="mb-3">
                                <label for="keterangan_eskalator" class="form-label">Keterangan</label>
                                <input type="text" class="form-control" id="keterangan_eskalator"
                                    name="keterangan_eskalator">
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button type="submit" name="add" class="btn btn-primary">Simpan</button>
                        </div>
                    </div>
                </form>
            </div>
        </div>

        <!-- Modal Edit -->
        <div class="modal fade" id="editModalEskalator" tabindex="-1" aria-labelledby="editModalLabel"
            aria-hidden="true">
            <div class="modal-dialog">
                <form action="proses_master_komponen.php" method="post">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h5 class="modal-title" id="editModalLabel">Edit Komponen</h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <input type="hidden" id="edit_id" name="id_eskalator">
                            <div class="mb-3">
                                <label for="edit_code" class="form-label">Code Komponen</label>
                                <input type="text" class="form-control" id="edit_code" name="code_eskalator" required>
                            </div>
                            <div class="mb-3">
                                <label for="edit_name" class="form-label">Nama Komponen</label>
                                <input type="text" class="form-control" id="edit_name" name="nama_eskalator" required>
                            </div>
                            <div class="mb-3">
                                <label for="edit_keterangan" class="form-label">Keterangan</label>
                                <input type="text" class="form-control" id="edit_keterangan"
                                    name="keterangan_eskalator">
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button type="submit" name="update" class="btn btn-warning">Update</button>
                        </div>
                    </div>
                </form>
            </div>
        </div>


        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.min.js"></script>
        <script>
            document.addEventListener('DOMContentLoaded', function () {
                var editModal = document.getElementById('editModal');
                editModal.addEventListener('show.bs.modal', function (event) {
                    // Tombol yang diklik
                    var button = event.relatedTarget;

                    // Ambil data dari tombol
                    var id = button.getAttribute('data-id');
                    var code = button.getAttribute('data-code');
                    var name = button.getAttribute('data-name');
                    var keterangan = button.getAttribute('data-keterangan');

                    // Isi data ke dalam form modal
                    document.getElementById('edit_id').value = id;
                    document.getElementById('edit_code_komponen').value = code;
                    document.getElementById('edit_nama_komponen').value = name;
                    document.getElementById('edit_keterangan').value = keterangan;
                });
            });
        </script>
</body>

</html>