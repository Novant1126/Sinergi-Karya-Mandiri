<?php
include '../../Koneksi/Koneksi.php';
// Ambil data solusi komponen
$solusi_query = "SELECT * FROM solusi_komponen";
$temuan_query = "SELECT * FROM temuan_komponen";
$solusi_result = $conn->query($solusi_query);
$temuan_result = $conn->query($temuan_query);
?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Master Komponen</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.0.0/dist/tailwind.min.css" rel="stylesheet">
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
                <a href="../../View/Master/Master_Komponen.php" class="text-white hover:text-gray-200">Master Komponen</a>
                <a href="../../View/Master/Master_Solusi.php" class="text-white hover:text-gray-200">Master Solusi Temuan</a>
                <a href="../../View/Audit/Hasil_Audit.php" class="text-white hover:text-gray-200">Hasil Audit</a>
                <a href="../../View/Audit/FORM_AUDIT.php" class="text-white hover:text-gray-200">Formulir Audit</a>
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
        <a href="../../View/Master/Master_Komponen.php" class="block px-4 py-2 hover:bg-gray-100">Master komponen</a>
        <a href="../../View/Master/Master_Solusi.php" class="block px-4 py-2 hover:bg-gray-100">Master Solusi dan Temuan</a>
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
    <div class="row">
        <!-- Kolom untuk Master Solusi -->
        <div class="col-md-6">
            <div class="container-fluid bg-white shadow p-4 rounded">
                <h2 class="text-center fs-5 fw-bold">Master Solusi</h2>
                <button type="button" class="btn btn-primary mb-3" data-bs-toggle="modal" data-bs-target="#addModalSolusi">Tambah Solusi</button>
                <div class="table-responsive overflow-auto rounded-lg shadow">
                    <table class="min-w-full bg-white border border-gray-300 rounded-lg text-sm">
                        <thead class="bg-gray-200 text-gray-700">
                            <tr>
                                <th class="py-3 px-3 text-left border border-gray-300">Code Solusi</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Nama Solusi</th>
                                <th class="py-3 px-3 text-center border border-gray-300">Aksi</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php while ($row = $solusi_result->fetch_assoc()): ?>
                                <tr class="hover:bg-gray-100 border-b border-gray-200">
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['code_solusi']) ?></td>
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['nama_solusi']) ?></td>
                                    <td class="py-3 px-3 text-center">
                                        <button class="btn btn-warning btn-sm" data-bs-toggle="modal"
                                            data-bs-target="#editModalSolusi" data-id="<?= $row['id_solusi'] ?>"
                                            data-code="<?= $row['code_solusi'] ?>" data-name="<?= $row['nama_solusi'] ?>">Edit</button>
                                        <a href="proses_master_solusi.php?delete=<?= $row['id_solusi'] ?>"
                                            class="btn btn-danger btn-sm"
                                            onclick="return confirm('Yakin ingin menghapus data ini?')">Hapus</a>
                                    </td>
                                </tr>
                            <?php endwhile; ?>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Kolom untuk Master Temuan -->
        <div class="col-md-6">
            <div class="container-fluid bg-white shadow p-4 rounded">
                <h2 class="text-center fs-5 fw-bold">Master Temuan</h2>
                <button type="button" class="btn btn-primary mb-3" data-bs-toggle="modal" data-bs-target="#addModalTemuan">Tambah Temuan</button>
                <div class="table-responsive overflow-auto rounded-lg shadow">
                    <table class="min-w-full bg-white border border-gray-300 rounded-lg text-sm">
                        <thead class="bg-gray-200 text-gray-700">
                            <tr>
                                <th class="py-3 px-3 text-left border border-gray-300">Code Temuan</th>
                                <th class="py-3 px-3 text-left border border-gray-300">Nama Temuan</th>
                                <th class="py-3 px-3 text-center border border-gray-300">Aksi</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php while ($row = $temuan_result->fetch_assoc()): ?>
                                <tr class="hover:bg-gray-100 border-b border-gray-200">
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['code_temuan']) ?></td>
                                    <td class="py-3 px-3 text-left"><?= htmlspecialchars($row['nama_temuan']) ?></td>
                                    <td class="py-3 px-3 text-center">
                                        <button class="btn btn-warning btn-sm" data-bs-toggle="modal"
                                            data-bs-target="#editModalTemuan" data-id="<?= $row['id_temuan'] ?>"
                                            data-code="<?= $row['code_temuan'] ?>" data-desc="<?= $row['nama_temuan'] ?>">Edit</button>
                                        <a href="proses_master_temuan.php?delete_temuan=<?= $row['id_temuan'] ?>"
                                            class="btn btn-danger btn-sm"
                                            onclick="return confirm('Yakin ingin menghapus data ini?')">Hapus</a>
                                    </td>
                                </tr>
                            <?php endwhile; ?>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Modal Tambah Solusi -->
<div class="modal fade" id="addModalSolusi" tabindex="-1" aria-labelledby="addModalSolusiLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="addModalSolusiLabel">Tambah Solusi</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <form action="proses_master_solusi.php" method="POST">
                    <input type="hidden" name="add" value="1">
                    <div class="mb-3">
                        <label for="code_solusi" class="form-label">Code Solusi</label>
                        <input type="text" name="code_solusi" id="code_solusi" class="form-control" required>
                    </div>
                    <div class="mb-3">
                        <label for="nama_solusi" class="form-label">Nama Solusi</label>
                        <input type="text" name="nama_solusi" id="nama_solusi" class="form-control" required>
                    </div>
                    <button type="submit" class="btn btn-primary">Simpan</button>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- Modal Edit Solusi -->
<div class="modal fade" id="editModalSolusi" tabindex="-1" aria-labelledby="editModalSolusiLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="editModalSolusiLabel">Edit Solusi</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <form action="proses_master_solusi.php" method="POST">
                    <input type="hidden" name="id_solusi" id="edit_id_solusi">
                    <div class="mb-3">
                        <label for="edit_code_solusi" class="form-label">Code Solusi</label>
                        <input type="text" class="form-control" id="edit_code_solusi" name="code_solusi" required>
                    </div>
                    <div class="mb-3">
                        <label for="edit_nama_solusi" class="form-label">Nama Solusi</label>
                        <input type="text" class="form-control" id="edit_nama_solusi" name="nama_solusi" required>
                    </div>
                    <button type="submit" class="btn btn-primary">Simpan Perubahan</button>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- Modal Tambah Temuan -->
<div class="modal fade" id="addModalTemuan" tabindex="-1" aria-labelledby="addModalTemuanLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="addModalTemuanLabel">Tambah Temuan</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <form action="proses_master_temuan.php" method="POST">
                    <input type="hidden" name="add_temuan" value="1">
                    <div class="mb-3">
                        <label for="code_temuan" class="form-label">Code Temuan</label>
                        <input type="text" name="code_temuan" id="code_temuan" class="form-control" required>
                    </div>
                    <div class="mb-3">
                        <label for="nama_temuan" class="form-label">Nama Temuan</label>
                        <input type="text" name="nama_temuan" id="nama_temuan" class="form-control" required>
                    </div>
                    <button type="submit" class="btn btn-primary">Simpan</button>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- Modal Edit Temuan -->
<div class="modal fade" id="editModalTemuan" tabindex="-1" aria-labelledby="editModalTemuanLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="editModalTemuanLabel">Edit Temuan</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <form action="proses_master_temuan.php" method="POST">
                    <input type="hidden" name="id_temuan" id="edit_id_temuan">
                    <div class="mb-3">
                        <label for="edit_code_temuan" class="form-label">Code Temuan</label>
                        <input type="text" class="form-control" id="edit_code_temuan" name="code_temuan" required>
                    </div>
                    <div class="mb-3">
                        <label for="edit_nama_temuan" class="form-label">Nama Temuan</label>
                        <input type="text" class="form-control" id="edit_nama_temuan" name="nama_temuan" required>
                    </div>
                    <button type="submit" class="btn btn-primary">Simpan Perubahan</button>
                </form>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.bundle.min.js"></script>

<script>
// Script untuk mengisi form edit
var editModalSolusi = document.getElementById('editModalSolusi')
editModalSolusi.addEventListener('show.bs.modal', function (event) {
    var button = event.relatedTarget;
    var id = button.getAttribute('data-id');
    var code = button.getAttribute('data-code');
    var name = button.getAttribute('data-name');

    var modalCodeInput = editModalSolusi.querySelector('#edit_code_solusi');
    var modalNameInput = editModalSolusi.querySelector('#edit_nama_solusi');
    var modalIdInput = editModalSolusi.querySelector('#edit_id_solusi');

    modalCodeInput.value = code;
    modalNameInput.value = name;
    modalIdInput.value = id;
});
</script>

<script>
// Script untuk mengisi form edit
var editModalTemuan = document.getElementById('editModalTemuan')
editModalTemuan.addEventListener('show.bs.modal', function (event) {
    var button = event.relatedTarget;
    var id = button.getAttribute('data-id');
    var code = button.getAttribute('data-code');
    var name = button.getAttribute('data-desc');

    var modalCodeInput = editModalTemuan.querySelector('#edit_code_temuan');
    var modalNameInput = editModalTemuan.querySelector('#edit_nama_temuan');
    var modalIdInput = editModalTemuan.querySelector('#edit_id_temuan');

    modalCodeInput.value = code;
    modalNameInput.value = name;
    modalIdInput.value = id;
});
</script>

</body>
</html>
