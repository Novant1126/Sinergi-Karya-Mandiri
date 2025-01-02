<?php
include '../../Koneksi/Koneksi.php';

// Fetch data from the gedung table
$sql = "SELECT * FROM gedung_eskalator";
$result = $conn->query($sql);
?>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tabel Gedung</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons/font/bootstrap-icons.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.0.0/dist/tailwind.min.css" rel="stylesheet">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <style>
        @media (max-width: 576px) {
            table {
            font-size: 12px; /* Kurangi ukuran font */
            word-wrap: break-word; /* Pecah kata jika terlalu panjang */
            }

            th, td {
            white-space: nowrap; /* Hindari teks yang memanjang keluar */
            }

            .table-responsive {
            overflow-x: auto; /* Memungkinkan scroll horizontal jika dibutuhkan */
            }


        .modal-dialog {
            max-width: 90%; /* Kurangi ukuran modal */
            margin: 1rem auto; /* Tambahkan jarak di sisi atas dan bawah */
            }

            .modal-content {
            font-size: 14px; /* Kurangi ukuran font untuk menghemat ruang */
            }

            .modal-header, .modal-footer {
            padding: 1rem; /* Kurangi padding di bagian header dan footer */
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
                <img 
                    src="../../IMG/Logo_Sinergi.png" 
                    alt="Your Company" 
                    class="h-8 w-auto"
                />
            </div>

            <!-- Menu Desktop -->
            <div class="hidden md:flex space-x-4">
                <a href="" class="text-white hover:text-gray-200">Master Komponen</a>
                <a href="../Master/Master_Solusi.php" class="text-white hover:text-gray-200">Master Solusi Temuan</a>
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
    
    <!-- Tabel Master Gedung -->
    <div class="container-fluid p-5">
        <h1 class="text-center mb-4 fs-4 fw-bold">Daftar Gedung</h1>
        <button type="button" class="btn btn-primary mb-3 btn-sm" data-bs-toggle="modal" data-bs-target="#tambahModal"><i class="bi bi-plus-circle"></i> Tambah</button>
        <div class="table-responsive">
            <table class="table table-sm table-bordered text-center">
                <thead class="table-info ">
                    <tr>
                        <th scope="col" class="py-2 px-2 small">Isi Form</th>
                        <th scope="col" class="py-2 px-2 small">Nama Gedung</th>
                        <th scope="col" class="py-2 px-2 small">Kode Proyek</th>
                        <th scope="col" class="py-2 px-2 small">Alamat</th>
                        <th scope="col" class="py-2 px-2 text-center small">Aksi</th>
                    </tr>
                </thead>
                <tbody>
                    <?php
                    if ($result->num_rows > 0) {
                        while ($row = $result->fetch_assoc()) {
                            echo "<tr>";
                            echo "<td class='py-2 px-2 text-center small'>";
                            echo "<button class='btn btn-warning btn-sm me-1 mb-1 mb-md-0' data-bs-toggle='modal' data-bs-target='#ubahModal' onclick='setModalData(" . json_encode($row) . ")'>";
                            echo "<i class='bi bi-pencil-square'></i>"; // Ikon edit
                            echo "</button>";                            
                            echo "<button class='btn btn-danger btn-sm ms-1 mb-1 mb-md-0' data-bs-toggle='modal' data-bs-target='#hapusModal' onclick='setHapusId(" . $row['id_gedung'] . ")'>";
                            echo "<i class='bi bi-trash'></i>"; // Ikon hapus
                            echo "</button>";                                                  
                            echo "</td>";               
                            echo "<td class='py-2 px-2 small'>" . htmlspecialchars($row['nama_gedung']) . "</td>";
                            echo "<td class='py-2 px-2 small'>" . htmlspecialchars($row['project_code']) . "</td>";
                            echo "<td class='py-2 px-2 small'>" . htmlspecialchars($row['address']) . "</td>";  

                            echo "<td class='py-2 px-2 text-center small'>";
                            echo "<a href='Hasil_Audit_Tower.php?id_gedung=" . urlencode($row['id_gedung']) . "' class='btn btn-success btn-sm'>";
                            echo "<i class='bi bi-arrow-right-circle'></i> Isi Data Tower";
                            echo "</a>";
                            echo "</td>";                          
                            echo "</tr>";
                        }
                    } else {
                        echo "<tr><td colspan='5' class='text-center small'>Tidak ada data.</td></tr>";
                    }
                    ?>
                </tbody>
            </table>
        </div>
    </div>


    <!-- Modal Tambah -->
    <div class="modal fade" id="tambahModal" tabindex="-1" aria-labelledby="tambahModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header">
                    <h3 class="modal-title fw-bold" id="tambahModalLabel">Tambah Gedung</h3>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <form action="Proses/proses_gedung.php" method="POST">
                        <input type="hidden" name="action" value="add">
                        <div class="mb-3">
                            <label for="nama_gedung" class="form-label">Nama Gedung</label>
                            <input type="text" class="form-control" id="nama_gedung" name="nama_gedung" required>
                        </div>
                        <div class="mb-3">
                            <label for="project_code" class="form-label">Kode Proyek</label>
                            <input type="text" class="form-control" id="project_code" name="project_code" required>
                        </div>
                        <div class="mb-3">
                            <label for="address" class="form-label">Alamat</label>
                            <textarea class="form-control" id="address" name="address" required></textarea>
                        </div>
                        <button type="submit" class="btn btn-primary">Simpan</button>
                    </form>
                </div>
            </div>
        </div>
    </div>


    <!-- Modal Ubah -->
    <div class="modal fade" id="ubahModal" tabindex="-1" aria-labelledby="ubahModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header">
                    <h3 class="modal-title fw-bold" id="ubahModalLabel">Ubah Gedung</h3>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <form action="Proses/proses_gedung.php" method="POST">
                        <input type="hidden" name="action" value="edit">
                        <input type="hidden" name="id_gedung" id="modal-id_gedung">
                        <div class="mb-3">
                            <label for="modal-nama_gedung" class="form-label">Nama Gedung</label>
                            <input type="text" class="form-control" id="modal-nama_gedung" name="nama_gedung" required>
                        </div>
                        <div class="mb-3">
                            <label for="modal-project_code" class="form-label">Kode Proyek</label>
                            <input type="text" class="form-control" id="modal-project_code" name="project_code" required>
                        </div>
                        <div class="mb-3">
                            <label for="modal-address" class="form-label">Alamat</label>
                            <textarea class="form-control" id="modal-address" name="address" required></textarea>
                        </div>
                        <button type="submit" class="btn btn-primary">Update</button>
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
                    <p class="mt-3 fw-bold">Data gedung berhasil dihapus!</p>
                    <form action="Proses/proses_gedung.php" method="POST">
                        <input type="hidden" name="action" value="delete">
                        <input type="hidden" name="id_gedung" id="hapus-id_gedung">
                        <button type="submit" class="btn btn-danger mt-3">Tutup</button>
                    </form>
                </div>
            </div>
        </div>
    </div>



    <script>
        // Scrip Edit Data Gedung
        function setModalData(rowData) {
            document.getElementById('modal-id_gedung').value = rowData.id_gedung;
            document.getElementById('modal-nama_gedung').value = rowData.nama_gedung;
            document.getElementById('modal-project_code').value = rowData.project_code;
            document.getElementById('modal-address').value = rowData.address;
        }
        // Scrip Hapus Data Gedung
        function setHapusId(idGedung) {
            document.getElementById('hapus-id_gedung').value = idGedung;
        }
    </script>

    

    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.9.3/dist/umd/popper.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.min.js"></script>
</body>

</html>