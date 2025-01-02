<?php
include '../../../Koneksi/Koneksi.php';

// Ambil data dari form
$id_tower = $_POST['id_tower'];
$id_gedung = $_POST['id_gedung'];
$lift_no = $_POST['lift_no'];  // Array dari nomor lift
$lift_brand = $_POST['lift_brand'];  // Array dari merek lift
$lift_type = $_POST['lift_type'];  // Array dari tipe lift
$foto_instalasi = $_FILES['foto_instalasi'];  // File gambar
$nama_instalasi = isset($_POST['nama_instalasi']) ? $_POST['nama_instalasi'] : [];  // Array nama instalasi
$deskripsi_instalasi = isset($_POST['deskripsi_instalasi']) ? $_POST['deskripsi_instalasi'] : [];  // Array deskripsi instalasi

// Proses data lift
for ($i = 0; $i < count($lift_no); $i++) {
    $no = $lift_no[$i];
    $brand = $lift_brand[$i];
    $type = $lift_type[$i];

    // Insert data lift ke tabel `audit_lift`
    $stmt = $conn->prepare("INSERT INTO audit_lift (id_tower, lift_no, lift_brand, lift_type) VALUES (?, ?, ?, ?)");
    $stmt->bind_param("isss", $id_tower, $no, $brand, $type);
    if ($stmt->execute()) {
        // Ambil ID lift yang baru saja dimasukkan
        $id_lift = $conn->insert_id;

        // Jika data instalasi ada, proses foto dan deskripsi
        if (!empty($foto_instalasi['name'][0])) {
            for ($j = 0; $j < count($foto_instalasi['name']); $j++) {
                $nama = $nama_instalasi[$j] ?? '';  // Gunakan empty string jika nama instalasi kosong
                $deskripsi = $deskripsi_instalasi[$j] ?? '';  // Gunakan empty string jika deskripsi kosong

                // Proses upload foto
                $foto_name = $foto_instalasi['name'][$j];
                $foto_tmp = $foto_instalasi['tmp_name'][$j];
                $foto_path = 'uploads/foto_instalasi/' . basename($foto_name);
                
                // Pindahkan foto ke folder 'uploads'
                if (move_uploaded_file($foto_tmp, $foto_path)) {
                    // Insert data foto instalasi ke tabel `instalations`
                    $stmt = $conn->prepare("INSERT INTO instalations (id_lift, foto_instalasi, nama_instalasi, deskripsi) VALUES (?, ?, ?, ?)");
                    $stmt->bind_param("isss", $id_lift, $foto_path, $nama, $deskripsi);
                    $stmt->execute();
                }
            }
        }
    }
}

// Tutup koneksi database
$conn->close();
header("Location: ../Formulir_Komponen.php?id_lift=" . $id_lift);
exit; // Always call exit after a header redirection.

?>
