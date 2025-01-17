<?php
include '../../../Koneksi/Koneksi.php';

// Fungsi untuk validasi id_gedung
function validateGedung($conn, $id_gedung) {
    $queryCheckGedung = "SELECT COUNT(*) as count FROM gedung WHERE id_gedung = ?";
    $stmt = $conn->prepare($queryCheckGedung);
    $stmt->bind_param("i", $id_gedung);
    $stmt->execute();
    $result = $stmt->get_result()->fetch_assoc();
    return $result['count'] > 0;
}

// Proses untuk menambahkan tower
if (isset($_POST['nama_tower'], $_POST['pic'], $_POST['jumlah_lantai'], $_POST['id_gedung'])) {
    // Mengambil data dari form
    $nama_tower = $_POST['nama_tower'];
    $pic = $_POST['pic'];
    $jumlah_lantai = $_POST['jumlah_lantai'];
    $id_gedung = $_POST['id_gedung'];

    // Validasi id_gedung
    if (!validateGedung($conn, $id_gedung)) {
        die("Error: id_gedung tidak valid. Pastikan id_gedung sesuai dengan data di tabel gedung.");
    }

    // Query untuk memasukkan data tower baru
    $query = "INSERT INTO audit_tower (id_gedung, nama_tower, pic, jumlah_lantai) VALUES (?, ?, ?, ?)";

    // Persiapkan dan eksekusi query
    if ($stmt = $conn->prepare($query)) {
        $stmt->bind_param("isss", $id_gedung, $nama_tower, $pic, $jumlah_lantai);
        if ($stmt->execute()) {
            // Redirect setelah berhasil
            header("Location: ../Hasil_Audit_Tower.php?id_gedung=" . $id_gedung . "&message=Data tower berhasil ditambahkan");
            exit();
        } else {
            echo "Error saat menambahkan data tower: " . $stmt->error;
        }
    } else {
        echo "Error dalam persiapan query: " . $conn->error;
    }
}

// Proses untuk menghapus tower
if (isset($_POST['action']) && $_POST['action'] == 'delete' && isset($_POST['id_tower'])) {
    $id_tower = $_POST['id_tower'];

    // Query untuk menghapus data tower
    $query = "DELETE FROM audit_tower WHERE id_tower = ?";

    // Persiapkan dan eksekusi query
    if ($stmt = $conn->prepare($query)) {
        $stmt->bind_param("i", $id_tower);  // "i" untuk tipe data integer
        if ($stmt->execute()) {
            // Redirect ke halaman detail gedung setelah penghapusan
            header("Location: ../Hasil_Audit_Tower.php?id_gedung=" . $_POST['id_gedung'] . "&message=Data tower berhasil dihapus");
            exit();
        } else {
            echo "Error saat menghapus data tower: " . $stmt->error;
        }
    } else {
        echo "Error dalam persiapan query: " . $conn->error;
    }
}
?>
 // Hapus Data
    elseif ($action === 'delete' && $idGedung) {
        $query = "DELETE FROM gedung WHERE id_gedung = ?";
        $stmt = $conn->prepare($query);
        $stmt->bind_param("i", $idGedung);

        if ($stmt->execute()) {
            header("Location: ../Identitas_Gedung.php?message=delete-success");
            exit();
        } else {
            header("Location: ../Identitas_Gedung.php?message=delete-error");
            exit();
        }
    } 
    // Jika data tidak valid
    else {
        header("Location: ../Identitas_Gedung.php?message=invalid");
        exit();
    }
}