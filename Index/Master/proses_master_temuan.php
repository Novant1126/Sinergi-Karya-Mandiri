<?php
include '../../Koneksi/Koneksi.php';

// Proses Tambah untuk Temuan
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['add_temuan'])) {
    $code_temuan = trim($_POST['code_temuan']);
    $nama_temuan = trim($_POST['nama_temuan']);

    if (!empty($code_temuan) && !empty($nama_temuan)) {
        $sql = "INSERT INTO temuan_komponen (code_temuan, nama_temuan) VALUES (?, ?)";
        $stmt = $conn->prepare($sql);

        if ($stmt) {
            $stmt->bind_param("ss", $code_temuan, $nama_temuan);

            if ($stmt->execute()) {
                header("Location: Master_Solusi.php?status=success_temuan");
                exit;
            } else {
                echo "Error: " . $stmt->error;
            }
        } else {
            echo "Error: " . $conn->error;
        }
    } else {
        echo "Error: Data tidak lengkap.";
    }
}

// Proses Edit untuk Temuan
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['edit_temuan'])) {
    $id_temuan = intval($_POST['id_temuan']);
    $code_temuan = trim($_POST['code_temuan']);
    $nama_temuan = trim($_POST['nama_temuan']);

    if (!empty($id_temuan) && !empty($code_temuan) && !empty($nama_temuan)) {
        $sql = "UPDATE temuan_komponen SET code_temuan = ?, nama_temuan = ? WHERE id_temuan = ?";
        $stmt = $conn->prepare($sql);

        if ($stmt) {
            $stmt->bind_param("ssi", $code_temuan, $nama_temuan, $id_temuan);

            if ($stmt->execute()) {
                header("Location: Master_Solusi.php?status=updated_temuan");
                exit;
            } else {
                echo "Error: " . $stmt->error;
            }
        } else {
            echo "Error: " . $conn->error;
        }
    } else {
        echo "Error: Data tidak lengkap atau ID tidak valid.";
    }
}

// Proses Hapus untuk Temuan
if (isset($_GET['delete_temuan'])) {
    $id = $_GET['delete_temuan'];

    $sql = "DELETE FROM temuan_komponen WHERE id_temuan = ?";
    $stmt = $conn->prepare($sql);
    $stmt->bind_param("i", $id);

    if ($stmt->execute()) {
        header("Location: Master_Solusi.php?status=deleted_temuan");
    } else {
        echo "Error: " . $stmt->error;
    }
}

$conn->close();
?>
