<?php
include '../../Koneksi/Koneksi.php';

// Proses Tambah untuk Solusi
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['add'])) {
    $code_solusi = trim($_POST['code_solusi']);
    $nama_solusi = trim($_POST['nama_solusi']);

    if (!empty($code_solusi) && !empty($nama_solusi)) {
        $sql = "INSERT INTO solusi_komponen (code_solusi, nama_solusi) VALUES (?, ?)";
        $stmt = $conn->prepare($sql);

        if ($stmt) {
            $stmt->bind_param("ss", $code_solusi, $nama_solusi);

            if ($stmt->execute()) {
                header("Location: Master_Solusi.php?status=success");
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

// Proses Edit untuk Solusi
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['edit'])) {
    $id_solusi = intval($_POST['id_solusi']);
    $code_solusi = trim($_POST['code_solusi']);
    $nama_solusi = trim($_POST['nama_solusi']);

    if (!empty($id_solusi) && !empty($code_solusi) && !empty($nama_solusi)) {
        $sql = "UPDATE solusi_komponen SET code_solusi = ?, nama_solusi = ? WHERE id_solusi = ?";
        $stmt = $conn->prepare($sql);

        if ($stmt) {
            $stmt->bind_param("ssi", $code_solusi, $nama_solusi, $id_solusi);

            if ($stmt->execute()) {
                header("Location: Master_Solusi.php?status=updated");
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

// Proses Hapus untuk Solusi
if (isset($_GET['delete'])) {
    $id = $_GET['delete'];

    $sql = "DELETE FROM solusi_komponen WHERE id_solusi = ?";
    $stmt = $conn->prepare($sql);
    $stmt->bind_param("i", $id);

    if ($stmt->execute()) {
        header("Location: Master_Solusi.php?status=deleted");
    } else {
        echo "Error: " . $stmt->error;
    }
}

$conn->close();
?>
