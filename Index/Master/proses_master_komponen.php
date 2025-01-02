<?php
include '../../Koneksi/Koneksi.php';

// Proses tambah komponen
if (isset($_POST['add']) && $_POST['add'] == 1) {
    // Ambil data dari form tambah
    $code_komponen = $_POST['code_komponen'];
    $nama_komponen = $_POST['nama_komponen'];
    $keterangan = $_POST['keterangan'];

    // Query untuk menambahkan data
    $query = "INSERT INTO komponen (code_komponen, nama_komponen, keterangan) VALUES (?, ?, ?)";
    $stmt = $conn->prepare($query);
    $stmt->bind_param("sss", $code_komponen, $nama_komponen, $keterangan);

    if ($stmt->execute()) {
        // Redirect setelah berhasil menambah data
        header("Location: Master_Komponen.php?status=success");
        exit;
    } else {
        echo "Error: " . $stmt->error;
    }

    $stmt->close();
}

// Proses edit komponen
if (isset($_POST['edit'])) {
    $id_komponen = $_POST['id_komponen'];
    $code_komponen = $_POST['code_komponen'];
    $nama_komponen = $_POST['nama_komponen'];
    $keterangan = $_POST['keterangan'];

    // Query untuk memperbarui data
    $query = "UPDATE komponen SET code_komponen = ?, nama_komponen = ?, keterangan = ? WHERE id_komponen = ?";
    $stmt = $conn->prepare($query);
    $stmt->bind_param("sssi", $code_komponen, $nama_komponen, $keterangan, $id_komponen);

    if ($stmt->execute()) {
        // Redirect setelah berhasil mengedit data
        header("Location: Master_Komponen.php?status=updated");
        exit;
    } else {
        echo "Error: " . $stmt->error;
    }

    $stmt->close();
}

// Proses Hapus untuk Solusi
if (isset($_GET['delete'])) {
    $id = $_GET['delete'];

    $sql = "DELETE FROM komponen WHERE id_komponen = ?";
    $stmt = $conn->prepare($sql);
    $stmt->bind_param("i", $id);

    if ($stmt->execute()) {
        header("Location: Master_Komponen.php?status=deleted");
    } else {
        echo "Error: " . $stmt->error;
    }
}

if (isset($_POST['add'])) {
    $code = $_POST['code_eskalator'];
    $name = $_POST['nama_eskalator'];
    $keterangan = $_POST['keterangan_eskalator'];

    $stmt = $conn->prepare("INSERT INTO komponen_eskalator (code_eskalator, nama_eskalator, keterangan_eskalator) VALUES (?, ?, ?)");
    $stmt->bind_param("sss", $code, $name, $keterangan);
    $stmt->execute();

    header("Location: master_komponen.php");
}

// Update Data
if (isset($_POST['update'])) {
    $id = $_POST['id_eskalator'];
    $code = $_POST['code_eskalator'];
    $name = $_POST['nama_eskalator'];
    $keterangan = $_POST['keterangan_eskalator'];

    $stmt = $conn->prepare("UPDATE komponen_eskalator SET code_eskalator = ?, nama_eskalator = ?, keterangan_eskalator = ? WHERE id_eskalator = ?");
    $stmt->bind_param("sssi", $code, $name, $keterangan, $id);
    $stmt->execute();

    header("Location: master_komponen.php");
}

// Hapus Data
if (isset($_GET['delete'])) {
    $id = $_GET['delete'];

    $stmt = $conn->prepare("DELETE FROM komponen_eskalator WHERE id_eskalator = ?");
    $stmt->bind_param("i", $id);
    $stmt->execute();

    header("Location: master_komponen.php");
}

$conn->close();
?>
