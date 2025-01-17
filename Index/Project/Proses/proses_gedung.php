<?php
include '../../../Koneksi/Koneksi.php';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Ambil data dari form
    $action = $_POST['action'] ?? '';
    $idGedung = $_POST['id_gedung'] ?? '';
    $namaGedung = $_POST['nama_gedung'] ?? '';
    $projectCode = $_POST['project_code'] ?? '';
    $address = $_POST['address'] ?? '';
    $fotoGedung = $_FILES['foto_gedung'] ?? null;  // Ambil file gambar dari form
    
    // Tentukan folder tempat foto disimpan
    $uploadDir = 'uploads/foto_gedung/';
    $fotoGedungName = '';
    if (!is_dir($uploadDir)) {
        mkdir($uploadDir, 0755, true); // Buat folder jika belum ada
    }
    
    // Format nama file foto menggunakan nama gedung dan tanggal
    if ($fotoGedung && $fotoGedung['tmp_name']) {
        // Format nama file menggunakan nama gedung dan tanggal
        $timestamp = date('Ymd-His');  // Format tanggal seperti 20250114-153021
        $fotoGedungName = preg_replace('/\s+/', '_', strtolower($namaGedung)) . '-' . $timestamp . '.' . pathinfo($fotoGedung['name'], PATHINFO_EXTENSION);
        $targetFile = $uploadDir . $fotoGedungName;

        // Pindahkan file dari temporary location ke folder tujuan
        if (!move_uploaded_file($fotoGedung['tmp_name'], $targetFile)) {
            // Foto gagal diupload
            header("Location: ../Identitas_Gedung.php?message=upload-error");
            exit();
        }
    }

    // Tambah Data
    if ($action === 'add' && $namaGedung && $projectCode && $address) {
        $query = "INSERT INTO gedung (nama_gedung, project_code, address, foto_gedung) VALUES (?, ?, ?, ?)";
        $stmt = $conn->prepare($query);
        $stmt->bind_param("ssss", $namaGedung, $projectCode, $address, $fotoGedungName);

        if ($stmt->execute()) {
            header("Location: ../Identitas_Gedung.php?message=success");
            exit();
        } else {
            header("Location: ../Identitas_Gedung.php?message=error");
            exit();
        }
    }
    // Ubah Data
    elseif ($action === 'edit' && $idGedung && $namaGedung && $projectCode && $address) {
        $query = "UPDATE gedung SET nama_gedung = ?, project_code = ?, address = ?, foto_gedung = ? WHERE id_gedung = ?";
        $stmt = $conn->prepare($query);
        $stmt->bind_param("ssssi", $namaGedung, $projectCode, $address, $fotoGedungName, $idGedung);

        if ($stmt->execute()) {
            header("Location: ../Identitas_Gedung.php?message=update-success");
            exit();
        } else {
            header("Location: ../Identitas_Gedung.php?message=update-error");
            exit();
        }
    }
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
