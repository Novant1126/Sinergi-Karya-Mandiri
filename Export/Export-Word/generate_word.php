<?php
require_once __DIR__ . '/../../vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

// Koneksi ke database
$host = 'localhost'; // Ganti dengan host database lo
$dbname = 'sinergi'; // Nama database
$username = 'root'; // Username database
$password = ''; // Password database (kosongkan jika tidak ada password)

try {
    // Membuat koneksi ke database
    $pdo = new PDO("mysql:host=$host;dbname=$dbname", $username, $password);
    // Set mode error untuk menangkap masalah
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
} catch (PDOException $e) {
    // Tangkap error jika koneksi gagal
    echo "Koneksi gagal: " . $e->getMessage();
    exit;
}

if (isset($_GET['id_gedung'])) {
    $id_gedung = (int) $_GET['id_gedung'];

    $query = "
    SELECT 
        g.id_gedung,
        g.nama_gedung,
        g.foto_gedung,
        g.created_at AS tanggal_dibuat,
        i.foto_instalasi,
        i.nama_instalasi,
        i.deskripsi AS deskripsi_instalasi,
        ac.id_komponen,
        k.nama_komponen,
        ac.keterangan AS keterangan_komponen,
        ac.foto_bukti,
        ac.prioritas,
        tc.nama_temuan,
        sc.nama_solusi,
        at.nama_tower,
        al.lift_no
    FROM gedung g
    LEFT JOIN audit_tower at ON g.id_gedung = at.id_gedung
    LEFT JOIN audit_lift al ON at.id_tower = al.id_tower
    LEFT JOIN instalations i ON al.id_lift = i.id_lift
    LEFT JOIN audit_komponen ac ON al.id_lift = ac.id_lift
    LEFT JOIN komponen k ON ac.id_komponen = k.id_komponen
    LEFT JOIN temuan_komponen tc ON ac.id_temuan = tc.id_temuan
    LEFT JOIN solusi_komponen sc ON ac.id_solusi = sc.id_solusi
    WHERE g.id_gedung = :id_gedung
    ";

    $params = [':id_gedung' => $id_gedung];

    try {
        $stmt = $pdo->prepare($query);
        $stmt->execute($params);

        if ($stmt->rowCount() > 0) {
            $data = $stmt->fetchAll(PDO::FETCH_ASSOC);
        } else {
            echo "Tidak ada data ditemukan untuk id_gedung: $id_gedung.";
            exit;
        }
    } catch (PDOException $e) {
        echo "Query gagal: " . $e->getMessage();
        exit;
    }
}



// Buat instance PHPWord
$phpWord = new PhpWord();

// Tambahin section baru
$section = $phpWord->addSection();

// Tambahin teks ke section
$section->addText('Hello, ini teks dari PHPWord!');





// bagian untuk nama file
$namaGedung = isset($data[0]['nama_gedung']) ? $data[0]['nama_gedung'] : 'Gedung_Tidak_Diketahui';
$namaGedung = preg_replace('/[^a-zA-Z0-9_]/', '_', $namaGedung);
$tanggalSekarang = date('Ymd');
$namaFile = $namaGedung . '_' . $tanggalSekarang . '.docx';
// Output file Word langsung ke browser
try {
    // Buat writer
    $writer = IOFactory::createWriter($phpWord, 'Word2007');

    // Set headers untuk download
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment;filename="' . $namaFile . '"');
    header('Cache-Control: max-age=0');

    // Output file
    $writer->save('php://output'); // Langsung kirim ke browser
} catch (Exception $e) {
    echo 'Error: ' . $e->getMessage();
}