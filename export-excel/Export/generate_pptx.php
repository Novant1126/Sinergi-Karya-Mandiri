<?php
require_once __DIR__ . '/../../vendor/autoload.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;


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
// Ambil ID gedung, tower, dan lift dari URL
$id_gedung = isset($_GET['id_gedung']) ? $_GET['id_gedung'] : null;
$id_towers = isset($_GET['id_tower']) ? explode(',', $_GET['id_tower']) : [];
$id_lifts = isset($_GET['id_lift']) ? explode(',', $_GET['id_lift']) : [];

// Debug untuk mengecek apakah parameter sudah diterima dengan benar
echo "ID Gedung: $id_gedung<br>";
echo "ID Tower: " . implode(', ', $id_towers) . "<br>";
echo "ID Lift: " . implode(', ', $id_lifts) . "<br>";

// Pastikan variabel sudah terdefinisi dan tidak kosong
if (empty($id_towers) || empty($id_lifts) || !$id_gedung) {
    echo "ID gedung, tower, atau lift belum didefinisikan dengan benar.";
    exit;
}

// Query dengan placeholder dinamis
$query = "
SELECT 
    g.id_gedung,
    g.nama_gedung,
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
AND at.id_tower IN (" . implode(',', array_map(function($i) { return ":id_tower_$i"; }, range(1, count($id_towers)))) . ")
AND al.id_lift IN (" . implode(',', array_map(function($i) { return ":id_lift_$i"; }, range(1, count($id_lifts)))) . ")
GROUP BY g.id_gedung, g.nama_gedung, i.foto_instalasi, i.nama_instalasi, i.deskripsi, ac.id_komponen, k.nama_komponen
";

// Persiapkan array untuk parameter
$params = [':id_gedung' => $id_gedung];

// Bind parameter untuk tower
foreach ($id_towers as $key => $tower) {
    $params[":id_tower_" . ($key + 1)] = $tower;
}

// Bind parameter untuk lift
foreach ($id_lifts as $key => $lift) {
    $params[":id_lift_" . ($key + 1)] = $lift;
}

// Persiapkan dan eksekusi query
$stmt = $pdo->prepare($query);
$stmt->execute($params);

// Ambil hasil data
$result = $stmt->fetchAll(PDO::FETCH_ASSOC);

// Cek hasilnya
if ($result) {
    print_r($result); // Atau manipulasi sesuai kebutuhan
} else {
    echo "Tidak ada data yang ditemukan.";
}


    // Tampilkan data untuk debugging (opsional)
    foreach ($data as $row) {
        // echo "Nama Gedung: " . $row['nama_gedung'] . "<br>";
        // echo "Tanggal Dibuat: " . $row['tanggal_dibuat'] . "<br>";
        echo "Foto Instalasi: " . $row['foto_instalasi'] . "<br>";
        echo "Nama Instalasi: " . $row['nama_instalasi'] . "<br>";
        echo "Deskripsi Instalasi: " . $row['deskripsi_instalasi'] . "<br>";
        // echo "Nama Komponen: " . $row['nama_komponen'] . "<br>";
        // echo "Nama Tower: " . $row['nama_tower'] . "<br>";
        // echo "Nomor Lift: " . $row['lift_no'] . "<br>";
        // echo "Keterangan Komponen: " . $row['keterangan_komponen'] . "<br>";
        // echo "Nama Temuan: " . $row['nama_temuan'] . "<br>";
        // echo "Nama Solusi: " . $row['nama_solusi'] . "<br><hr>";
    }




    // Buat presentasi baru
    $presentation = new PhpPresentation();

    // Set layout ke 16:9
    $presentation->getLayout()->setDocumentLayout(
        DocumentLayout::LAYOUT_SCREEN_16X9
    );

    // Slide 1
    $slide1 = $presentation->getActiveSlide(); // Slide pertama otomatis dibuat

    // Gambar
    $shape1 = $slide1->createDrawingShape();
    $shape1->setName('Full Slide Image 1')
        ->setDescription('Gambar Full Slide 1')
        ->setPath(__DIR__ . '/img/1.png') // Pastikan gambar ada di folder img
        ->setWidth(960) // Sesuaikan lebar (16:9)
        ->setHeight(540) // Sesuaikan tinggi (16:9)
        ->setOffsetX(0)
        ->setOffsetY(0); // Gambar di posisi awal

    // Buat objek RichText untuk menambahkan teks
    $text = $slide1->createRichTextShape();
    $text->setHeight(50)
        ->setWidth(960);

    // Set posisi manual untuk teks
    $text->setOffsetX(400) // Posisi horizontal (dari kiri)
        ->setOffsetY(450); // Posisi vertikal (dari atas)

    // Menambahkan teks langsung
    $text->createTextRun($row['tanggal_dibuat'])->getFont()->setSize(12)->setColor(new Color('FF000000')); // Styling teks

    // Membuat Slide ke-2
    $slide2 = $presentation->createSlide(); // Tambah slide kedua

    // Gambar di slide ke-2
    $shape2 = $slide2->createDrawingShape();
    $shape2->setName('Full Slide Image 2')
        ->setDescription('Gambar Full Slide 2')
        ->setPath(__DIR__ . '/img/2.png') // Pastikan gambar ada di folder img
        ->setWidth(960) // Sesuaikan lebar (16:9)
        ->setHeight(540) // Sesuaikan tinggi (16:9)
        ->setOffsetX(0)
        ->setOffsetY(0); // Gambar di posisi awal

    // Loop untuk menambahkan slide untuk setiap instalasi
// Slide pertama
    $slide3 = $presentation->createSlide(); // Buat slide baru untuk instalasi pertama

    // Gambar di slide ke-3
    $shape3 = $slide3->createDrawingShape();
    $shape3->setName('Full Slide Image 3')
        ->setDescription('Gambar Full Slide 3')
        ->setPath(__DIR__ . '/img/3.png') // Pastikan gambar ada di folder img
        ->setWidth(960) // Sesuaikan lebar (16:9)
        ->setHeight(540) // Sesuaikan tinggi (16:9)
        ->setOffsetX(0)
        ->setOffsetY(0); // Gambar di posisi awal

    // Menambahkan teks di slide ke-3
    $text = $slide3->createRichTextShape();
    $text->setHeight(50)
        ->setWidth(960);

    // Set posisi teks (sesuaikan dengan posisi yang diinginkan)
    $text->setOffsetX(0) // Posisi horizontal (dari kiri)
        ->setOffsetY(35); // Posisi vertikal (dari atas)
    $text->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
    // Menambahkan teks ke dalam slide
    $text->createTextRun('UNIT LIFT TERPASANG')
        ->getFont()
        ->setSize(20) // Ukuran font
        ->setColor(new Color('FF000000')) // Warna font
        ->setName('Arial') // Mengatur font ke Arial
        ->setBold(true); // Membuat font menjadi bold

    // Menambahkan data instalasi ke dalam slide
    foreach ($data as $key => $row) {
        // Setiap 2 data instalasi akan tampil dalam 1 slide
        if ($key % 2 == 0 && $key != 0) {
            // Jika sudah ada 2 data, buat slide baru
            $slide3 = $presentation->createSlide();
            // Gambar di slide ke-3
            $shape3 = $slide3->createDrawingShape();
            $shape3->setName('Full Slide Image 3')
                ->setDescription('Gambar Full Slide 3')
                ->setPath(__DIR__ . '/img/3.png') // Pastikan gambar ada di folder img
                ->setWidth(960) // Sesuaikan lebar (16:9)
                ->setHeight(540) // Sesuaikan tinggi (16:9)
                ->setOffsetX(0)
                ->setOffsetY(0); // Gambar di posisi awal

            // Menambahkan teks di slide ke-3
            $text = $slide3->createRichTextShape();
            $text->setHeight(50)
                ->setWidth(960);

            // Set posisi teks (sesuaikan dengan posisi yang diinginkan)
            $text->setOffsetX(0) // Posisi horizontal (dari kiri)
                ->setOffsetY(35); // Posisi vertikal (dari atas)
            $text->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
            // Menambahkan teks ke dalam slide
            $text->createTextRun('UNIT LIFT TERPASANG')
                ->getFont()
                ->setSize(20) // Ukuran font
                ->setColor(new Color('FF000000')) // Warna font
                ->setName('Arial') // Mengatur font ke Arial
                ->setBold(true); // Membuat font menjadi bold
        }


        $shape1 = $slide3->createDrawingShape();
        $shape1->setName('Installation Image ' . ($key + 1))
            ->setDescription('Gambar Instalasi ' . ($key + 1))
            ->setPath(__DIR__ . '/../Proses/uploads/foto_instalasi/' . $row['foto_instalasi']) // Pastikan path gambar benar
            ->setWidth(480) // Setengah dari slide (16:9)
            ->setHeight(270) // Proporsional
            ->setOffsetX(10 + ($key % 2) * 490) // Posisi gambar secara berdampingan
            ->setOffsetY(10); // Posisi vertikal

        // Teks instalasi
        $text1 = $slide3->createRichTextShape();
        $text1->setHeight(30)
            ->setWidth(350)
            ->setOffsetX(105 + ($key % 2) * 400) // Posisi teks agar sejajar dengan gambar
            ->setOffsetY(440); // Posisi teks di bawah gambar
        $text1->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
        $text1->createTextRun($row['nama_instalasi'])->getFont()->setSize(16)->setBold(true)->setColor(new Color('FF000000'));
        $text2 = $slide3->createRichTextShape();
        $text2->setHeight(30)
            ->setWidth(350)
            ->setOffsetX(105 + ($key % 2) * 400) // Posisi teks sejajar dengan teks pertama
            ->setOffsetY(470); // Posisi teks sedikit lebih bawah
        $text2->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
        $text2->createTextRun($row['deskripsi_instalasi'])
            ->getFont()->setSize(14) // Ukuran font lebih kecil
            ->setColor(new Color('FF000000')); // Warna teks lebih terang untuk perbedaan
    }


    // Mengecek apakah ada data
    if ($data) {
        // Loop untuk setiap temuan
        $slide4 = $presentation->createSlide(); // Buat slide baru untuk instalasi pertama

        // Gambar di slide ke-3
        $shape4 = $slide4->createDrawingShape();
        $shape4->setName('Full Slide Image 3')
            ->setDescription('Gambar Full Slide 3')
            ->setPath(__DIR__ . '/img/4.png') // Pastikan gambar ada di folder img
            ->setWidth(960) // Sesuaikan lebar (16:9)
            ->setHeight(540) // Sesuaikan tinggi (16:9)
            ->setOffsetX(0)
            ->setOffsetY(0); // Gambar di posisi awal

        // Menambahkan teks di slide ke-3
        $text4 = $slide4->createRichTextShape();
        $text4->setHeight(50)
            ->setWidth(960);
        $text4->setOffsetX(0) // Posisi horizontal (dari kiri)
            ->setOffsetY(35); // Posisi vertikal (dari atas)
        $text4->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
        // Menambahkan teks ke dalam slide
        $text4->createTextRun('TEMUAN')
            ->getFont()
            ->setSize(20) // Ukuran font
            ->setColor(new Color('FF000000')) // Warna font
            ->setName('Arial') // Mengatur font ke Arial
            ->setBold(true); // Membuat font menjadi bold

        // Array untuk menyimpan id_komponen yang sudah diproses
        $processed_komponen = []; // Pastikan array ini dideklarasikan di luar loop
        $invalid_ids = [1];
        // Loop untuk menambahkan data ke slide
        foreach ($data as $key => $row) {
            // Cek apakah id_komponen sudah pernah diproses
            if (in_array($row['id_komponen'], $processed_komponen)) {
                // Lewatkan jika id_komponen sudah ada dalam array
                continue;
            }
            // Cek apakah id_temuan atau id_solusi ada dalam array dan memiliki nilai
            if ($row['nama_temuan'] == 'Tidak Ada' || $row['nama_solusi'] == 'Tidak Ada') {
                continue; // Lewatkan data yang memiliki "Tidak Ada"
            }


            // Tambahkan id_komponen ke dalam array untuk menandakan sudah diproses
            $processed_komponen[] = $row['id_komponen'];

            // Setiap 2 data instalasi akan tampil dalam 1 slide
            if ($key % 2 == 0 && $key != 0) {
                // Jika sudah ada 2 data, buat slide baru
                $slide4 = $presentation->createSlide();

                // Gambar di slide
                $shape4 = $slide4->createDrawingShape();
                $shape4->setName('Full Slide Image 3')
                    ->setDescription('Gambar Full Slide 3')
                    ->setPath(__DIR__ . '/img/4.png') // Pastikan gambar ada di folder img
                    ->setWidth(960) // Sesuaikan lebar (16:9)
                    ->setHeight(540) // Sesuaikan tinggi (16:9)
                    ->setOffsetX(0)
                    ->setOffsetY(0); // Gambar di posisi awal

                // Menambahkan teks di slide
                $text4 = $slide4->createRichTextShape();
                $text4->setHeight(50)
                    ->setWidth(960);
                $text4->setOffsetX(0) // Posisi horizontal (dari kiri)
                    ->setOffsetY(35); // Posisi vertikal (dari atas)
                $text4->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
                // Menambahkan teks ke dalam slide
                $text4->createTextRun('TEMUAN')
                    ->getFont()
                    ->setSize(20) // Ukuran font
                    ->setColor(new Color('FF000000')) // Warna font
                    ->setName('Arial') // Mengatur font ke Arial
                    ->setBold(true); // Membuat font menjadi bold
            }

            // Teks instalasi (Komponen : Temuan)
            $text4 = $slide4->createRichTextShape();
            $text4->setHeight(25)
                ->setWidth(400)
                ->setOffsetX(65 + ($key % 2) * 427) // Posisi teks agar sejajar dengan gambar
                ->setOffsetY(408); // Posisi teks di bawah gambar
            $text4->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
            $text4->createTextRun($row['nama_komponen'] . ' : ' . $row['nama_temuan'])->getFont()->setSize(12)->setBold(true)->setColor(new Color('FF000000'));

            // Teks Tower
            $text5 = $slide4->createRichTextShape();
            $text5->setHeight(25)
                ->setWidth(400)
                ->setOffsetX(65 + ($key % 2) * 427) // Posisi teks sejajar dengan teks pertama
                ->setOffsetY(433); // Posisi teks sedikit lebih bawah
            $text5->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
            $text5->createTextRun('Tower: ' . $row['nama_tower'])
                ->getFont()->setSize(12) // Ukuran font lebih kecil
                ->setColor(new Color('FF000000')); // Warna teks lebih terang untuk perbedaan

            // Teks Prioritas
            $text6 = $slide4->createRichTextShape();
            $text6->setHeight(25)
                ->setWidth(400)
                ->setOffsetX(65 + ($key % 2) * 427) // Posisi teks sejajar dengan teks pertama
                ->setOffsetY(458); // Posisi teks sedikit lebih bawah
            $text6->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
            $text6->createTextRun($row['prioritas'])
                ->getFont()->setSize(12) // Ukuran font lebih kecil
                ->setColor(new Color('FF000000')); // Warna teks lebih terang untuk perbedaan

            // Teks Solusi
            $text7 = $slide4->createRichTextShape();
            $text7->setHeight(25)
                ->setWidth(400)
                ->setOffsetX(65 + ($key % 2) * 427) // Posisi teks sejajar dengan teks pertama
                ->setOffsetY(483); // Posisi teks sedikit lebih bawah
            $text7->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
            $text7->createTextRun('Solusi: ' . $row['nama_solusi'])
                ->getFont()->setSize(12) // Ukuran font lebih kecil
                ->setColor(new Color('FF000000')); // Warna teks lebih terang untuk perbedaan
        }

    }


// Pastikan folder 'PPTX' ada dan memiliki izin tulis
$folder = 'PPTX';
if (!is_dir($folder)) {
    mkdir($folder, 0777, true); // Membuat folder jika belum ada
}
if (is_writable($folder)) {
    try {
        // Menyimpan file PPTX ke dalam folder 'PPTX' dengan nama 'Audit.pptx'
        $oWriter = IOFactory::createWriter($presentation, 'PowerPoint2007');
        $oWriter->save($folder . '/Audit-1.pptx');
        echo "File PPTX berhasil disimpan di " . $folder . "/Audit.pptx";
    } catch (Exception $e) {
        // Menangkap error dan menampilkan pesan kesalahan
        echo 'Error: ' . $e->getMessage();
    }
} else {
    echo "Folder '$folder' tidak memiliki izin tulis. Pastikan folder dapat diakses untuk menulis file.";
}
?>