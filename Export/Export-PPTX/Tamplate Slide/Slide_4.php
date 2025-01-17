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

// Buat presentasi baru
$presentation = new PhpPresentation();

// Set layout ke 16:9
$presentation->getLayout()->setDocumentLayout(
    DocumentLayout::LAYOUT_SCREEN_16X9
);

// slide pertama
    // Deklarasikan fungsi addTextWithOutline di luar loop
    function addTextWithOutline($slide, $text, $x, $y, $width, $height, $fontSize, $textColor, $outlineColor) {
        // Tambahkan teks outline (lapisan luar) dengan posisi sedikit bergeser
        $offsets = [
            [-1, 0], [1, 0], [0, -1], [0, 1], // Atas, bawah, kiri, kanan
            [-1, -1], [1, -1], [-1, 1], [1, 1], // Sudut
        ];
        foreach ($offsets as $offset) {
            $txtOutline = $slide->createRichTextShape();
            $txtOutline->setHeight($height)
                ->setWidth($width)
                ->setOffsetX($x + $offset[0])
                ->setOffsetY($y + $offset[1]);
            $txtOutline->getParagraph(0)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $txtOutline->createTextRun($text)
                ->getFont()->setSize($fontSize)
                ->setColor(new Color($outlineColor)) // Warna outline
                ->setBold(true);
        }
        // Tambahkan teks utama di tengah
        $txtMain = $slide->createRichTextShape();
        $txtMain->setHeight($height)
            ->setWidth($width)
            ->setOffsetX($x)
            ->setOffsetY($y);
        $txtMain->getParagraph(0)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $txtMain->createTextRun($text)
            ->getFont()->setSize($fontSize)
            ->setColor(new Color($textColor)) // Warna teks utama
            ->setBold(true);
    }

    // Pastikan slide aktif sudah ada
    $slide1 = $presentation->getActiveSlide(); 

    // Gambar pertama
    $shape1 = $slide1->createDrawingShape();
    $shape1->setName('Full Slide Image 1')
        ->setDescription('Gambar Full Slide 1')
        ->setPath(__DIR__ . '/img/1.png')
        ->setWidth(960)
        ->setHeight(540)
        ->setOffsetX(0)
        ->setOffsetY(0);

    // Ambil data dari hasil query
    if (isset($data)) {
        foreach ($data as $row) {
            // Ambil nama file gambar dari database
            $fotoGedung = $row['foto_gedung'];
            $imagePath = __DIR__ . '/../../Index/Project/Proses/uploads/foto_gedung/' . $fotoGedung;

            $slideWidth = 960;  // Lebar slide (misalnya untuk 16:9, 960px)
            $slideHeight = 540; // Tinggi slide (misalnya untuk 16:9, 540px)
            $width = 760;   // Lebar gambar dalam piksel
            $height = 340;  // Tinggi gambar dalam piksel

            // Menghitung posisi otomatis agar gambar berada di tengah
            $offsetX = ($slideWidth - $width) / 2;
            $offsetY = ($slideHeight - $height) / 2;
            // Cek apakah gambar perlu sedikit geser ke kanan
            $offsetX += 250;  // Menambahkan sedikit offset agar gambar agak ke kanan

            // Pastikan gambar ada di folder yang sesuai
            if (file_exists($imagePath)) {
                // Gambar
                $shape1 = $slide1->createDrawingShape();
                $shape1->setName('Full Slide Image 1')
                    ->setDescription('Gambar Full Slide 1')
                    ->setPath($imagePath) // Menggunakan path dinamis dari DB
                    ->setWidth($width)    // Menggunakan lebar yang sudah di-set
                    ->setHeight($height)  // Menggunakan tinggi yang sudah di-set
                    ->setOffsetX($offsetX)       // Posisi gambar horizontal di tengah
                    ->setOffsetY($offsetY);      // Posisi gambar vertikal di tengah
            } else {
                echo "Gambar tidak ditemukan: " . $imagePath;
                exit;
            }

            // Gunakan fungsi untuk menambahkan teks dengan outline
            addTextWithOutline(
                $slide1,                      // Slide target
                $row['nama_gedung'],          // Teks yang akan ditampilkan
                305, 360,                     // Posisi X, Y
                340, 70,                      // Lebar dan tinggi teks
                22,                           // Ukuran font
                'FFFF00',                     // Warna teks utama (kuning)
                'FF000000'                    // Warna outline (hitam)
            );

            // Tambahkan teks di atas gambar
            $txtsled2 = $slide1->createRichTextShape();
            $txtsled2->setHeight(25)
                ->setWidth(960)
                ->setOffsetX(0) // Offset X di-set ke 0 agar di tengah horizontal
                ->setOffsetY(450); // Posisi teks secara vertikal (di atas gambar)

            // Set alignment teks ke tengah
            $txtsled2->getParagraph(0)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

            // Menambahkan teks dan format font
            $txtsled2->createTextRun($row['tanggal_dibuat'])
                ->getFont()->setSize(12)
                ->setColor(new Color('FF000000'));
        }
    } else {
        echo "Data tidak ditemukan atau format data tidak sesuai.";
        exit;
    }
//

// slide kedua
    $slide2 = $presentation->createSlide(); // Tambah slide kedua
    $imageslide2 = __DIR__ . '/img/2.png'; // Pastikan gambar ada di folder img
    // Cek apakah gambar ada
    if (file_exists($imageslide2)) {
        // Gambar di slide ke-2
        $imgslide2 = $slide2->createDrawingShape();
        $imgslide2->setName('Full Slide Image 2')
            ->setDescription('Gambar Full Slide 2')
            ->setPath($imageslide2) // Menggunakan path gambar
            ->setWidth(960)        // Sesuaikan lebar (16:9)
            ->setHeight(540)       // Sesuaikan tinggi (16:9)
            ->setOffsetX(0)
            ->setOffsetY(0);       // Gambar di posisi awal
    } else {
        echo "Gambar untuk slide 2 tidak ditemukan di: " . $imageslide2;
        exit;
    }
//

// slide ketiga
    // Slide background dan layout logic
    $slide3 = $presentation->createSlide();
    $backgroundImage = __DIR__ . '/img/3.png'; // Background image path
    if (file_exists($backgroundImage)) {
        $slide3->createDrawingShape()
            ->setPath($backgroundImage)
            ->setWidth(960)
            ->setHeight(540)
            ->setOffsetX(0)
            ->setOffsetY(0);
    } else {
        echo "Background untuk slide 3 tidak ditemukan!";
        exit;
    }

    // Looping data
    if ($data) {
        $uniqueData = [];
        $counter = 0;
        $maxPerSlide = 2;

        foreach ($data as $row) {
            $uniqueKey = $row['nama_instalasi']; // Penanda unik
            if (!isset($uniqueData[$uniqueKey])) {
                $uniqueData[$uniqueKey] = $row;
                $counter++;

                // Buat slide baru setelah limit
                if ($counter > $maxPerSlide) {
                    $slide3 = $presentation->createSlide();
                    $slide3->createDrawingShape()
                        ->setPath($backgroundImage)
                        ->setWidth(960)
                        ->setHeight(540)
                        ->setOffsetX(0)
                        ->setOffsetY(0);
                    $counter = 1; // Reset counter
                }

                // Judul
                $text = $slide3->createRichTextShape();
                $text->setHeight(50)
                    ->setWidth(960)
                    ->setOffsetX(0)
                    ->setOffsetY(35);
                $text->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
                $text->createTextRun('UNIT LIFT TERPASANG')->getFont()->setSize(20)->setColor(new Color('FF000000'))->setName('Arial')->setBold(true);

                $fotoinstalasi = $row['foto_instalasi'];
                $imagePathinstalasi = __DIR__ . '/../../Index/Project/Proses/uploads/foto_instalasi/' . $fotoinstalasi;

                // Memeriksa apakah file gambar ada
                if (file_exists($imagePathinstalasi)) {
                    // Gambar di slide
                    $shapeinstalasi = $slide3->createDrawingShape(); // Gambar baru di atas nama instalasi
                    $shapeinstalasi->setName('Image Above Installation')
                        ->setDescription('Gambar di atas Nama Instalasi')
                        ->setPath($imagePathinstalasi) // Path gambar dari database
                        ->setWidth(340) // Ukuran gambar yang lebih kecil
                        ->setHeight(340)
                        ->setOffsetX(110 + (($counter - 1) % 2) * 395) // Posisi gambar bergantian berdasarkan counter
                        ->setOffsetY(100); // Offset diatur supaya berada di atas nama_instalasi
                } else {
                    echo "Gambar tidak ditemukan: " . $imagePathinstalasi;
                }

                // Teks nama instalasi
                $text1 = $slide3->createRichTextShape();
                $text1->setHeight(30)
                    ->setWidth(350)
                    ->setOffsetX(105 + (($counter - 1) % 2) * 400) // Posisi teks (bergantian)
                    ->setOffsetY(445); // Posisi vertikal teks
                $text1->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
                $text1->createTextRun($row['nama_instalasi'])->getFont()->setSize(16)->setBold(true)->setColor(new Color('FF000000'));

                // Teks deskripsi instalasi (jika diperlukan)
                // $text2 = $slide3->createRichTextShape();
                // $text2->setHeight(30)
                //     ->setWidth(350)
                //     ->setOffsetX(105 + (($counter - 1) % 2) * 400)
                //     ->setOffsetY(475);
                // $text2->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
                // $text2->createTextRun($row['deskripsi_instalasi'])->getFont()->setSize(14)->setColor(new Color('FF000000'));
            }
        }
    }
//

// slide keempat
usort($data, function ($a, $b) {
    return $a['id_komponen'] <=> $b['id_komponen']; // Ascending order
});
// Mengecek apakah ada data
$processed_komponen = []; // Array untuk track komponen yang sudah diproses
$counter = 0; // Counter buat posisi data
$maxPerSlide = 2; // Maksimum data per slide

// Path gambar background
$imageBackgroundTemuan = __DIR__ . '/img/4.png';
if (!file_exists($imageBackgroundTemuan)) {
    die("Gambar background tidak ditemukan: $imageBackgroundTemuan");
}
// Loop data untuk menambahkan ke slide
foreach ($data as $row) {
    // Skip kalau komponen udah diproses atau temuan/solusi tidak valid
    if (in_array($row['id_komponen'], $processed_komponen) || 
        $row['nama_temuan'] === 'Tidak Ada' || 
        $row['nama_solusi'] === 'Tidak Ada') {
        continue;
    }

    // Tambahin ID ke processed list
    $processed_komponen[] = $row['id_komponen'];

    // Buat slide baru setelah limit data tercapai
    if ($counter % $maxPerSlide === 0) {
        $slide4 = $presentation->createSlide();

        // Tambahkan background gambar
        $shape4 = $slide4->createDrawingShape();
        $shape4->setPath($imageBackgroundTemuan)
            ->setWidth(960)
            ->setHeight(540)
            ->setOffsetX(0)
            ->setOffsetY(0);

        // Tambahkan teks "TEMUAN"
        $textTitle = $slide4->createRichTextShape();
        $textTitle->setHeight(50)
            ->setWidth(960)
            ->setOffsetX(0)
            ->setOffsetY(35);
        $textTitle->getParagraph(0)->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
        $textTitle->createTextRun('TEMUAN')
            ->getFont()
            ->setSize(20)
            ->setBold(true)
            ->setColor(new Color('FF000000'));
    }

    // Posisi dinamis berdasarkan counter (kiri atau kanan)
    $offsetX = 65 + ($counter % $maxPerSlide) * 427;
    $offsetYBase = 408;

    // Memeriksa apakah file gambar ada
    $fotobukti = $row['foto_bukti'];
    $imagePathBukti = realpath(__DIR__ . '/../../Index/Project/Proses/uploads/foto_bukti/' . $fotobukti);

    // Cek apakah path valid
    if ($imagePathBukti && file_exists($imagePathBukti)) {
        // Mendapatkan ukuran gambar
        $imageSize = @getimagesize($imagePathBukti); 

        // Tentukan ukuran tetap untuk gambar
        $fixedWidth = 420;  // Lebar gambar tetap
        $fixedHeight = 320; // Tinggi gambar tetap

        // Jika getimagesize berhasil
        if ($imageSize !== false) {
            $width = $fixedWidth;
            $height = $fixedHeight;

            $shapeBukti = $slide4->createDrawingShape();
            $shapeBukti->setPath($imagePathBukti)
                ->setWidth($width)  // Menggunakan ukuran yang sudah di-fix
                ->setHeight($height) // Menggunakan ukuran yang sudah di-fix
                ->setOffsetX(110 + (($counter % $maxPerSlide) % 2) * 420)
                ->setOffsetY(88);
        } 
    } else {
        // Jika gambar tidak ditemukan atau path tidak valid
        echo "Gambar tidak ditemukan atau path tidak valid: $imagePathBukti\n";
        continue; // Skip data ini dan lanjutkan ke data berikutnya
    }

    // Menambahkan nomor gambar di pojok kiri atau kanan
    $nomorGambar = str_pad($counter + 1, 2, '0', STR_PAD_LEFT); // Format nomor menjadi 01, 02, dst.

    // Tentukan posisi nomor berdasarkan gambar
    $textBox = $slide4->createRichTextShape();
    $textBox->setWidth(60)               // Lebar kotak teks
            ->setHeight(45)              // Tinggi kotak teks
            ->setOffsetX(($counter % 2 == 0) ? 405 : 960 - 125) // Posisi X (Kiri untuk gambar 1, Kanan untuk gambar 2)
            ->setOffsetY(88);            // Posisi Y (Sesuaikan dengan posisi gambar)

    // Menambahkan warna latar belakang Hijau untuk kotak
    $textBox->getFill()
            ->setFillType(\PhpOffice\PhpPresentation\Style\Fill::FILL_SOLID)
            ->setStartColor(new Color('FF06BD59')); // Warna Hijau

    // Menambahkan border pada kotak
    $textBox->getBorder()->setLineWidth(1) // Ketebalan border
            ->setColor(new Color('FF06BD59')); // Warna hijau untuk border


    // Tambahkan teks nomor gambar
    $textBox->createTextRun($nomorGambar) // Menambahkan nomor gambar
            ->getFont()
            ->setSize(26)
            ->setBold(true)
            ->setColor(new Color('FFFFFF00')); // Warna teks putih


    // Tambahkan detail data lainnya
    $textKomponen = $slide4->createRichTextShape();
    $textKomponen->setHeight(25)->setWidth(400)
        ->setOffsetX($offsetX)->setOffsetY($offsetYBase);
    $textKomponen->getParagraph(0)->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
    $textKomponen->createTextRun($row['nama_komponen'] . ' : ' . $row['nama_temuan'])
        ->getFont()->setSize(12)->setBold(true)->setColor(new Color('FF000000'));

    $textTower = $slide4->createRichTextShape();
    $textTower->setHeight(25)->setWidth(400)
        ->setOffsetX($offsetX)->setOffsetY($offsetYBase + 25);
    $textTower->getParagraph(0)->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
    $textTower->createTextRun('Tower: ' . $row['nama_tower'])
        ->getFont()->setSize(12)->setColor(new Color('FF000000'));

    $textPrioritas = $slide4->createRichTextShape();
    $textPrioritas->setHeight(25)->setWidth(400)
        ->setOffsetX($offsetX)->setOffsetY($offsetYBase + 50);
    $textPrioritas->getParagraph(0)->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
    $textPrioritas->createTextRun('Prioritas: ' . $row['prioritas'])
        ->getFont()->setSize(12)->setColor(new Color('FF000000'));

    $textSolusi = $slide4->createRichTextShape();
    $textSolusi->setHeight(25)->setWidth(400)
        ->setOffsetX($offsetX)->setOffsetY($offsetYBase + 75);
    $textSolusi->getParagraph(0)->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
    $textSolusi->createTextRun('Solusi: ' . $row['nama_solusi'])
        ->getFont()->setSize(12)->setColor(new Color('FF000000'));

    $counter++;
}
//

// Pastikan folder 'PPTX' ada dan memiliki izin tulis
$folder = 'PPTX';
if (!is_dir($folder)) {
    mkdir($folder, 0777, true);
}

if (is_writable($folder)) {
    try {
        $filePath = $folder . DIRECTORY_SEPARATOR . 'Audit_3.pptx';
        $oWriter = IOFactory::createWriter($presentation, 'PowerPoint2007');
        $oWriter->save($filePath);
        echo "File PPTX berhasil disimpan di <a href='$filePath' download>Download</a>";
    } catch (Exception $e) {
        echo 'Error: ' . $e->getMessage();
    }
} else {
    echo "Folder '$folder' tidak memiliki izin tulis. Pastikan folder dapat diakses untuk menulis file.";
}
?>