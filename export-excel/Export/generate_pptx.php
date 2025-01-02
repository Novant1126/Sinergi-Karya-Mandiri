<?php
require_once __DIR__ . '/../../vendor/autoload.php';
include '../../../adam/Koneksi/Koneksi.php';
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;
//
    // Koneksi ke database
    // $host = 'localhost';
    // $dbname = 'sinergi';
    // $username = 'root';
    // $password = '';
    
    // try {
    //     $pdo = new PDO("mysql:host=$host;dbname=$dbname", $username, $password);
    //     $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    // } catch (PDOException $e) {
    //     echo "Koneksi gagal: " . $e->getMessage();
    //     exit;
    // }
    // Ambil ID gedung, tower, dan lift dari URL
    // $id_gedung = isset($_GET['id_gedung']) ? $_GET['id_gedung'] : null;
    // $id_towers = isset($_GET['id_tower']) ? explode(',', $_GET['id_tower']) : [];
    // $id_lifts = isset($_GET['id_lift']) ? explode(',', $_GET['id_lift']) : [];

    // // Debug untuk mengecek apakah parameter sudah diterima dengan benar
    // echo "ID Gedung: $id_gedung<br>";
    // echo "ID Tower: " . implode(', ', $id_towers) . "<br>";
    // echo "ID Lift: " . implode(', ', $id_lifts) . "<br>";

    // // Pastikan variabel sudah terdefinisi dan tidak kosong
    // if (empty($id_towers) || empty($id_lifts) || !$id_gedung) {
    //     echo "ID gedung, tower, atau lift belum didefinisikan dengan benar.";
    //     exit;
    // }

    // Query dengan placeholder dinamis
    // $query = "
    // SELECT 
    //     g.id_gedung,
    //     g.nama_gedung,
    //     g.created_at AS tanggal_dibuat,
    //     i.foto_instalasi,
    //     i.nama_instalasi,
    //     i.deskripsi AS deskripsi_instalasi,
    //     ac.id_komponen,
    //     k.nama_komponen,
    //     ac.keterangan AS keterangan_komponen,
    //     ac.foto_bukti,
    //     ac.prioritas,
    //     tc.nama_temuan,
    //     sc.nama_solusi,
    //     at.nama_tower,
    //     al.lift_no
    // FROM gedung g
    // LEFT JOIN audit_tower at ON g.id_gedung = at.id_gedung
    // LEFT JOIN audit_lift al ON at.id_tower = al.id_tower
    // LEFT JOIN instalations i ON al.id_lift = i.id_lift
    // LEFT JOIN audit_komponen ac ON al.id_lift = ac.id_lift
    // LEFT JOIN komponen k ON ac.id_komponen = k.id_komponen
    // LEFT JOIN temuan_komponen tc ON ac.id_temuan = tc.id_temuan
    // LEFT JOIN solusi_komponen sc ON ac.id_solusi = sc.id_solusi
    // WHERE g.id_gedung = :id_gedung
    // AND at.id_tower IN (" . implode(',', array_map(function($i) { return ":id_tower_$i"; }, range(1, count($id_towers)))) . ")
    // AND al.id_lift IN (" . implode(',', array_map(function($i) { return ":id_lift_$i"; }, range(1, count($id_lifts)))) . ")
    // GROUP BY g.id_gedung, g.nama_gedung, i.foto_instalasi, i.nama_instalasi, i.deskripsi, ac.id_komponen, k.nama_komponen
    // ";
    // // Persiapkan array untuk parameter
    // $params = [':id_gedung' => $id_gedung];

    // // Bind parameter untuk tower
    // foreach ($id_towers as $key => $tower) {
    //     $params[":id_tower_" . ($key + 1)] = $tower;
    // }

    // // Bind parameter untuk lift
    // foreach ($id_lifts as $key => $lift) {
    //     $params[":id_lift_" . ($key + 1)] = $lift;
    // }

    // // Persiapkan dan eksekusi query
    // $stmt = $pdo->prepare($query);
    // $stmt->execute($params);

    // // Ambil hasil data
    // $result = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // // Cek hasilnya
    // if ($result) {
    //     print_r($result); // Atau manipulasi sesuai kebutuhan
    // } else {
    //     echo "Tidak ada data yang ditemukan.";
    // }
    // Tampilkan data untuk debugging (opsional)
    // foreach ($data as $row) {
    //     // echo "Nama Gedung: " . $row['nama_gedung'] . "<br>";
    //     // echo "Tanggal Dibuat: " . $row['tanggal_dibuat'] . "<br>";
    //     // echo "Nomor Lift: " . $row['lift_no'] . "<br>";
    //     // echo "Foto Instalasi: " . $row['foto_instalasi'] . "<br>";
    //     // echo "Nama Instalasi: " . $row['nama_instalasi'] . "<br>";
    //     // echo "Deskripsi Instalasi: " . $row['deskripsi_instalasi'] . "<br>";
    //     // echo "Nama Komponen: " . $row['nama_komponen'] . "<br>";
    //     // echo "Nama Tower: " . $row['nama_tower'] . "<br>";
    //     // echo "Keterangan Komponen: " . $row['keterangan_komponen'] . "<br>";
    //     // echo "Nama Temuan: " . $row['nama_temuan'] . "<br>";
    //     // echo "Nama Solusi: " . $row['nama_solusi'] . "<br><hr>";
    // }

//

if (isset($_GET['id_gedung'])) {
    $id_gedung = (int)$_GET['id_gedung'];

$query = "
    SELECT
        gedung.id_gedung,
        gedung.nama_gedung,
        gedung.project_code,
        gedung.address,
        gedung.created_at AS gedung_created_at,

        audit_tower.id_tower,
        audit_tower.nama_tower,
        audit_tower.pic,
        audit_tower.jumlah_lantai,
        audit_tower.created_at AS tower_created_at,

        audit_lift.id_lift,
        audit_lift.lift_no,
        audit_lift.lift_brand,
        audit_lift.lift_type,

        audit_komponen.id AS audit_komponen_id,
        audit_komponen.keterangan AS audit_komponen_keterangan,
        audit_komponen.foto_bukti AS audit_komponen_foto_bukti,
        audit_komponen.prioritas AS audit_komponen_prioritas,
        temuan_komponen.nama_temuan AS audit_komponen_temuan,
        solusi_komponen.nama_solusi AS audit_komponen_solusi,

   (
        SELECT GROUP_CONCAT(
            CONCAT(
                'ID: ', instalations.id_instalasi,
                ', Foto: ', instalations.foto_instalasi,
                ', Nama: ', instalations.nama_instalasi,
                ', Deskripsi: ', instalations.deskripsi
            ) SEPARATOR ' | '
        )
        FROM instalations
        WHERE instalations.id_lift = audit_lift.id_lift
    ) AS instalasi_data,

        komponen.id_komponen,
        komponen.code_komponen,
        komponen.nama_komponen,
        komponen.keterangan AS komponen_keterangan
        

    FROM audit_komponen

    -- Join ke table tower
    LEFT JOIN gedung ON gedung.id_gedung = audit_komponen.id_gedung


    -- Join ke table audit_komponen
    LEFT JOIN audit_tower ON audit_komponen.id_tower = audit_tower.id_tower

    -- Join ke table lift
    LEFT JOIN audit_lift ON audit_komponen.id_lift = audit_lift.id_lift

    -- Join ke table temuan_komponen
    LEFT JOIN temuan_komponen ON audit_komponen.id_temuan = temuan_komponen.id_temuan

    -- Join ke table solusi_komponen
    LEFT JOIN solusi_komponen ON audit_komponen.id_solusi = solusi_komponen.id_solusi

 

    -- Join ke table komponen
    LEFT JOIN komponen ON audit_komponen.id_komponen = komponen.id_komponen

    WHERE audit_komponen.id_gedung = ?
    ORDER BY komponen.id_komponen ASC
  ";
  // Eksekusi query
  $stmt = $conn->prepare($query);
  $stmt->bind_param('i', $id_gedung);
  $stmt->execute();
  $results = $stmt->get_result();
  // Ambil data
  if ($results->num_rows > 0) {
    // Proses hasil query menjadi array
    $data = [];
    $lift_komp = [];
    $lift = [];
    $komp_keterangan = [];
    $komp_keterangan_tes = [];

    $test = [];
    while ($row = $results->fetch_assoc()) {
      // Menambahkan semua data ke array $data
      $data[] = $row;
  

      // Ambil keterangan dan no_lift
      $keterangan = $row['komponen_keterangan'];
      $no_lift = $row['lift_no'];

      $komp_keterangan[$keterangan][] = $row;
      $lift[$no_lift][] = $row;
      // Jika sudah ada lift dengan nomor yang sama, tambahkan entri keterangan baru
      if (!isset($lift_komp[$no_lift])) {
        $lift_komp[$no_lift] = [];
      }

      // Menambahkan data ke keterangan yang sesuai
      if (!isset($lift_komp[$no_lift][$keterangan])) {
        $lift_komp[$no_lift][$keterangan] = [];
      }
      // Tambahkan row ke dalam keterangan
      $lift_komp[$no_lift][$keterangan][] = $row;
      if (isset($row['instalasi_data'])) {
        // Proses instalasi_data
        $instalasi_data = $row['instalasi_data'];
    } 
    
    }
  } else {
    echo "Data tidak ditemukan.";
  }


//   echo '<pre>';
//   print_r($row);
//   echo '</pre>';
//   die();

 





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
    if (isset($data['tanggal_dibuat'])) {
        $text->createTextRun($data['tanggal_dibuat'])->getFont()->setSize(12)->setColor(new Color('FF000000')); // Styling teks
    } else {
        // Tangani jika tanggal_dibuat tidak ada atau kosong
        $text->createTextRun('Tanggal tidak tersedia')->getFont()->setSize(12)->setColor(new Color('FF000000'));
    }
    

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
    if ($data) {
        $uniqueData = [];
        $counter = 0;
        $slide3 = $presentation->createSlide(); // Buat slide pertama
        
        foreach ($data as $index => $row) {
            $uniqueKey = $row['foto_instalasi']; // Penanda unik
            
            // Cek jika instalasi belum ada di $uniqueData
            if (!isset($uniqueData[$uniqueKey])) {
                $uniqueData[$uniqueKey] = $row; // Tambahkan data unik
                $counter++; // Tambahkan counter
                
                // Jika lebih dari 2 instalasi, buat slide baru
                if ($counter > 2) {
                    $slide3 = $presentation->createSlide();
                    $counter = 1; // Reset counter untuk slide baru
                }
    
                // Gambar di slide
                $shape3 = $slide3->createDrawingShape();
                $shape3->setName('Full Slide Image 3')
                    ->setDescription('Gambar Full Slide 3')
                    ->setPath(__DIR__ . '/img/3.png')
                    ->setWidth(960)
                    ->setHeight(540)
                    ->setOffsetX(0)
                    ->setOffsetY(0);
    
                // Menambahkan teks utama di slide
                $text = $slide3->createRichTextShape();
                $text->setHeight(50)
                    ->setWidth(960)
                    ->setOffsetX(0)
                    ->setOffsetY(35);
                $text->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
                $text->createTextRun('UNIT LIFT TERPASANG')->getFont()->setSize(20)->setColor(new Color('FF000000'))->setName('Arial')->setBold(true);
                
                // Ambil nama gambar dari database
                
                $imagePath = realpath(__DIR__ . '/../Proses/uploads/foto_instalasi/' . $row['foto_instalasi']);
    
                // Cek apakah gambar ada
                // if (!file_exists($imagePath)) {
                //     die("File gambar tidak ditemukan: " . $imagePath);
                // }
    
                // Gambar di slide
                $shapeinstalasi = $slide3->createDrawingShape(); // Gambar baru di atas nama instalasi
                $shapeinstalasi->setName('Image Above Installation')
                    ->setDescription('Gambar di atas Nama Instalasi')
                    ->setPath($imagePath) // Path gambar dari database
                    ->setWidth(230) // Ukuran gambar yang lebih kecil
                    ->setHeight(270)
                    ->setOffsetX(105) // Posisi gambar di atas teks nama_instalasi
                    ->setOffsetY(150); // Offset diatur supaya berada di atas nama_instalasi
    
                // Teks nama instalasi
                $text1 = $slide3->createRichTextShape();
                $text1->setHeight(30)
                    ->setWidth(350)
                    ->setOffsetX(105 + (($counter - 1) % 2) * 400) // Posisi teks (bergantian)
                    ->setOffsetY(440);
                $text1->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
                $text1->createTextRun($row['nama_instalasi'])->getFont()->setSize(16)->setBold(true)->setColor(new Color('FF000000'));
    
                // Teks deskripsi instalasi
                $text2 = $slide3->createRichTextShape();
                $text2->setHeight(30)
                    ->setWidth(350)
                    ->setOffsetX(105 + (($counter - 1) % 2) * 400)
                    ->setOffsetY(470);
                $text2->getActiveParagraph()->getAlignment()->setHorizontal(\PhpOffice\PhpPresentation\Style\Alignment::HORIZONTAL_CENTER);
                $text2->createTextRun($row['deskripsi_instalasi'])->getFont()->setSize(14)->setColor(new Color('FF000000'));
            }
        }
    }


//
// Mengecek apakah ada data
$processed_komponen = []; // Array untuk track komponen yang sudah diproses
$key = 0; // Counter buat posisi data
// Path gambar background
$imageBackground = __DIR__ . '/img/4.png';
if (!file_exists($imageBackground)) {
    die("Gambar background tidak ditemukan: $imageBackground");
}

// Loop data untuk menambahkan ke slide
foreach ($data as $row) {
    // Path untuk gambar foto_bukti di database
    $imagedatabase = realpath(__DIR__ . '/../Proses/' . $row['foto_bukti']);
    
    // Pengecekan file gambar
    if (!file_exists($imagedatabase)) {
        die("Gambar foto_bukti tidak ditemukan untuk ID Komponen: " . $row['id_komponen']);
    }

    // Skip kalau komponen udah diproses atau temuan/solusi tidak valid
    if (in_array($row['id_komponen'], $processed_komponen) || 
        $row['nama_temuan'] === 'Tidak Ada' || 
        $row['nama_solusi'] === 'Tidak Ada') {
        continue;
    }

    // Tambahin ID ke processed list
    $processed_komponen[] = $row['id_komponen'];

    // Buat slide baru setiap 2 data
    if ($key % 2 === 0) {
        $slide4 = $presentation->createSlide();

        // Tambahkan background gambar
        $shape4 = $slide4->createDrawingShape();
        $shape4->setPath($imageBackground)
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

    // Posisi dinamis berdasarkan key (kiri atau kanan)
    $offsetX = 65 + ($key % 2) * 427;
    $offsetYBase = 408;

    // Menambahkan gambar dari path foto_bukti
    $shapefotobukti = $slide4->createDrawingShape();
    $shapefotobukti->setPath($imagedatabase)
        ->setWidth(290) 
        ->setHeight(310)
        ->setOffsetX(110) 
        ->setOffsetY(98);

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

    $key++;
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
        $oWriter->save($folder . '/Audit_1.pptx');
        echo "File PPTX berhasil disimpan di " . $folder . "/Audit_1.pptx";
    } catch (Exception $e) {
        // Menangkap error dan menampilkan pesan kesalahan
        echo 'Error: ' . $e->getMessage();
    }
} else {
    echo "Folder '$folder' tidak memiliki izin tulis. Pastikan folder dapat diakses untuk menulis file.";
}
}
?>
