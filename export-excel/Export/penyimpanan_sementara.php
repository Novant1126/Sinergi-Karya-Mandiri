
// Membuat Slide ke-3
$slide3 = $presentation->createSlide(); // Tambah slide 3

// Gambar di slide ke-3
$shape3 = $slide3->createDrawingShape();
$shape3->setName('Full Slide Image 3')
       ->setDescription('Gambar Full Slide 3')
       ->setPath('img/3.png') // Pastikan gambar ada di folder img
       ->setWidth(960) // Sesuaikan lebar (16:9)
       ->setHeight(540) // Sesuaikan tinggi (16:9)
       ->setOffsetX(0)
       ->setOffsetY(0); // Gambar di posisi awal

// Menambahkan teks di slide ke-3
$text = $slide3->createRichTextShape();
$text->setHeight(50)
     ->setWidth(960);

// Set posisi teks (sesuaikan dengan posisi yang diinginkan)
$text->setOffsetX(320) // Posisi horizontal (dari kiri)
     ->setOffsetY(55); // Posisi vertikal (dari atas)

// Menambahkan teks ke dalam slide
$text->createTextRun('UNIT LIFT TERPASANG')
     ->getFont()
     ->setSize(20) // Ukuran font
     ->setColor(new Color('FF000000')) // Warna font
     ->setName('Arial') // Mengatur font ke Arial
     ->setBold(true); // Membuat font menjadi bold
     



// Membuat Slide ke-2
$slide4 = $presentation->createSlide(); // Tambah slide kedua

// Gambar di slide ke-2
$shape4 = $slide4->createDrawingShape();
$shape4->setName('Full Slide Image 4')
       ->setDescription('Gambar Full Slide 4')
       ->setPath('img/4.png') // Pastikan gambar ada di folder img
       ->setWidth(960) // Sesuaikan lebar (16:9)
       ->setHeight(540) // Sesuaikan tinggi (16:9)
       ->setOffsetX(0)
       ->setOffsetY(0); // Gambar di posisi awal
