<?php
include '../../Koneksi/Koneksi.php';

// Ambil id_tower dari parameter URL
$id_tower = isset($_GET['id_tower']) ? $_GET['id_tower'] : '';

if (empty($id_tower)) {
    header("Location: Hasil_Audit.php?error=id_tower_missing");
    exit;
}

// Query untuk data komponen
$sql_komponen = "
    SELECT 
        al.id_lift, al.lift_no, al.lift_brand, al.lift_type, 
        k.nama_komponen, ak.keterangan AS komponen_keterangan, 
        ak.foto_bukti, ak.prioritas, 
        sc.nama_solusi, tc.nama_temuan
    FROM 
        audit_lift al
    LEFT JOIN 
        audit_komponen ak ON al.id_lift = ak.id_lift
    LEFT JOIN 
        solusi_komponen sc ON ak.id_solusi = sc.id_solusi
    LEFT JOIN 
        temuan_komponen tc ON ak.id_temuan = tc.id_temuan
    LEFT JOIN
        komponen k ON ak.id_komponen = k.id_komponen
    WHERE 
        al.id_tower = ?
    ORDER BY 
        al.id_lift
";

$stmt_komponen = $conn->prepare($sql_komponen);
if ($stmt_komponen === false) {
    die("Error preparing komponen query: " . $conn->error);
}
$stmt_komponen->bind_param("i", $id_tower);
$stmt_komponen->execute();
$komponen_result = $stmt_komponen->get_result();

// Query untuk data instalasi
$sql_instalasi = "
    SELECT 
        al.id_lift, al.lift_no, 
        i.foto_instalasi, i.nama_instalasi, i.deskripsi AS instalasi_deskripsi
    FROM 
        audit_lift al
    LEFT JOIN 
        instalations i ON al.id_lift = i.id_lift
    WHERE 
        al.id_tower = ?
    ORDER BY 
        al.id_lift
";

$stmt_instalasi = $conn->prepare($sql_instalasi);
if ($stmt_instalasi === false) {
    die("Error preparing instalasi query: " . $conn->error);
}
$stmt_instalasi->bind_param("i", $id_tower);
$stmt_instalasi->execute();
$instalasi_result = $stmt_instalasi->get_result();

// Proses data untuk grouping
$data = [];

// Grouping data komponen
while ($row = $komponen_result->fetch_assoc()) {
    $id_lift = $row['id_lift'];
    if (!isset($data[$id_lift])) {
        $data[$id_lift] = [
            'lift' => [
                'lift_no' => $row['lift_no'],
                'lift_brand' => $row['lift_brand'],
                'lift_type' => $row['lift_type'],
            ],
            'komponen' => [],
            'instalasi' => [],
        ];
    }
    $data[$id_lift]['komponen'][] = [
        'nama_komponen' => $row['nama_komponen'],
        'prioritas' => $row['prioritas'],
        'nama_solusi' => $row['nama_solusi'],
        'nama_temuan' => $row['nama_temuan'],
        'komponen_keterangan' => $row['komponen_keterangan'],
        'foto_bukti' => $row['foto_bukti'],
    ];
}

// Grouping data instalasi
while ($row = $instalasi_result->fetch_assoc()) {
    $id_lift = $row['id_lift'];
    if (isset($data[$id_lift])) {
        $data[$id_lift]['instalasi'][] = [
            'foto_instalasi' => $row['foto_instalasi'],
            'nama_instalasi' => $row['nama_instalasi'],
            'instalasi_deskripsi' => $row['instalasi_deskripsi'],
        ];
    }
}

?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Data Audit Tower</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="css/mobile-hasil-audit.css">
</head>

<body>
    <div class="container-fluid my-5">
        <h2 class="text-center mb-4">Data Audit</h2>
        <h5 class="mb-4">Data Audit Komponen</h5>
        <div class="table-responsive table-scroll">
            <table class="table table-bordered table-sm">
                <thead class="table-info">
                    <tr>
                        <th>Lift No</th>
                        <th>Nama Komponen</th>
                        <th>Prioritas</th>
                        <th>Solusi</th>
                        <th>Temuan</th>
                        <th>Keterangan</th>
                        <th>Foto Bukti</th>
                    </tr>
                </thead>
                <tbody>
                    <?php
                    foreach ($data as $id_lift => $lift_data) {
                        $rowspan = max(count($lift_data['komponen']), count($lift_data['instalasi']));
                        for ($i = 0; $i < $rowspan; $i++) {
                            if (isset($lift_data['komponen'][$i])) {
                                $komponen = $lift_data['komponen'][$i];
                                echo "<tr>";

                                if ($i === 0) {

                                    echo "<td rowspan='$rowspan'>" . htmlspecialchars($lift_data['lift']['lift_no']) . "</td>";
                                }

                                echo "<td>" . htmlspecialchars($komponen['nama_komponen']) . "</td>";
                                echo "<td>" . htmlspecialchars($komponen['prioritas']) . "</td>";
                                echo "<td>" . htmlspecialchars($komponen['nama_solusi']) . "</td>";
                                echo "<td>" . htmlspecialchars($komponen['nama_temuan']) . "</td>";
                                echo "<td class='truncate-text'>" . htmlspecialchars($komponen['komponen_keterangan']) . "</td>";
                                $base_url = "/sinergi_karya_mandiri/Index/Project/Proses/uploads/foto_bukti/";

                                echo "<td>";
                                if (!empty($komponen['foto_bukti'])) {
                                    // Gabungkan base URL dengan nama file dari database
                                    $image_url = $base_url . htmlspecialchars($komponen['foto_bukti']);
                                    
                                    // Tambahkan teks "Lihat Foto" sebagai link
                                    echo "<a href='" . $image_url . "' target='_blank'>Lihat Foto</a>";
                                } else {
                                    echo "Foto tidak tersedia.";
                                }
                                echo "</td>";
 
                                echo "</tr>";
                            }
                        }
                    }
                    ?>
                </tbody>
            </table>
        </div>
    </div>

    <div class="container-fluid my-5">
        <h5 class="mb-4">Data Unit Terpasang</h5>
        <div class="table-responsive table-scroll">
            <table class="table table-bordered table-sm">
                <thead class="table-info">
                    <tr>
                        <th>Lift No</th>
                        <th>Nama Instalasi</th>
                        <!-- <th>Deskripsi Instalasi</th> -->
                        <th>Foto Instalasi</th>
                    </tr>
                </thead>
                <tbody>
                    <?php
                    foreach ($data as $id_lift => $lift_data) {
                        $rowspan = count($lift_data['instalasi']);
                        for ($i = 0; $i < $rowspan; $i++) {
                            if (isset($lift_data['instalasi'][$i])) {
                                $instalasi = $lift_data['instalasi'][$i];
                                echo "<tr>";

                                if ($i === 0) {
                                    echo "<td rowspan='$rowspan' class='align-middle'>" . htmlspecialchars($lift_data['lift']['lift_no']) . "</td>";
                                }

                                echo "<td>" . htmlspecialchars($instalasi['nama_instalasi']) . "</td>";
                                // echo "<td class='truncate-text' style='max-width: 200px;'>" . htmlspecialchars($instalasi['instalasi_deskripsi']) . "</td>";
                                echo "<td>";
                                if (!empty($instalasi['foto_instalasi'])) {
                                    $base_url = "/sinergi_karya_mandiri/Index/Project/Proses/uploads/foto_instalasi/";

                                    // Gabungkan base URL dengan nama file dari database
                                    $foto_instalasi_url = $base_url . htmlspecialchars($instalasi['foto_instalasi']);

                                    // Tampilkan link "Lihat Foto"
                                    echo "<a href='" . $foto_instalasi_url . "' target='_blank'>Lihat Foto</a>";
                                } else {
                                    echo "Foto tidak tersedia.";
                                }
                                echo "</td>";

                                echo "</tr>";
                            }
                        }
                    }
                    ?>
                </tbody>
            </table>
        </div>
    </div>

</body>

</html>