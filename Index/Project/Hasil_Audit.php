<?php
include '../../Koneksi/Koneksi.php';

// Ambil id_tower dari parameter URL
$id_tower = isset($_GET['id_tower']) ? $_GET['id_tower'] : '';

if (empty($id_tower)) {
    header("Location: Hasil_Audit.php?error=id_tower_missing");
    exit;
}

// Query gabungan untuk data
$sql = "
    SELECT 
        al.id_lift, al.lift_no, al.lift_brand, al.lift_type, 
        k.nama_komponen, ak.keterangan AS komponen_keterangan, 
        ak.foto_bukti, ak.prioritas, 
        i.foto_instalasi, i.nama_instalasi, i.deskripsi AS instalasi_deskripsi,
        sc.nama_solusi, tc.nama_temuan
    FROM 
        audit_lift al
    LEFT JOIN 
        audit_komponen ak ON al.id_lift = ak.id_lift
    LEFT JOIN 
        instalations i ON al.id_lift = i.id_lift
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

$stmt = $conn->prepare($sql);
if ($stmt === false) {
    die("Error preparing query: " . $conn->error);
}

$stmt->bind_param("i", $id_tower);
$stmt->execute();
$result = $stmt->get_result();

// Proses data untuk grouping
$data = [];
while ($row = $result->fetch_assoc()) {
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

    // Tambahkan data komponen
    $data[$id_lift]['komponen'][] = [
        'nama_komponen' => $row['nama_komponen'],
        'prioritas' => $row['prioritas'],
        'nama_solusi' => $row['nama_solusi'],
        'nama_temuan' => $row['nama_temuan'],
        'komponen_keterangan' => $row['komponen_keterangan'],
        'foto_bukti' => $row['foto_bukti'],
    ];

    // Tambahkan data instalasi
    $data[$id_lift]['instalasi'][] = [
        'foto_instalasi' => $row['foto_instalasi'],
        'nama_instalasi' => $row['nama_instalasi'],
        'instalasi_deskripsi' => $row['instalasi_deskripsi'],
    ];
}
?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Data Audit Tower</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        th, td {
            white-space: nowrap;
            text-overflow: ellipsis;
            overflow: hidden;
            padding: 10px;
        }

        th {
            text-align: center;
            background-color: #f8f9fa;
        }

        td {
            text-align: center;
            vertical-align: middle;
        }

        .truncate-text {
            max-width: 150px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .img-thumbnail {
            width: 60px;
            height: 60px;
            object-fit: cover;
            cursor: pointer;
        }

        .modal-dialog {
            max-width: 80%;
            margin: 30px auto;
        }

        .modal-body img {
            max-width: 100%;
            height: auto;
        }

        .table th, .table td {
            vertical-align: middle;
            text-align: center;
        }

        .table-bordered td, .table-bordered th {
            border: 1px solid #ddd;
        }

        .table th {
            background-color: #f8f9fa;
            font-weight: bold;
        }
    </style>
</head>

<body>
    <div class="container my-5">
        <h1 class="text-center mb-4">Data Audit Tower</h1>
        <div class="table-responsive">
            <table class="table table-bordered table-sm">
                <thead class="table-info">
                    <tr>
                        <th>Lift No</th>
                        <!-- <th>Lift Brand</th>
                        <th>Lift Type</th> -->
                        <th>Nama Komponen</th>
                        <th>Prioritas</th>
                        <th>Solusi</th>
                        <th>Temuan</th>
                        <th>Keterangan</th>
                        <th>Foto Bukti</th>
                        <th>Foto Instalasi</th>
                        <th>Nama Instalasi</th>
                        <th>Deskripsi Instalasi</th>
                    </tr>
                </thead>
                <tbody>
                    <?php
                    foreach ($data as $id_lift => $lift_data) {
                        $rowspan = max(count($lift_data['komponen']), count($lift_data['instalasi']));
                        for ($i = 0; $i < $rowspan; $i++) {
                            echo "<tr>";

                            if ($i === 0) {
                                echo "<td rowspan='$rowspan'>" . htmlspecialchars($lift_data['lift']['lift_no']) . "</td>";
                                // echo "<td rowspan='$rowspan'>" . htmlspecialchars($lift_data['lift']['lift_brand']) . "</td>";
                                // echo "<td rowspan='$rowspan'>" . htmlspecialchars($lift_data['lift']['lift_type']) . "</td>";
                            }

                            if (isset($lift_data['komponen'][$i])) {
                                $komponen = $lift_data['komponen'][$i];
                                echo "<td>" . htmlspecialchars($komponen['nama_komponen']) . "</td>";
                                echo "<td>" . htmlspecialchars($komponen['prioritas']) . "</td>";
                                echo "<td>" . htmlspecialchars($komponen['nama_solusi']) . "</td>";
                                echo "<td>" . htmlspecialchars($komponen['nama_temuan']) . "</td>";
                                echo "<td class='truncate-text'>" . htmlspecialchars($komponen['komponen_keterangan']) . "</td>";
                                echo "<td>";
                                if (!empty($komponen['foto_bukti'])) {
                                    echo "<img src='Proses/upload/" . htmlspecialchars($komponen['foto_bukti']) . "' class='img-thumbnail' data-bs-toggle='modal' data-bs-target='#modal" . $komponen['foto_bukti'] . "'>";
                                    // Modal untuk gambar besar
                                    echo "
                                    <div class='modal fade' id='modal" . $komponen['foto_bukti'] . "' tabindex='-1' aria-labelledby='modalLabel" . $komponen['foto_bukti'] . "' aria-hidden='true'>
                                        <div class='modal-dialog modal-dialog-centered modal-lg'>
                                            <div class='modal-content'>
                                                <div class='modal-header'>
                                                    <h5 class='modal-title' id='modalLabel" . $komponen['foto_bukti'] . "'>Foto Bukti</h5>
                                                    <button type='button' class='btn-close' data-bs-dismiss='modal' aria-label='Close'></button>
                                                </div>
                                                <div class='modal-body text-center'>
                                                    <img src='Proses/upload/" . htmlspecialchars($komponen['foto_bukti']) . "' class='img-fluid'>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    ";
                                }
                                echo "</td>";
                            } else {
                                echo "<td colspan='6'></td>";
                            }

                            if (isset($lift_data['instalasi'][$i])) {
                                $instalasi = $lift_data['instalasi'][$i];
                                echo "<td>";
                                if (!empty($instalasi['foto_instalasi'])) {
                                    echo "<img src='Proses/upload/" . htmlspecialchars($instalasi['foto_instalasi']) . "' class='img-thumbnail' data-bs-toggle='modal' data-bs-target='#modal" . $instalasi['foto_instalasi'] . "'>";
                                    // Modal untuk gambar instalasi
                                    echo "
                                    <div class='modal fade' id='modal" . $instalasi['foto_instalasi'] . "' tabindex='-1' aria-labelledby='modalLabel" . $instalasi['foto_instalasi'] . "' aria-hidden='true'>
                                        <div class='modal-dialog modal-dialog-centered modal-lg'>
                                            <div class='modal-content'>
                                                <div class='modal-header'>
                                                    <h5 class='modal-title' id='modalLabel" . $instalasi['foto_instalasi'] . "'>Foto Instalasi</h5>
                                                    <button type='button' class='btn-close' data-bs-dismiss='modal' aria-label='Close'></button>
                                                </div>
                                                <div class='modal-body text-center'>
                                                    <img src='Proses/upload/" . htmlspecialchars($instalasi['foto_instalasi']) . "' class='img-fluid'>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    ";
                                }
                                echo "</td>";
                                echo "<td>" . htmlspecialchars($instalasi['nama_instalasi']) . "</td>";
                                echo "<td class='truncate-text'>" . htmlspecialchars($instalasi['instalasi_deskripsi']) . "</td>";
                            } else {
                                echo "<td colspan='3'></td>";
                            }

                            echo "</tr>";
                        }
                    }
                    ?>
                </tbody>
            </table>
        </div>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.bundle.min.js"></script>
</body>

</html>
