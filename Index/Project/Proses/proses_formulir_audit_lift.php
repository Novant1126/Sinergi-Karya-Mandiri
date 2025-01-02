


//        $conn->commit();
//         header("Location: ../Hasil_Audit.php?id_tower=" . $_POST['id_tower']);
//         exit;


//     } catch (Exception $e) {
//         $conn->rollback();
//         error_log($e->getMessage());
//         echo "Error: " . $e->getMessage();
//     }
// }
// ?>


        <!-- // 3. Simpan data ke tabel `audit_komponen`
        if (isset($_POST['id_komponen'])) {
            $id_komponen = $_POST['id_komponen'];
            $id_temuan = $_POST['id_temuan'];
            $id_solusi = $_POST['id_solusi'];
            $prioritas = $_POST['prioritas'];
            $keterangan = $_POST['keterangan'];
            $foto_bukti = $_FILES['foto_bukti'];

            foreach ($id_komponen as $key => $komponen_id) {
                $temuan_id = $id_temuan[$key];
                $solusi_id = $id_solusi[$key];
                $prior = $prioritas[$key];
                $ket = $keterangan[$key];
                $foto_name = $foto_bukti['name'][$key];
                $foto_tmp = $foto_bukti['tmp_name'][$key];

                // Upload file
                $target_dir_1 = "uploads/foto_bukti";
                $foto_path = $target_dir_1 . basename($foto_name);
                move_uploaded_file($foto_tmp, $foto_path);

                $stmt = $conn->prepare("INSERT INTO audit_komponen (id_komponen, id_lift, id_tower, id_gedung, id_temuan, id_solusi, prioritas, keterangan, foto_bukti) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)");
                $stmt->bind_param("iiiisiiss", $komponen_id, $id_lift, $id_tower, $id_gedung, $temuan_id, $solusi_id, $prior, $ket, $foto_path);
                $stmt->execute();
            }
        }

        // Commit transaksi
        $conn->commit();

        // Redirect ke halaman Hasil_Audit.php setelah data berhasil disimpan
        header("Location: ../Hasil_Audit.php");
        exit; // Pastikan tidak ada kode lain yang dijalankan setelah redirect

    } catch (Exception $e) {
        // Rollback transaksi jika terjadi kesalahan
        $conn->rollback();
        echo "Error: " . $e->getMessage();
    }
}
?> -->

<?php
// Terima data yang dikirim melalui POST
$data = json_decode(file_get_contents('php://input'), true);

// Koneksi ke database
include '../../../Koneksi/Koneksi.php';

foreach ($data as $item) {
    // Ambil nilai data yang dikirim
    $komponen = $item['komponen'];
    $lift = $item['lift'];
    $tower = $item['tower'];
    $gedung = $item['gedung'];
    $temuan = $item['temuan'];
    $solusi = $item['solusi'];
    $prioritas = $item['prioritas'];
    $lantai = $item['lantai'];

    // Masukkan data ke dalam tabel audit_komponen
    $stmt = $conn->prepare("INSERT INTO audit_komponen (id_komponen, id_lift, id_tower, id_gedung, keterangan, prioritas, id_temuan, id_solusi) VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
    $stmt->bind_param("iiiiiiii", $komponen, $lift, $tower, $gedung, $lantai, $prioritas, $temuan, $solusi);
    $stmt->execute();
}

// Menutup koneksi
$conn->close();

echo "Data berhasil disimpan!";
?>
