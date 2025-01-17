<?php
$servername = "localhost";
$username = "root"; // ganti sesuai user database
$password = ""; // ganti sesuai password database
$dbname = "sinergi"; // nama database yang dituju

// Membuat koneksi
$conn = new mysqli($servername, $username, $password, $dbname);

// Cek koneksi
if ($conn->connect_error) {
    die("Koneksi gagal: " . $conn->connect_error);
}
