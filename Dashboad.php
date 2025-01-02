<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Sinergi</title>
    <script src="https://cdn.tailwindcss.com"></script>

    <style>
        /* Tambahkan padding lebih besar untuk kartu di layar kecil */
        @media (max-width: 768px) {
            .card {
                margin-bottom: 1rem;
                padding: 1.5rem;
            }
        }
    </style>

</head>
<body class="bg-gradient-to-br from-blue-50 to-blue-100 min-h-screen">

    <!-- Navbar -->
    <nav class="bg-blue-600 shadow-md">
        <div class="container mx-auto px-4 py-3 flex justify-between items-center">
            <!-- Logo dan Menu Desktop -->
            <div class="flex items-center">
                <!-- Logo -->
                <div class="mr-6">
                    <img 
                        src="IMG/Logo_Sinergi.png" 
                        alt="Sinergi Logo" 
                        class="h-10 w-auto"
                    />
                </div>
                <!-- Menu Desktop -->
                <div class="hidden md:flex space-x-4">
                    <a href="Index/Master/Master_Komponen.php" class="text-white font-medium hover:text-gray-200 transition">Master Komponen</a>
                    <a href="../Master/Master_Solusi.php" class="text-white font-medium hover:text-gray-200 transition">Master Solusi Temuan</a>
                    <a href="../../View/Audit/Hasil_Audit.php" class="text-white font-medium hover:text-gray-200 transition">Hasil Audit</a>
                    <a href="../../View/Audit/FORM_AUDIT.php" class="text-white font-medium hover:text-gray-200 transition">Formulir Audit</a>
                </div>
            </div>

            <!-- Hamburger Menu (Mobile) -->
            <div class="md:hidden">
                <button 
                    class="text-white focus:outline-none" 
                    onclick="toggleMobileMenu()"
                    aria-label="Open main menu"
                >
                    <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5"></path>
                    </svg>
                </button>
            </div>
        </div>
    </nav>

    <!-- Mobile Menu -->
    <div id="mobileMenu" class="hidden md:hidden bg-white text-gray-800 px-4 py-3 space-y-2">
        <a href="#" class="block px-4 py-2 hover:bg-gray-100">Dashboard</a>
        <div class="border-t pt-2">
            <p class="text-gray-600 font-semibold">Master</p>
            <a href="Index/Master/Master_Komponen.php" class="block px-4 py-2 hover:bg-gray-100">Master Komponen</a>
            <a href="../../View/Master/Master_Solusi.php" class="block px-4 py-2 hover:bg-gray-100">Master Solusi dan Temuan</a>
        </div>
        <div class="border-t pt-2">
            <p class="text-gray-600 font-semibold">Audit</p>
            <a href="../../View/Audit/FORM_AUDIT.php" class="block px-4 py-2 hover:bg-gray-100">Formulir Audit Lift</a>
            <a href="../../View/Audit/Hasil_Audit.php" class="block px-4 py-2 hover:bg-gray-100">Hasil Audit Lift</a>
        </div>
    </div>


    <div class="container mx-auto my-10 px-4">
    <h4 class="text-gray-800 text-center mb-1">Selamat Datang di Dashboard</h4>
    <h1 class="text-2xl font-bold text-gray-800 text-center mb-1">PT. Sinergi Karya Mandiri</h1>


    <!-- Bagian 1: Formulir -->
    <h2 class="text-2xl font-semibold text-gray-800 mt-12 mb-4">Formulir Audit</h2>
    <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
        <!-- Card 1: Formulir Lift -->
        <div class="bg-white shadow-lg rounded-lg p-6 hover:shadow-xl transition transform hover:scale-105">
            <div class="flex items-center space-x-4">
                <div class="p-4 bg-purple-100 rounded-full">
                    <svg class="h-6 w-6 text-purple-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M4 16h16M4 12h16M4 8h16M4 4h16" />
                    </svg>
                </div>
                <h2 class="text-lg font-semibold text-gray-700">Formulir Lift</h2>
            </div>
            <p class="text-gray-500 mt-2">Isi formulir audit lift secara lengkap.</p>
            <a href="Index/Project/Identitas_Gedung.php" class="text-purple-600 hover:underline mt-2 block">Lihat Detail</a>
        </div>

        <!-- Card 2: Formulir Eskalator -->
        <div class="bg-white shadow-lg rounded-lg p-6 hover:shadow-xl transition transform hover:scale-105">
            <div class="flex items-center space-x-4">
                <div class="p-4 bg-red-100 rounded-full">
                    <svg class="h-6 w-6 text-red-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M3 7l9 4.5L21 7M3 17l9-4.5L21 17M3 7v10M21 7v10" />
                    </svg>
                </div>
                <h2 class="text-lg font-semibold text-gray-700">Formulir Eskalator</h2>
            </div>
            <p class="text-gray-500 mt-2">Isi formulir audit eskalator dengan detail.</p>
            <a href="../../View/Audit/FORM_ESKALATOR.php" class="text-red-600 hover:underline mt-2 block">Lihat Detail</a>
        </div>
    </div>

    <!-- Bagian 2: Master Data -->
    <h2 class="text-2xl font-semibold text-gray-800 mt-8 mb-4">Master Data & Laporan</h2>
    <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
        <!-- Card 1: Master Komponen -->
        <div class="bg-white shadow-lg rounded-lg p-6 hover:shadow-xl transition transform hover:scale-105">
            <div class="flex items-center space-x-4">
                <div class="p-4 bg-blue-100 rounded-full">
                    <svg class="h-6 w-6 text-blue-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m-6-8h6m-9 4a9 9 0 1118 0 9 9 0 01-18 0z" />
                    </svg>
                </div>
                <h2 class="text-lg font-semibold text-gray-700">Master Komponen</h2>
            </div>
            <p class="text-gray-500 mt-2">Kelola data komponen dengan mudah.</p>
            <a href="Index/Master/Master_Komponen.php" class="text-blue-600 hover:underline mt-2 block">Lihat Detail</a>
        </div>

        <!-- Card 2: Master Solusi -->
        <div class="bg-white shadow-lg rounded-lg p-6 hover:shadow-xl transition transform hover:scale-105">
            <div class="flex items-center space-x-4">
                <div class="p-4 bg-green-100 rounded-full">
                    <svg class="h-6 w-6 text-green-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M13 16h-1v-4h-1m1 4h1m-1 0v-4m-4-2h6m2 0h2m-8 0a9 9 0 110-18 9 9 0 010 18z" />
                    </svg>
                </div>
                <h2 class="text-lg font-semibold text-gray-700">Master Solusi</h2>
            </div>
            <p class="text-gray-500 mt-2">Kelola solusi temuan audit.</p>
            <a href="Index/Master/Master_Solusi.php" class="text-green-600 hover:underline mt-2 block">Lihat Detail</a>
        </div>

        <!-- Card 3: Hasil Audit -->
        <div class="bg-white shadow-lg rounded-lg p-6 hover:shadow-xl transition transform hover:scale-105">
            <div class="flex items-center space-x-4">
                <div class="p-4 bg-yellow-100 rounded-full">
                    <svg class="h-6 w-6 text-yellow-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m-6-8h6m-3 4a9 9 0 110 18 9 9 0 010-18z" />
                    </svg>
                </div>
                <h2 class="text-lg font-semibold text-gray-700">Hasil Audit</h2>
            </div>
            <p class="text-gray-500 mt-2">Lihat laporan hasil audit lift.</p>
            <a href="../../View/Audit/Hasil_Audit.php" class="text-yellow-600 hover:underline mt-2 block">Lihat Detail</a>
        </div>
    </div>
</div>

<script>
    function toggleMobileMenu() {
        const menu = document.getElementById('mobileMenu');
        menu.classList.toggle('hidden');
    }
</script>


</body>
</html>
