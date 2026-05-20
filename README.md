# e-raport-smp-it-nurhyat-iman-al-falah
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistem Informasi e-Raport & SKHU Digital - SMP IT Al-Falah</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        brand: {
                            50: '#f0fdf4',
                            100: '#dcfce7',
                            200: '#bbf7d0',
                            300: '#86efac',
                            400: '#4ade80',
                            500: '#22c55e',
                            600: '#16a34a',
                            700: '#15803d',
                            800: '#166534',
                            900: '#14532d',
                            950: '#052e16',
                        },
                        accent: {
                            emerald: '#10b981',
                        }
                    }
                }
            }
        }
    </script>
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <!-- SheetJS (Excel) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');
        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
        }
        @media print {
            .no-print {
                display: none !important;
            }
            .print-only {
                display: block !important;
            }
            body {
                background: white !important;
                color: black !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            .print-page {
                page-break-after: always;
                padding: 1.5cm !important;
                border: none !important;
                box-shadow: none !important;
                background: white !important;
            }
        }
        /* Custom scrollbar */
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 10px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
    </style>
</head>
<body class="text-slate-800 min-h-screen flex flex-col antialiased bg-slate-50">

    <!-- Cloud Syncing Loading Overlay -->
    <div id="cloud-loader" class="fixed inset-0 z-50 bg-slate-900/85 backdrop-blur-md flex flex-col items-center justify-center text-white transition-opacity duration-300">
        <div class="w-14 h-14 border-4 border-brand-500 border-t-transparent rounded-full animate-spin mb-4"></div>
        <h3 class="text-md font-bold tracking-tight" id="cloud-loader-title">Menghubungkan ke Database Cloud...</h3>
        <p class="text-xs text-slate-400 mt-1">Mengambil & menyelaraskan data e-raport terbaru dari server SMP IT Al-Falah.</p>
    </div>

    <!-- Login Screen -->
    <div id="login-screen" class="fixed inset-0 z-40 flex items-center justify-center bg-brand-950">
        <div class="bg-white p-8 sm:p-10 rounded-3xl shadow-2xl w-full max-w-md border border-brand-900/50 transform transition-all">
            <div class="text-center mb-8">
                <div class="w-16 h-16 mx-auto bg-gradient-to-tr from-brand-600 to-emerald-600 rounded-2xl flex items-center justify-center shadow-lg shadow-brand-500/30 mb-4">
                    <i class="fa-solid fa-mosque text-white text-3xl"></i>
                </div>
                <h2 class="text-2xl font-extrabold text-slate-900 tracking-tight">Portal e-Raport</h2>
                <p class="text-xs text-slate-500 font-semibold mt-1">SMP IT Nurhyat Iman Al-Falah</p>
            </div>
            <form onsubmit="handleLogin(event)" class="space-y-5">
                <div>
                    <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1.5 tracking-wider">Wali Kelas / Pengguna</label>
                    <select id="login-user" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold text-slate-700 focus:outline-none focus:ring-2 focus:ring-brand-500 appearance-none">
                        <option value="admin">Administrator (Akses Penuh)</option>
                        <option value="guru7">Wali Kelas VII</option>
                        <option value="guru8">Wali Kelas VIII</option>
                        <option value="guru9">Wali Kelas IX</option>
                    </select>
                </div>
                <div>
                    <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1.5 tracking-wider">Kata Sandi</label>
                    <input type="password" id="login-pass" required placeholder="Masukkan kata sandi..." class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm font-semibold text-slate-700 focus:outline-none focus:ring-2 focus:ring-brand-500">
                    <p id="login-error" class="hidden text-[10px] text-rose-500 font-bold mt-2"><i class="fa-solid fa-circle-exclamation mr-1"></i>Kata sandi tidak sesuai!</p>
                </div>
                <button type="submit" class="w-full bg-brand-600 hover:bg-brand-700 text-white font-bold text-sm px-4 py-3.5 rounded-xl shadow-lg shadow-brand-500/30 transition duration-200 mt-2">
                    Masuk Sistem
                </button>
            </form>
            <div class="mt-8 text-center border-t border-slate-100 pt-6">
                <p class="text-[10px] text-slate-400 font-semibold">Info Default Sandi: admin / kelas7 / kelas8 / kelas9</p>
            </div>
        </div>
    </div>

    <!-- Application Header -->
    <header id="app-header" class="hidden bg-brand-950 text-white shadow-xl no-print sticky top-0 z-30 border-b border-brand-900/50">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
            <div class="flex items-center gap-3">
                <div class="w-10 h-10 bg-brand-600 rounded-xl flex items-center justify-center text-white font-extrabold shadow-inner shadow-black/20">
                    <i class="fa-solid fa-mosque text-xl"></i>
                </div>
                <div>
                    <h1 class="text-sm font-extrabold tracking-tight">SMP IT AL-FALAH</h1>
                    <p class="text-[10px] text-brand-300 font-semibold uppercase tracking-wider flex items-center gap-2">
                        e-Raport & SKHU Digital v4
                        <span id="cloud-badge" class="px-2 py-0.5 rounded-full text-[8px] bg-amber-500/20 text-amber-300 border border-amber-500/30 flex items-center gap-1 font-bold">
                            <span class="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse"></span> Mode Lokal
                        </span>
                    </p>
                </div>
            </div>
            
            <div class="flex items-center space-x-4">
                <div class="text-right hidden sm:block">
                    <p class="text-[10px] text-brand-300 font-bold uppercase tracking-widest" id="header-kelas-wali">Mengambil Data...</p>
                    <p class="text-xs font-bold text-slate-100" id="header-nama-wali">Administrator</p>
                </div>
                <button onclick="handleLogout()" class="w-10 h-10 rounded-xl bg-rose-500/20 hover:bg-rose-500 flex items-center justify-center text-rose-100 font-extrabold shadow-inner border border-rose-400/20 transition-all" title="Keluar Akun">
                    <i class="fa-solid fa-power-off"></i>
                </button>
            </div>
        </div>
    </header>

    <!-- Main Container -->
    <div id="app-content" class="hidden flex-1 max-w-7xl w-full mx-auto px-0 sm:px-6 lg:px-8 py-8 gap-8">
        <!-- Sidebar Navigation -->
        <aside class="w-64 shrink-0 hidden md:flex flex-col gap-6 no-print">
            <nav class="bg-white rounded-3xl border border-slate-100 p-4 shadow-sm flex flex-col gap-1">
                <p class="text-[10px] font-extrabold text-slate-400 px-3 py-2.5 uppercase tracking-wider">Sajian Menu</p>
                
                <button onclick="switchTab('dashboard')" id="nav-dashboard" class="nav-btn flex items-center gap-3.5 px-3 py-3 rounded-2xl transition-all duration-150 text-xs font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-chart-pie w-5 text-center text-base text-brand-500"></i>
                    Dasbor Utama
                </button>

                <button onclick="switchTab('siswa')" id="nav-siswa" class="nav-btn flex items-center gap-3.5 px-3 py-3 rounded-2xl transition-all duration-150 text-xs font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-graduation-cap w-5 text-center text-base text-brand-500"></i>
                    Manajemen Siswa
                </button>

                <button onclick="switchTab('nilai')" id="nav-nilai" class="nav-btn flex items-center gap-3.5 px-3 py-3 rounded-2xl transition-all duration-150 text-xs font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-pen-to-square w-5 text-center text-base text-brand-500"></i>
                    Input Nilai Akademik
                </button>

                <button onclick="switchTab('raport')" id="nav-raport" class="nav-btn flex items-center gap-3.5 px-3 py-3 rounded-2xl transition-all duration-150 text-xs font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-print w-5 text-center text-base text-brand-500"></i>
                    Cetak Raport Resmi
                </button>

                <button onclick="switchTab('skhu')" id="nav-skhu" class="nav-btn flex items-center gap-3.5 px-3 py-3 rounded-2xl transition-all duration-150 text-xs font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-certificate w-5 text-center text-base text-brand-500"></i>
                    Cetak SKHU Lulus
                </button>
            </nav>

            <nav class="bg-white rounded-3xl border border-slate-100 p-4 shadow-sm flex flex-col gap-1">
                <p class="text-[10px] font-extrabold text-slate-400 px-3 py-2.5 uppercase tracking-wider">Integrasi Data</p>

                <button onclick="exportTabelExcel()" class="flex items-center gap-3.5 px-3 py-2 rounded-2xl transition-all duration-150 text-[11px] font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-file-export w-5 text-center text-sm text-indigo-500"></i>
                    Ekspor Nilai (Tabel)
                </button>
                <button onclick="triggerExcelImportTabel()" class="flex items-center gap-3.5 px-3 py-2 rounded-2xl transition-all duration-150 text-[11px] font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-file-import w-5 text-center text-sm text-indigo-500"></i>
                    Impor Nilai (Tabel)
                </button>
                <input type="file" id="excel-tabel-input" onchange="importDariTabel(event)" class="hidden" accept=".xlsx, .xls">

                <div class="h-px bg-slate-100 my-1"></div>

                <button onclick="exportSiswaExcel()" class="flex items-center gap-3.5 px-3 py-2 rounded-2xl transition-all duration-150 text-[11px] font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-user-export w-5 text-center text-sm text-brand-600"></i>
                    Ekspor Data Siswa
                </button>
                <button onclick="triggerSiswaImport()" class="flex items-center gap-3.5 px-3 py-2 rounded-2xl transition-all duration-150 text-[11px] font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-user-plus w-5 text-center text-sm text-brand-600"></i>
                    Impor Data Siswa
                </button>
                <input type="file" id="excel-siswa-input" onchange="importSiswaExcel(event)" class="hidden" accept=".xlsx, .xls">

                <div class="h-px bg-slate-100 my-1"></div>

                <button onclick="exportEkskulExcel()" class="flex items-center gap-3.5 px-3 py-2 rounded-2xl transition-all duration-150 text-[11px] font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-file-export w-5 text-center text-sm text-amber-500"></i>
                    Ekspor Ekskul & Absen
                </button>
                <button onclick="triggerEkskulImport()" class="flex items-center gap-3.5 px-3 py-2 rounded-2xl transition-all duration-150 text-[11px] font-bold text-slate-600 hover:bg-slate-50 w-full text-left">
                    <i class="fa-solid fa-file-import w-5 text-center text-sm text-amber-500"></i>
                    Impor Ekskul & Absen
                </button>
                <input type="file" id="excel-ekskul-input" onchange="importEkskulExcel(event)" class="hidden" accept=".xlsx, .xls">

                <div class="h-px bg-slate-100 my-1"></div>

                <button onclick="exportJSONBackup()" class="flex items-center gap-3.5 px-3 py-2 rounded-2xl transition-all duration-150 text-[11px] font-bold text-slate-500 hover:bg-rose-50 hover:text-rose-600 w-full text-left">
                    <i class="fa-solid fa-database w-5 text-center text-sm text-rose-500"></i>
                    Unduh Cadangan JSON
                </button>
                <button onclick="triggerJSONImport()" class="flex items-center gap-3.5 px-3 py-2 rounded-2xl transition-all duration-150 text-[11px] font-bold text-slate-500 hover:bg-rose-50 hover:text-rose-600 w-full text-left">
                    <i class="fa-solid fa-upload w-5 text-center text-sm text-rose-500"></i>
                    Impor Cadangan JSON
                </button>
                <input type="file" id="json-backup-input" onchange="importJSONBackup(event)" class="hidden" accept=".json">
            </nav>
        </aside>

        <!-- Main Workspace Area -->
        <main class="flex-1 min-w-0 flex flex-col gap-6 px-4 sm:px-0">
            <!-- Global Feedback Toast Notification -->
            <div id="toast-notif" class="hidden no-print items-center gap-3 bg-slate-900 text-white rounded-2xl px-5 py-4 shadow-xl fixed bottom-6 right-6 z-50 animate-bounce">
                <i id="toast-icon" class="fa-solid fa-circle-check text-brand-400 text-lg"></i>
                <div>
                    <h4 id="toast-title" class="text-xs font-extrabold">Notifikasi</h4>
                    <p id="toast-desc" class="text-[10px] text-slate-300 font-semibold">Tindakan berhasil diselesaikan.</p>
                </div>
            </div>

            <!-- Tab: Dashboard -->
            <section id="tab-dashboard" class="tab-content flex flex-col gap-6">
                <!-- Welcome Widget -->
                <div class="bg-gradient-to-r from-brand-900 to-emerald-950 p-6 sm:p-8 rounded-3xl text-white relative overflow-hidden shadow-sm">
                    <div class="relative z-10">
                        <span class="bg-brand-500/20 text-brand-300 text-[10px] font-extrabold px-3 py-1 rounded-full border border-brand-500/30 uppercase tracking-widest">SMP IT AL-FALAH</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold tracking-tight mt-3">Sistem Pengisian e-Raport Mandiri</h2>
                        <p class="text-xs text-slate-300 font-semibold mt-1 max-w-xl">Ubah isian identitas murid, nilai per semester, maupun kop surat dengan dinamis. Seluruh dokumen disimpan otomatis pada penyimpanan cloud atau lokal secara aman.</p>
                    </div>
                </div>

                <!-- Stats Widgets -->
                <div class="grid grid-cols-1 sm:grid-cols-3 gap-6">
                    <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex items-center gap-4">
                        <div class="w-12 h-12 rounded-2xl bg-brand-50 text-brand-600 flex items-center justify-center text-xl font-bold"><i class="fa-solid fa-user-group"></i></div>
                        <div>
                            <p class="text-[10px] text-slate-400 font-extrabold uppercase tracking-wider">Total Siswa Aktif</p>
                            <h3 id="stat-siswa-total" class="text-xl font-extrabold text-slate-900">0 Siswa</h3>
                        </div>
                    </div>
                    <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex items-center gap-4">
                        <div class="w-12 h-12 rounded-2xl bg-indigo-50 text-indigo-600 flex items-center justify-center text-xl font-bold"><i class="fa-solid fa-book"></i></div>
                        <div>
                            <p class="text-[10px] text-slate-400 font-extrabold uppercase tracking-wider">Mata Pelajaran</p>
                            <h3 class="text-xl font-extrabold text-slate-900">15 Mapel</h3>
                        </div>
                    </div>
                    <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex items-center gap-4">
                        <div class="w-12 h-12 rounded-2xl bg-amber-50 text-amber-600 flex items-center justify-center text-xl font-bold"><i class="fa-solid fa-circle-check"></i></div>
                        <div>
                            <p class="text-[10px] text-slate-400 font-extrabold uppercase tracking-wider">Semester Aktif</p>
                            <h3 id="stat-semester-active" class="text-xl font-extrabold text-slate-900">Semester Ganjil (1)</h3>
                        </div>
                    </div>
                </div>

                <!-- Configuration Panels (Two columns) -->
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    <!-- Left: Quick Settings -->
                    <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex flex-col gap-4">
                        <h3 class="text-sm font-extrabold text-slate-900 tracking-tight flex items-center gap-2"><i class="fa-solid fa-sliders text-brand-600"></i> Kontrol Kelas & Semester</h3>
                        <div class="grid grid-cols-3 gap-3">
                            <button id="btn-class-VII" onclick="changeActiveKelas('VII')" class="py-2.5 px-4 rounded-xl text-xs font-bold border transition-all text-slate-500 border-slate-200 hover:bg-slate-50">Kelas VII</button>
                            <button id="btn-class-VIII" onclick="changeActiveKelas('VIII')" class="py-2.5 px-4 rounded-xl text-xs font-bold border transition-all text-slate-500 border-slate-200 hover:bg-slate-50">Kelas VIII</button>
                            <button id="btn-class-IX" onclick="changeActiveKelas('IX')" class="py-2.5 px-4 rounded-xl text-xs font-bold border transition-all text-slate-500 border-slate-200 hover:bg-slate-50">Kelas IX</button>
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <button id="btn-semester-1" onclick="changeActiveSemester(1)" class="py-2.5 px-4 rounded-xl text-xs font-bold border transition-all text-slate-500 border-slate-200 hover:bg-slate-50">Semester 1 (Ganjil)</button>
                            <button id="btn-semester-2" onclick="changeActiveSemester(2)" class="py-2.5 px-4 rounded-xl text-xs font-bold border transition-all text-slate-500 border-slate-200 hover:bg-slate-50">Semester 2 (Genap)</button>
                        </div>
                    </div>

                    <!-- Right: Document Sign & Meta Information Configuration -->
                    <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex flex-col gap-4">
                        <h3 class="text-sm font-extrabold text-slate-900 tracking-tight flex items-center gap-2"><i class="fa-solid fa-pen-fancy text-brand-600"></i> Pengenal & Tanda Tangan Dokumen</h3>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-[9px] font-extrabold text-slate-500 uppercase mb-1">Tempat Penerbitan</label>
                                <input type="text" id="config-tempat" oninput="updateConfig('tempat', this.value)" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                            </div>
                            <div>
                                <label class="block text-[9px] font-extrabold text-slate-500 uppercase mb-1">Nomor Surat SKHU</label>
                                <input type="text" id="config-no-surat" oninput="updateConfig('noSurat', this.value)" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                            </div>
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-[9px] font-extrabold text-slate-500 uppercase mb-1">Tanggal Rapor</label>
                                <input type="text" id="config-tgl-rapor" oninput="updateConfig('tglRapor', this.value)" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                            </div>
                            <div>
                                <label class="block text-[9px] font-extrabold text-slate-500 uppercase mb-1">Tanggal SKHU</label>
                                <input type="text" id="config-tgl-skhu" oninput="updateConfig('tglSkhu', this.value)" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                            </div>
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-[9px] font-extrabold text-slate-500 uppercase mb-1">Nama Wali Kelas</label>
                                <input type="text" id="config-wali-nama" oninput="updateConfig('waliNama', this.value)" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                            </div>
                            <div>
                                <label class="block text-[9px] font-extrabold text-slate-500 uppercase mb-1">Nama Kepala Sekolah</label>
                                <input type="text" id="config-kepsek-nama" oninput="updateConfig('kepsekNama', this.value)" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                            </div>
                        </div>
                    </div>
                </div>
            </section>

            <!-- Tab: Data Siswa -->
            <section id="tab-siswa" class="tab-content hidden flex-col gap-6">
                <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
                    <div class="flex items-center justify-between mb-4">
                        <h3 class="text-sm font-extrabold text-slate-900 flex items-center gap-2"><i class="fa-solid fa-users text-brand-600"></i> Daftar Siswa Kelas</h3>
                        <button onclick="openAddSiswaModal()" class="bg-brand-600 hover:bg-brand-700 text-white font-bold text-[10px] px-3.5 py-2 rounded-xl transition duration-150">
                            Tambah Siswa Baru
                        </button>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-xs border-collapse">
                            <thead>
                                <tr class="bg-slate-50 text-slate-400 font-extrabold border-b border-slate-100 uppercase tracking-widest text-[9px]">
                                    <th class="py-3 px-4">Nama Lengkap</th>
                                    <th class="py-3 px-4">NISN</th>
                                    <th class="py-3 px-4">Tempat, Tgl Lahir</th>
                                    <th class="py-3 px-4">Nama Orang Tua</th>
                                    <th class="py-3 px-4 text-center">Tindakan</th>
                                </tr>
                            </thead>
                            <tbody id="siswa-table-body" class="divide-y divide-slate-100 font-medium text-slate-700">
                                <!-- Rendered dynamically -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- Tab: Input Nilai -->
            <section id="tab-nilai" class="tab-content hidden flex-col gap-6">
                <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex flex-col gap-6">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div>
                            <h3 class="text-sm font-extrabold text-slate-900 flex items-center gap-2"><i class="fa-solid fa-square-poll-vertical text-brand-600"></i> Pengisian Nilai & Catatan Wali Kelas</h3>
                            <p class="text-[10px] text-slate-400 font-bold mt-0.5">Silakan pilih murid terlebih dahulu.</p>
                        </div>
                        <select id="nilai-student-selector" onchange="loadStudentGrades()" class="bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-bold text-slate-700 focus:outline-none focus:ring-1 focus:ring-brand-500">
                            <!-- Dynamic Option -->
                        </select>
                    </div>

                    <div id="nilai-input-area" class="hidden flex-col gap-6">
                        <!-- Nilai Mapel Grid -->
                        <div>
                            <h4 class="text-xs font-bold text-slate-400 uppercase tracking-widest mb-3">Daftar Nilai Akademik</h4>
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-4" id="grades-input-grid">
                                <!-- Dynamic input form for 15 mapels -->
                            </div>
                        </div>

                        <!-- Tambahan Ekskul & Absensi Manual Input -->
                        <div class="border-t border-slate-100 pt-6 flex flex-col gap-4">
                            <h4 class="text-xs font-bold text-slate-400 uppercase tracking-widest">Ekstrakurikuler & Absensi (Kehadiran)</h4>
                            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                                <!-- Ekskul 1 -->
                                <div class="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-2">
                                    <h5 class="text-[11px] font-bold text-slate-600 uppercase">Ekstrakurikuler 1</h5>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Nama Kegiatan</label>
                                        <input type="text" id="ekskul1-nama" placeholder="Misal: Pramuka" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Nilai (A / B / C)</label>
                                        <input type="text" id="ekskul1-nilai" placeholder="A / B / C" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Keterangan / Deskripsi</label>
                                        <input type="text" id="ekskul1-ket" placeholder="Aktif mengikuti kegiatan" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                </div>

                                <!-- Ekskul 2 -->
                                <div class="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-2">
                                    <h5 class="text-[11px] font-bold text-slate-600 uppercase">Ekstrakurikuler 2</h5>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Nama Kegiatan</label>
                                        <input type="text" id="ekskul2-nama" placeholder="Misal: UKS / Paskibra" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Nilai (A / B / C)</label>
                                        <input type="text" id="ekskul2-nilai" placeholder="A / B / C" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Keterangan / Deskripsi</label>
                                        <input type="text" id="ekskul2-ket" placeholder="Melaksanakan tugas dengan baik" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                </div>

                                <!-- Absensi -->
                                <div class="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-2">
                                    <h5 class="text-[11px] font-bold text-slate-600 uppercase">Ketidakhadiran (Absensi)</h5>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Sakit (Hari)</label>
                                        <input type="number" id="absen-sakit" min="0" placeholder="0" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Izin (Hari)</label>
                                        <input type="number" id="absen-izin" min="0" placeholder="0" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                    <div>
                                        <label class="text-[9px] font-extrabold text-slate-400 block mb-0.5">Tanpa Keterangan / Alpa (Hari)</label>
                                        <input type="number" id="absen-alpa" min="0" placeholder="0" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none">
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- Catatan Wali Kelas -->
                        <div class="border-t border-slate-100 pt-6 flex flex-col gap-4">
                            <h4 class="text-xs font-bold text-slate-400 uppercase tracking-widest">Catatan & Status Siswa</h4>
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                                <div>
                                    <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1.5 tracking-wider">Keputusan Akhir / Kenaikan / Kelulusan</label>
                                    <input type="text" id="catatan-keputusan" placeholder="Contoh: NAIK KE KELAS VIII atau LULUS" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                                </div>
                                <div>
                                    <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1.5 tracking-wider">Catatan Motivasi Wali Kelas</label>
                                    <textarea id="catatan-teks" placeholder="Berikan evaluasi, nasehat dan motivasi belajar murid..." class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500 h-24 resize-none"></textarea>
                                </div>
                            </div>
                        </div>

                        <div class="flex justify-end gap-3">
                            <button onclick="saveAllGrades()" class="bg-brand-600 hover:bg-brand-700 text-white font-bold text-xs px-6 py-3 rounded-xl shadow-lg shadow-brand-500/20 transition duration-150">
                                Simpan Nilai & Catatan
                            </button>
                        </div>
                    </div>
                </div>
            </section>

            <!-- Tab: Cetak Rapor -->
            <section id="tab-raport" class="tab-content hidden flex-col gap-6">
                <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex flex-col gap-4 no-print">
                    <div class="flex items-center justify-between">
                        <div>
                            <h3 class="text-sm font-extrabold text-slate-900">Lembar Cetak Raport</h3>
                            <p class="text-[10px] text-slate-400 font-bold">Tekan cetak untuk memunculkan lembar dialog cetak peramban asli.</p>
                        </div>
                        <div class="flex gap-2">
                            <select id="raport-student-selector" onchange="renderRaportDoc()" class="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-700 focus:outline-none">
                                <!-- Dynamic Selector -->
                            </select>
                            <button onclick="window.print()" class="bg-brand-600 hover:bg-brand-700 text-white font-bold text-xs px-4 py-2 rounded-xl transition">
                                <i class="fa-solid fa-print mr-1"></i> Cetak Dokumen
                            </button>
                        </div>
                    </div>
                </div>

                <!-- Printable Document Raport -->
                <div id="raport-document" class="bg-white rounded-3xl border border-slate-200 p-8 sm:p-12 shadow-sm max-w-[21cm] mx-auto w-full print-page text-sm leading-relaxed text-black">
                    <!-- Dynamic rendering in JS -->
                </div>
            </section>

            <!-- Tab: Cetak SKHU -->
            <section id="tab-skhu" class="tab-content hidden flex-col gap-6">
                <div class="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex flex-col gap-4 no-print">
                    <div class="flex items-center justify-between">
                        <div>
                            <h3 class="text-sm font-extrabold text-slate-900">Lembar Surat Keterangan Hasil Ujian (SKHU)</h3>
                            <p class="text-[10px] text-slate-400 font-bold">Silakan sesuaikan konfigurasi nomor surat dan tgl tanda tangan di Dasbor utama jika perlu.</p>
                        </div>
                        <div class="flex gap-2">
                            <select id="skhu-student-selector" onchange="renderSKHUDoc()" class="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-700 focus:outline-none">
                                <!-- Dynamic selector -->
                            </select>
                            <button onclick="window.print()" class="bg-brand-600 hover:bg-brand-700 text-white font-bold text-xs px-4 py-2 rounded-xl transition">
                                <i class="fa-solid fa-print mr-1"></i> Cetak SKHU
                            </button>
                        </div>
                    </div>
                </div>

                <!-- Printable Document SKHU -->
                <div id="skhu-document" class="bg-white rounded-3xl border border-slate-200 p-8 sm:p-12 shadow-sm max-w-[21cm] mx-auto w-full print-page text-sm leading-relaxed text-black">
                    <!-- Dynamic rendering in JS -->
                </div>
            </section>
        </main>
    </div>

    <!-- Modal: Tambah/Edit Siswa -->
    <div id="siswa-modal" class="hidden fixed inset-0 z-50 bg-black/50 items-center justify-center p-4">
        <div class="bg-white rounded-3xl w-full max-w-md overflow-hidden shadow-2xl border border-slate-100">
            <div class="p-6 bg-slate-50 border-b border-slate-100 flex items-center justify-between">
                <h3 id="siswa-modal-title" class="text-sm font-extrabold text-slate-900">Tambah Siswa Baru</h3>
                <button onclick="closeSiswaModal()" class="text-slate-400 hover:text-slate-600"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            <form onsubmit="handleSiswaSubmit(event)" class="p-6 space-y-4">
                <input type="hidden" id="siswa-form-id">
                <div>
                    <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1">Nama Lengkap Murid</label>
                    <input type="text" id="siswa-form-nama" required placeholder="Contoh: Adil Pranata" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-semibold focus:ring-1 focus:ring-brand-500">
                </div>
                <div>
                    <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1">Nomor Induk Siswa Nasional (NISN)</label>
                    <input type="text" id="siswa-form-nisn" required placeholder="10 digit nomor NISN" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-semibold focus:ring-1 focus:ring-brand-500">
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1">Tempat Lahir</label>
                        <input type="text" id="siswa-form-tempat" placeholder="Medan" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-semibold focus:ring-1 focus:ring-brand-500">
                    </div>
                    <div>
                        <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1">Tanggal Lahir</label>
                        <input type="text" id="siswa-form-tanggal" placeholder="12 Maret 2008" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-semibold focus:ring-1 focus:ring-brand-500">
                    </div>
                </div>
                <div>
                    <label class="block text-[10px] font-extrabold text-slate-500 uppercase mb-1">Nama Orang Tua / Wali</label>
                    <input type="text" id="siswa-form-ortu" required placeholder="Nama Ayah/Ibu/Wali" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-semibold focus:ring-1 focus:ring-brand-500">
                </div>
                <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
                    <button type="button" onclick="closeSiswaModal()" class="px-4 py-2.5 text-xs font-bold text-slate-500 hover:bg-slate-50 rounded-xl">Batal</button>
                    <button type="submit" class="px-5 py-2.5 text-xs font-bold bg-brand-600 hover:bg-brand-700 text-white rounded-xl">Simpan Data</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Firebase SDK Imports (Using ES Modules natively) -->
    <script type="module">
        import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
        import { getAuth, signInWithCustomToken, signInAnonymously } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
        import { getFirestore, doc, getDoc, setDoc, onSnapshot } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

        // Declare dynamic window handles to communicate with general legacy JS script block
        window.firebaseLoaded = false;
        if (typeof __firebase_config !== 'undefined' && __firebase_config) {
            try {
                const config = JSON.parse(__firebase_config);
                const app = initializeApp(config);
                const auth = getAuth(app);
                const db = getFirestore(app);
                const appId = typeof __app_id !== 'undefined' ? __app_id : 'erapor-alfalah-cloud-v1';

                window.firebaseAuth = auth;
                window.firebaseDb = db;
                window.firebaseAppId = appId;
                window.firebaseDoc = doc;
                window.firebaseGetDoc = getDoc;
                window.firebaseSetDoc = setDoc;
                window.firebaseOnSnapshot = onSnapshot;
                window.firebaseLoaded = true;
                window.signInAnonymously = signInAnonymously;
                window.signInWithCustomToken = signInWithCustomToken;
            } catch (e) {
                console.error("Gagal melakukan inisialisasi modul Firebase SDK:", e);
            }
        }
    </script>

    <!-- Script Block -->
    <script>
        // Global variables for cloud status
        let isCloudMode = false;

        // Database State
        let state = {
            activeKelas: 'IX',
            activeSemester: 1,
            config: {
                tempat: 'Tanjung Siram',
                noSurat: '045.2 / SMP-IT.AF / SKHU / V / 2026',
                tglRapor: '18 Juni 2026',
                tglSkhu: '12 Mei 2026',
                waliNama: 'Drs. H. Ahmad Fauzi',
                kepsekNama: 'M. Yusuf, S.Pd.I'
            },
            mapel: [],
            siswa: {
                'VII': [],
                'VIII': [],
                'IX': []
            },
            nilai: {},      // key: {studentId}_{semesterId}_{mapelId} -> { p: 90, k: 90 }
            catatan: {},    // key: {studentId}_{semesterId} -> { teks: X, keputusan: Y }
            ekskul: {},     // key: {studentId}_{semesterId} -> { e1Nama, e1Nilai, e1Ket, e2Nama, e2Nilai, e2Ket }
            absensi: {}     // key: {studentId}_{semesterId} -> { sakit, izin, alpa }
        };

        // Static 15 Mapel Kurikulum SMP IT Al-Falah
        const defaultMapel = [
            { id: 'mp-1', kode: 'SBD', nama: 'Seni Budaya', descP: 'memahami konsep, gaya, dan teknik karya seni rupa, musik, tari, atau teater', descK: 'membuat karya seni rupa, menyanyikan lagu, memperagakan gerakan tari, atau mementaskan teater' },
            { id: 'mp-2', kode: 'PPKN', nama: 'Pancasila & Kewarganegaraan', descP: 'memahami nilai-nilai Pancasila, konstitusi, tata hukum, dan hak kewajiban warga negara', descK: 'menerapkan komitmen kebangsaan, norma-norma, dan partisipasi aktif dalam masyarakat' },
            { id: 'mp-3', kode: 'BIND', nama: 'Bahasa Indonesia', descP: 'memahami struktur, aspek kebahasaan, gagasan, dan pesan dalam teks formal maupun sastra', descK: 'menyajikan ide, gagasan, dan ulasan kreatif secara tertulis maupun lisan dengan efektif' },
            { id: 'mp-4', kode: 'MTK', nama: 'Matematika', descP: 'memahami konsep aljabar, geometri, statistika, peluang, dan penalaran logika matematis', descK: 'menyelesaikan masalah kontekstual yang berkaitan dengan model matematika dan pola bilangan' },
            { id: 'mp-5', kode: 'IPA', nama: 'Ilmu Pengetahuan Alam', descP: 'memahami sistem klasifikasi, interaksi makhluk hidup, energi, materi, dan gejala alam sekitarnya', descK: 'melakukan penyelidikan ilmiah, eksperimen sederhana, dan menyajikan laporan hasil observasi' },
            { id: 'mp-6', kode: 'IPS', nama: 'Ilmu Pengetahuan Sosial', descP: 'memahami keruangan, interaksi sosial, sejarah perjuangan, dan aktivitas pemenuhan kebutuhan ekonomi', descK: 'menyajikan analisis perubahan sosial ekonomi dan peta interaksi kebudayaan antarwilayah' },
            { id: 'mp-7', kode: 'BING', nama: 'Bahasa Inggris', descP: 'memahami fungsi sosial, struktur teks, dan unsur kebahasaan interaksi transaksional sehari-hari', descK: 'menyusun teks lisan dan tulis pendek untuk berkomunikasi aktif dengan tata bahasa yang berterima' },
            { id: 'mp-8', kode: 'PJOK', nama: 'Pendidikan Jasmani & Kesehatan', descP: 'memahami prinsip kebugaran, pola hidup sehat, aturan permainan, dan teknik dasar olahraga', descK: 'mempraktikkan keterampilan gerak spesifik olahraga, kebugaran jasmani, dan menjaga kesehatan diri' },
            { id: 'mp-9', kode: 'BARB', nama: 'Bahasa Arab (Muatan Lokal)', descP: 'memahami struktur teks percakapan, mufrodat, dan makna ungkapan bertema aktivitas sekolah', descK: 'melakukan hiwar (percakapan) sederhana dan menulis kalimat terstruktur dalam Bahasa Arab' },
            { id: 'mp-10', kode: 'PRK', nama: 'Prakarya', descP: 'memahami prinsip perancangan, pembuatan, dan pengemasan produk kerajinan atau pengolahan pangan', descK: 'membuat produk kerajinan fungsional atau menyajikan masakan kreatif bernilai ekonomis' },
            { id: 'mp-11', kode: 'TIK', nama: 'Teknologi Informasi & Komputer', descP: 'memahami sistem komputer, jaringan internet, algoritma pemrograman, dan dampak sosial informatika', descK: 'mengoperasikan aplikasi perkantoran, membuat program logika sederhana, dan mengolah data digital' },
            { id: 'mp-12', kode: 'AA', nama: 'Akidah Akhlak', descP: 'memahami dasar-dasar akidah Islam, sifat-sifat Allah, serta keteladanan para nabi dan rasul', descK: 'menerapkan akhlak terpuji dalam kehidupan sehari-hari dan menghindari perilaku tercela' },
            { id: 'mp-13', kode: 'PAI', nama: 'Pendidikan Agama Islam', descP: 'memahami rukun iman, rukun islam, serta sejarah agama dan tata cara peribadatan', descK: 'menunjukkan akhlak mulia dan mempraktikkan ibadah harian sesuai syariat Islam' },
            { id: 'mp-14', kode: 'SKI', nama: 'Sejarah Kebudayaan Islam', descP: 'memahami sejarah perjuangan dakwah Nabi Muhammad SAW, Khulafaur Rasyidin, dan daulah Islamiyah', descK: 'mengambil ibrah dan keteladanan dari perkembangan peradaban serta tokoh-tokoh Islam terdahulu' },
            { id: 'mp-15', kode: 'PI', nama: 'Praktek Ibadah', descP: 'memahami tata cara thaharah, wudhu, shalat fardhu, shalat sunnah, zikir, dan doa setelah shalat', descK: 'mempraktikkan ibadah wajib dan sunnah harian secara tartil, fasih, serta tertib sesuai sunnah' }
        ];

        // Dummy Data untuk Kelas IX sebagai contoh awal jika database kosong
        function initDummyData() {
            state.mapel = [...defaultMapel];
            state.siswa['IX'] = [
                { id: 'sis-1', nama: 'Adil Pranata', nisn: '0089123001', tempatLahir: 'Tanjung Siram', tglLahir: '14 Januari 2008', ortu: 'Aris Pranata' },
                { id: 'sis-2', nama: 'Alif Alfiansyah', nisn: '0089123002', tempatLahir: 'Tanjung Siram', tglLahir: '21 Februari 2008', ortu: 'Ahmad Alfian' },
                { id: 'sis-3', nama: 'Amnahwiyah Harahap', nisn: '0089123003', tempatLahir: 'Rantau Prapat', tglLahir: '05 Mei 2008', ortu: 'Rahman Harahap' }
            ];

            const defaultGrades = [
                { sId: 'sis-1', g: [93, 90, 88, 92, 90, 92, 80, 89, 86, 90, 92, 89, 90, 90, 86] },
                { sId: 'sis-2', g: [93, 90, 88, 91, 90, 94, 80, 89, 86, 95, 95, 90, 95, 92, 90] },
                { sId: 'sis-3', g: [97, 92, 90, 98, 95, 98, 86, 89, 92, 97, 96, 96, 97, 94, 92] }
            ];

            defaultGrades.forEach(studentGrades => {
                state.mapel.forEach((m, idx) => {
                    const key = `${studentGrades.sId}_1_${m.id}`;
                    state.nilai[key] = { p: studentGrades.g[idx], k: studentGrades.g[idx] };
                });
                state.catatan[`${studentGrades.sId}_1`] = {
                    teks: 'Sangat baik, harap pertahankan dan kembangkan terus prestasimu demi menggapai cita-cita.',
                    keputusan: 'NAIK KE KELAS IX'
                };
                state.ekskul[`${studentGrades.sId}_1`] = {
                    e1Nama: 'Pramuka', e1Nilai: 'A', e1Ket: 'Aktif, disiplin, dan memiliki kepemimpinan yang baik.',
                    e2Nama: 'UKS', e2Nilai: 'B', e2Ket: 'Cukup berpartisipasi aktif dalam kegiatan kesehatan sekolah.'
                };
                state.absensi[`${studentGrades.sId}_1`] = { sakit: 0, izin: 1, alpa: 0 };
            });
        }

        // --- CLOUD SINKRONISASI & FALLBACK SYSTEM ---
        async function connectToDatabase() {
            // Berikan jeda toleransi bagi ES Module Firebase SDK untuk bersiap
            for (let i = 0; i < 20; i++) {
                if (window.firebaseLoaded) break;
                await new Promise(r => setTimeout(r, 100));
            }

            const loader = document.getElementById('cloud-loader');
            const badge = document.getElementById('cloud-badge');

            if (window.firebaseLoaded) {
                try {
                    const auth = window.firebaseAuth;
                    const db = window.firebaseDb;
                    const appId = window.firebaseAppId;

                    // Langkah 1: Autentikasi Cloud Terlebih Dahulu (RULE 3)
                    if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
                        await window.signInWithCustomToken(auth, __initial_auth_token);
                    } else {
                        await window.signInAnonymously(auth);
                    }

                    if (!auth.currentUser) throw new Error("Sesi login cloud tidak sah.");

                    isCloudMode = true;

                    // Mengatur Desain Badge di Header
                    if (badge) {
                        badge.className = "px-2 py-0.5 rounded-full text-[8px] bg-brand-500/20 text-brand-300 border border-brand-500/30 flex items-center gap-1 font-bold";
                        badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-brand-400"></span> Cloud Terhubung`;
                    }

                    // Jalur Dokumen Sesuai RULE 1 (Aman dari Firebase Permission Errors)
                    const stateDocRef = window.firebaseDoc(db, 'artifacts', appId, 'public', 'data', 'school_data', 'erapor_v4_state');

                    // Ambil Snapshot Waktu Nyata (Real-time Listener)
                    window.firebaseOnSnapshot(stateDocRef, (docSnap) => {
                        if (docSnap.exists()) {
                            state = docSnap.data();
                        } else {
                            // Inisialisasi Cloud State Baru saat pertama kali dijalankan
                            initDummyData();
                            window.firebaseSetDoc(stateDocRef, state);
                        }
                        
                        // Tutup Overlay Loading Cloud
                        if (loader) loader.classList.add('hidden');
                        checkLoginSession();
                    }, (err) => {
                        console.error("Gagal mendapatkan snapshot cloud:", err);
                        fallbackToLocal();
                    });

                } catch (e) {
                    console.error("Firebase auth/conn failed:", e);
                    fallbackToLocal();
                }
            } else {
                fallbackToLocal();
            }
        }

        function fallbackToLocal() {
            const loader = document.getElementById('cloud-loader');
            const badge = document.getElementById('cloud-badge');

            isCloudMode = false;
            if (loader) loader.classList.add('hidden');
            if (badge) {
                badge.className = "px-2 py-0.5 rounded-full text-[8px] bg-amber-500/20 text-amber-300 border border-amber-500/30 flex items-center gap-1 font-bold";
                badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-amber-400"></span> Penyimpanan Lokal`;
            }

            // Muat dari penyimpanan lokal localStorage
            const savedData = localStorage.getItem('erapor_data_v1');
            if (savedData) {
                try {
                    state = JSON.parse(savedData);
                } catch (e) {
                    initDummyData();
                }
            } else {
                initDummyData();
                localStorage.setItem('erapor_data_v1', JSON.stringify(state));
            }
            checkLoginSession();
        }

        function saveAllData() {
            if (isCloudMode && window.firebaseLoaded && window.firebaseAuth.currentUser) {
                const db = window.firebaseDb;
                const appId = window.firebaseAppId;
                const stateDocRef = window.firebaseDoc(db, 'artifacts', appId, 'public', 'data', 'school_data', 'erapor_v4_state');
                
                // Simpan langsung ke cloud (Setiap Wali Kelas di perangkat lain akan ter-update otomatis)
                window.firebaseSetDoc(stateDocRef, state)
                    .then(() => {
                        console.log("Cloud Database berhasil diperbarui.");
                    })
                    .catch(err => {
                        console.error("Gagal menyimpan ke Cloud:", err);
                        showToast('Gagal Cloud', 'Gagal sinkronisasi data ke Cloud.', 'error');
                    });
            } else {
                // Cadangan lokal jika berjalan offline
                localStorage.setItem('erapor_data_v1', JSON.stringify(state));
            }
            syncConfigUI();
        }

        // Update Configuration Values
        function updateConfig(field, val) {
            state.config[field] = val;
            saveAllData();
            // Real-time update documents if visible
            const activeTab = document.querySelector('.tab-content:not(.hidden)');
            if (activeTab && activeTab.id === 'tab-raport') {
                renderRaportDoc();
            } else if (activeTab && activeTab.id === 'tab-skhu') {
                renderSKHUDoc();
            }
        }

        function syncConfigUI() {
            const inputs = {
                'config-tempat': 'tempat',
                'config-no-surat': 'noSurat',
                'config-tgl-rapor': 'tglRapor',
                'config-tgl-skhu': 'tglSkhu',
                'config-wali-nama': 'waliNama',
                'config-kepsek-nama': 'kepsekNama'
            };
            for (let [id, field] of Object.entries(inputs)) {
                const el = document.getElementById(id);
                if (el) el.value = state.config[field] || '';
            }
        }

        // --- AUTHENTICATION & LOGIN LOGIC ---
        const authUsers = {
            'admin': { role: 'ADMIN', pass: 'admin' },
            'guru7': { role: 'VII', pass: 'kelas7' },
            'guru8': { role: 'VIII', pass: 'kelas8' },
            'guru9': { role: 'IX', pass: 'sandiganteng' }
        };

        function checkLoginSession() {
            const loggedInRole = localStorage.getItem('erapor_role');
            
            const loginScreen = document.getElementById('login-screen');
            const appHeader = document.getElementById('app-header');
            const appContent = document.getElementById('app-content');

            if (loggedInRole) {
                if (loginScreen) loginScreen.classList.add('hidden');
                if (appHeader) appHeader.classList.remove('hidden');
                if (appContent) {
                    appContent.classList.remove('hidden');
                    appContent.classList.add('flex');
                }

                // Atur pembatasan tombol kelas berdasar hak login
                applyRoleRestrictions(loggedInRole);
                
                if (loggedInRole !== 'ADMIN') {
                    changeActiveKelas(loggedInRole);
                } else {
                    changeActiveKelas('IX');
                }
                
                changeActiveSemester(1);
                switchTab('dashboard');
            } else {
                if (loginScreen) loginScreen.classList.remove('hidden');
                if (appHeader) appHeader.classList.add('hidden');
                if (appContent) {
                    appContent.classList.add('hidden');
                    appContent.classList.remove('flex');
                }
            }
        }

        function handleLogin(e) {
            e.preventDefault();
            const user = document.getElementById('login-user').value;
            const pass = document.getElementById('login-pass').value;
            const errorText = document.getElementById('login-error');

            if (authUsers[user] && authUsers[user].pass === pass) {
                localStorage.setItem('erapor_role', authUsers[user].role);
                if (errorText) errorText.classList.add('hidden');
                document.getElementById('login-pass').value = '';
                checkLoginSession();
            } else {
                if (errorText) errorText.classList.remove('hidden');
            }
        }

        function handleLogout() {
            localStorage.removeItem('erapor_role');
            location.reload();
        }

        function applyRoleRestrictions(role) {
            const btn7 = document.getElementById('btn-class-VII');
            const btn8 = document.getElementById('btn-class-VIII');
            const btn9 = document.getElementById('btn-class-IX');

            const classDisplayMap = {
                'VII': 'Wali Kelas VII',
                'VIII': 'Wali Kelas VIII',
                'IX': 'Wali Kelas IX',
                'ADMIN': 'Administrator (Akses Penuh)'
            };

            const headerKelasWali = document.getElementById('header-kelas-wali');
            const headerNamaWali = document.getElementById('header-nama-wali');

            if (headerKelasWali) headerKelasWali.innerText = classDisplayMap[role] || 'Wali Kelas';
            if (headerNamaWali) headerNamaWali.innerText = role === 'ADMIN' ? 'Administrator' : `Wali Kelas ${role}`;

            if (role === 'ADMIN') {
                if (btn7) btn7.classList.remove('hidden');
                if (btn8) btn8.classList.remove('hidden');
                if (btn9) btn9.classList.remove('hidden');
            } else {
                if (btn7) btn7.classList.add('hidden');
                if (btn8) btn8.classList.add('hidden');
                if (btn9) btn9.classList.add('hidden');
                
                if (role === 'VII' && btn7) btn7.classList.remove('hidden');
                if (role === 'VIII' && btn8) btn8.classList.remove('hidden');
                if (role === 'IX' && btn9) btn9.classList.remove('hidden');
            }
        }

        // --- NAVIGATION CONTROLLER ---
        function switchTab(tabId) {
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.add('hidden'));
            const activeTab = document.getElementById(`tab-${tabId}`);
            if (activeTab) activeTab.classList.remove('hidden');

            document.querySelectorAll('.nav-btn').forEach(btn => {
                btn.classList.remove('bg-brand-50', 'text-brand-700', 'border-r-4', 'border-brand-500');
            });
            const activeNav = document.getElementById(`nav-${tabId}`);
            if (activeNav) {
                activeNav.classList.add('bg-brand-50', 'text-brand-700', 'border-r-4', 'border-brand-500');
            }

            if (tabId === 'siswa') {
                renderSiswa();
            } else if (tabId === 'nilai') {
                populateStudentSelectors();
                loadStudentGrades();
            } else if (tabId === 'raport') {
                populateStudentSelectors();
                renderRaportDoc();
            } else if (tabId === 'skhu') {
                populateStudentSelectors();
                renderSKHUDoc();
            } else if (tabId === 'dashboard') {
                renderDashboard();
            }
        }

        function changeActiveKelas(kelas) {
            state.activeKelas = kelas;
            document.querySelectorAll('[id^="btn-class-"]').forEach(btn => {
                btn.classList.remove('bg-brand-600', 'text-white', 'border-brand-600');
                btn.classList.add('text-slate-500', 'border-slate-200');
            });
            const activeBtn = document.getElementById(`btn-class-${kelas}`);
            if (activeBtn) {
                activeBtn.classList.remove('text-slate-500', 'border-slate-200');
                activeBtn.classList.add('bg-brand-600', 'text-white', 'border-brand-600');
            }
            renderDashboard();
        }

        function changeActiveSemester(semester) {
            state.activeSemester = semester;
            document.querySelectorAll('[id^="btn-semester-"]').forEach(btn => {
                btn.classList.remove('bg-brand-600', 'text-white', 'border-brand-600');
                btn.classList.add('text-slate-500', 'border-slate-200');
            });
            const activeBtn = document.getElementById(`btn-semester-${semester}`);
            if (activeBtn) {
                activeBtn.classList.remove('text-slate-500', 'border-slate-200');
                activeBtn.classList.add('bg-brand-600', 'text-white', 'border-brand-600');
            }
            const semText = document.getElementById('stat-semester-active');
            if (semText) semText.innerText = `Semester ${semester === 1 ? 'Ganjil (1)' : 'Genap (2)'}`;
            renderDashboard();
        }

        // Render Dashboard Stats
        function renderDashboard() {
            const listSiswa = state.siswa[state.activeKelas] || [];
            const countEl = document.getElementById('stat-siswa-total');
            if (countEl) countEl.innerText = `${listSiswa.length} Siswa`;
        }

        // --- MANAJEMEN SISWA CONTROLLER ---
        function renderSiswa() {
            const tbody = document.getElementById('siswa-table-body');
            if (!tbody) return;
            tbody.innerHTML = '';
            const list = state.siswa[state.activeKelas] || [];

            if (list.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="py-8 text-center text-slate-400 font-bold">Tidak ada data siswa di Kelas ${state.activeKelas}.</td></tr>`;
                return;
            }

            list.forEach((s, idx) => {
                tbody.innerHTML += `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="py-3 px-4 font-bold text-slate-900">${s.nama}</td>
                        <td class="py-3 px-4 text-slate-500 font-bold tracking-wider">${s.nisn || '-'}</td>
                        <td class="py-3 px-4 text-slate-500">${s.tempatLahir || '-'}, ${s.tglLahir || '-'}</td>
                        <td class="py-3 px-4 text-slate-500">${s.ortu || '-'}</td>
                        <td class="py-3 px-4 text-center">
                            <div class="flex items-center justify-center gap-2">
                                <button onclick="openEditSiswaModal('${s.id}')" class="w-8 h-8 rounded-lg bg-indigo-50 hover:bg-indigo-100 text-indigo-600 flex items-center justify-center transition" title="Edit Identitas"><i class="fa-solid fa-pen text-xs"></i></button>
                                <button onclick="deleteSiswa('${s.id}')" class="w-8 h-8 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-600 flex items-center justify-center transition" title="Hapus Siswa"><i class="fa-solid fa-trash-can text-xs"></i></button>
                            </div>
                        </td>
                    </tr>
                `;
            });
        }

        function openAddSiswaModal() {
            document.getElementById('siswa-modal-title').innerText = 'Tambah Siswa Baru';
            document.getElementById('siswa-form-id').value = '';
            document.getElementById('siswa-form-nama').value = '';
            document.getElementById('siswa-form-nisn').value = '';
            document.getElementById('siswa-form-tempat').value = '';
            document.getElementById('siswa-form-tanggal').value = '';
            document.getElementById('siswa-form-ortu').value = '';
            const modal = document.getElementById('siswa-modal');
            if (modal) modal.classList.replace('hidden', 'flex');
        }

        function openEditSiswaModal(id) {
            const list = state.siswa[state.activeKelas] || [];
            const s = list.find(item => item.id === id);
            if (!s) return;

            document.getElementById('siswa-modal-title').innerText = 'Ubah Identitas Siswa';
            document.getElementById('siswa-form-id').value = s.id;
            document.getElementById('siswa-form-nama').value = s.nama;
            document.getElementById('siswa-form-nisn').value = s.nisn || '';
            document.getElementById('siswa-form-tempat').value = s.tempatLahir || '';
            document.getElementById('siswa-form-tanggal').value = s.tglLahir || '';
            document.getElementById('siswa-form-ortu').value = s.ortu || '';
            
            const modal = document.getElementById('siswa-modal');
            if (modal) modal.classList.replace('hidden', 'flex');
        }

        function closeSiswaModal() {
            const modal = document.getElementById('siswa-modal');
            if (modal) modal.classList.replace('flex', 'hidden');
        }

        function handleSiswaSubmit(e) {
            e.preventDefault();
            const id = document.getElementById('siswa-form-id').value;
            const nama = document.getElementById('siswa-form-nama').value.trim();
            const nisn = document.getElementById('siswa-form-nisn').value.trim();
            const tempat = document.getElementById('siswa-form-tempat').value.trim();
            const tanggal = document.getElementById('siswa-form-tanggal').value.trim();
            const ortu = document.getElementById('siswa-form-ortu').value.trim();

            if (!nama || !nisn) return;

            const list = state.siswa[state.activeKelas] || [];

            if (id) {
                // Update
                const idx = list.findIndex(s => s.id === id);
                if (idx !== -1) {
                    list[idx] = { ...list[idx], nama, nisn, tempatLahir: tempat, tglLahir: tanggal, ortu };
                    showToast('Berhasil Diperbarui', 'Data identitas siswa telah di-update.');
                }
            } else {
                // Create
                const newSiswa = {
                    id: 'sis-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
                    nama,
                    nisn,
                    tempatLahir: tempat,
                    tglLahir: tanggal,
                    ortu
                };
                list.push(newSiswa);
                showToast('Siswa Ditambahkan', 'Siswa baru berhasil terdaftar.');
            }

            saveAllData();
            closeSiswaModal();
            renderSiswa();
        }

        function deleteSiswa(id) {
            if (!confirm('Apakah Anda yakin ingin menghapus data siswa ini beserta seluruh nilai akademisnya? Tindakan ini permanen.')) return;
            const list = state.siswa[state.activeKelas] || [];
            const idx = list.findIndex(s => s.id === id);
            if (idx !== -1) {
                list.splice(idx, 1);
                saveAllData();
                renderSiswa();
                showToast('Hapus Selesai', 'Siswa telah terhapus dari sistem.');
            }
        }

        // --- EXCEL SPREADSHEET OPERATIONS (MATA PELAJARAN NILAI) ---
        function triggerExcelImportTabel() {
            const input = document.getElementById('excel-tabel-input');
            if (input) input.click();
        }

        function importDariTabel(e) {
            const files = e.target.files;
            if (files.length === 0) return;

            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

                    const mapelMapping = {
                        'senbud': 'SBD', 'seni budaya': 'SBD',
                        'pkn': 'PPKN', 'pancasila': 'PPKN',
                        'b.indo': 'BIND', 'bahasa indonesia': 'BIND',
                        'mm': 'MTK', 'matematika': 'MTK',
                        'ipa': 'IPA',
                        'ips': 'IPS',
                        'b.inggris': 'BING', 'bahasa inggris': 'BING',
                        'pjok': 'PJOK',
                        'b.arab': 'BARB', 'bahasa arab': 'BARB',
                        'prakarya': 'PRK',
                        'tik': 'TIK',
                        'aa': 'AA', 'akidah akhlak': 'AA',
                        'pai': 'PAI',
                        'ski': 'SKI',
                        'p.ibadah': 'PI', 'praktek ibadah': 'PI'
                    };

                    let addedStudents = 0;
                    let updatedGrades = 0;

                    jsonData.forEach(row => {
                        let namaSiswa = "";
                        for (let key in row) {
                            if (key.trim().toLowerCase() === 'nama') {
                                namaSiswa = String(row[key]);
                                break;
                            }
                        }
                        if (!namaSiswa || namaSiswa.trim() === '') return;

                        let student = state.siswa[state.activeKelas].find(s => s.nama.trim().toLowerCase() === namaSiswa.trim().toLowerCase());
                        
                        if (!student) {
                            const newId = 'sis-' + Date.now() + Math.floor(Math.random() * 1000);
                            student = {
                                id: newId,
                                nisn: '-',
                                nama: namaSiswa,
                                tempatLahir: '-',
                                tglLahir: '',
                                ortu: '-'
                            };
                            state.siswa[state.activeKelas].push(student);
                            addedStudents++;
                        }

                        for (let key in row) {
                            if (key.trim().toLowerCase() === 'nama') continue;
                            
                            const normalizedKey = key.trim().toLowerCase();
                            const mapelKode = mapelMapping[normalizedKey];

                            if (mapelKode) {
                                const mapel = state.mapel.find(m => m.kode === mapelKode);
                                if (mapel) {
                                    const scoreStr = String(row[key]).replace(',', '.');
                                    const score = parseFloat(scoreStr);
                                    if (!isNaN(score)) {
                                        const sKey = `${student.id}_${state.activeSemester}_${mapel.id}`;
                                        state.nilai[sKey] = { p: score, k: score };
                                        updatedGrades++;
                                    }
                                }
                            }
                        }
                    });

                    saveAllData();
                    showToast('Impor Tabel Selesai', `Berhasil! ${addedStudents} siswa baru dibuat & ${updatedGrades} nilai mapel disinkronkan.`);
                    populateStudentSelectors();
                    renderSiswa();
                    renderDashboard();
                } catch (err) {
                    console.error(err);
                    showToast('Gagal Impor', 'Format berkas Excel tidak dapat diproses.', 'error');
                }
            };
            reader.readAsArrayBuffer(files[0]);
            document.getElementById('excel-tabel-input').value = '';
        }

        function exportTabelExcel() {
            const listSiswa = state.siswa[state.activeKelas] || [];
            if (listSiswa.length === 0) {
                showToast('Kosong', `Tidak ada data siswa di kelas ${state.activeKelas} untuk diekspor.`, 'error');
                return;
            }

            const exportData = [];
            const mapelCols = [
                { key: 'senbud', kode: 'SBD' },
                { key: 'pkn', kode: 'PPKN' },
                { key: 'b.indo', kode: 'BIND' },
                { key: 'MM', kode: 'MTK' },
                { key: 'IPA', kode: 'IPA' },
                { key: 'IPS', kode: 'IPS' },
                { key: 'b.inggris', kode: 'BING' },
                { key: 'pjok', kode: 'PJOK' },
                { key: 'b.arab', kode: 'BARB' },
                { key: 'Prakarya', kode: 'PRK' },
                { key: 'TIK', kode: 'TIK' },
                { key: 'AA', kode: 'AA' },
                { key: 'PAI', kode: 'PAI' },
                { key: 'SKI', kode: 'SKI' },
                { key: 'P.ibadah', kode: 'PI' }
            ];

            listSiswa.forEach(s => {
                const row = { Nama: s.nama };
                mapelCols.forEach(col => {
                    const mObj = state.mapel.find(m => m.kode === col.kode);
                    if (mObj) {
                        const sKey = `${s.id}_${state.activeSemester}_${mObj.id}`;
                        const gradeObj = state.nilai[sKey];
                        row[col.key] = (gradeObj && gradeObj.p > 0) ? gradeObj.p : '';
                    } else {
                        row[col.key] = '';
                    }
                });
                exportData.push(row);
            });

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(exportData);
            const wscols = [{wch: 25}];
            for(let i=0; i<15; i++) wscols.push({wch: 10});
            ws['!cols'] = wscols;

            XLSX.utils.book_append_sheet(wb, ws, "Nilai Siswa");
            XLSX.writeFile(wb, `Tabel_Nilai_Kelas_${state.activeKelas}_Smt_${state.activeSemester}.xlsx`);
            showToast('Ekspor Tabel Berhasil', 'Template daftar nilai terunduh.');
        }

        // --- EXCEL SPREADSHEET OPERATIONS (DATA SISWA) ---
        function triggerSiswaImport() {
            const input = document.getElementById('excel-siswa-input');
            if (input) input.click();
        }

        function exportSiswaExcel() {
            const listSiswa = state.siswa[state.activeKelas] || [];
            if (listSiswa.length === 0) {
                showToast('Kosong', `Tidak ada data siswa di kelas ${state.activeKelas} untuk diekspor.`, 'error');
                return;
            }

            const exportData = listSiswa.map(s => ({
                'Nama': s.nama,
                'NISN': s.nisn || '',
                'Tempat Lahir': s.tempatLahir || '',
                'Tanggal Lahir': s.tglLahir || '',
                'Orang Tua': s.ortu || ''
            }));

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(exportData);
            ws['!cols'] = [{wch: 25}, {wch: 15}, {wch: 18}, {wch: 18}, {wch: 25}];

            XLSX.utils.book_append_sheet(wb, ws, "Data Siswa");
            XLSX.writeFile(wb, `Data_Siswa_Kelas_${state.activeKelas}.xlsx`);
            showToast('Ekspor Siswa Berhasil', 'Data siswa berhasil diunduh ke Excel.');
        }

        function importSiswaExcel(e) {
            const files = e.target.files;
            if (files.length === 0) return;

            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

                    let added = 0;
                    let updated = 0;

                    jsonData.forEach(row => {
                        let nama = "";
                        let nisn = "";
                        let tempat = "";
                        let tanggal = "";
                        let ortu = "";

                        for (let key in row) {
                            const cleanKey = key.trim().toLowerCase();
                            if (cleanKey === 'nama' || cleanKey === 'nama lengkap') nama = String(row[key]).trim();
                            if (cleanKey === 'nisn') nisn = String(row[key]).trim();
                            if (cleanKey === 'tempat lahir' || cleanKey === 'tempat') tempat = String(row[key]).trim();
                            if (cleanKey === 'tanggal lahir' || cleanKey === 'tanggal') tanggal = String(row[key]).trim();
                            if (cleanKey === 'orang tua' || cleanKey === 'ortu' || cleanKey === 'nama orang tua') ortu = String(row[key]).trim();
                        }

                        if (!nama) return;

                        let student = state.siswa[state.activeKelas].find(s => s.nama.toLowerCase() === nama.toLowerCase());

                        if (student) {
                            student.nisn = nisn || student.nisn;
                            student.tempatLahir = tempat || student.tempatLahir;
                            student.tglLahir = tanggal || student.tglLahir;
                            student.ortu = ortu || student.ortu;
                            updated++;
                        } else {
                            const newId = 'sis-' + Date.now() + Math.floor(Math.random() * 1000);
                            state.siswa[state.activeKelas].push({
                                id: newId,
                                nama,
                                nisn: nisn || '-',
                                tempatLahir: tempat || '-',
                                tglLahir: tanggal || '-',
                                ortu: ortu || '-'
                            });
                            added++;
                        }
                    });

                    saveAllData();
                    showToast('Impor Siswa Selesai', `Berhasil! Mendaftarkan ${added} siswa baru & memperbarui ${updated} profil.`);
                    populateStudentSelectors();
                    renderSiswa();
                    renderDashboard();
                } catch (err) {
                    console.error(err);
                    showToast('Gagal Impor', 'Format berkas Excel tidak dapat diproses.', 'error');
                }
            };
            reader.readAsArrayBuffer(files[0]);
            document.getElementById('excel-siswa-input').value = '';
        }

        // --- EXCEL SPREADSHEET OPERATIONS (EKSKUL & ABSENSI) ---
        function triggerEkskulImport() {
            const input = document.getElementById('excel-ekskul-input');
            if (input) input.click();
        }

        function exportEkskulExcel() {
            const listSiswa = state.siswa[state.activeKelas] || [];
            if (listSiswa.length === 0) {
                showToast('Kosong', `Tidak ada data siswa di kelas ${state.activeKelas} untuk diekspor.`, 'error');
                return;
            }

            const exportData = listSiswa.map(s => {
                const key = `${s.id}_${state.activeSemester}`;
                const eks = state.ekskul[key] || { e1Nama: '', e1Nilai: '', e1Ket: '', e2Nama: '', e2Nilai: '', e2Ket: '' };
                const abs = state.absensi[key] || { sakit: 0, izin: 0, alpa: 0 };

                return {
                    'Nama': s.nama,
                    'Ekskul 1 Nama': eks.e1Nama || '',
                    'Ekskul 1 Nilai': eks.e1Nilai || '',
                    'Ekskul 1 Ket': eks.e1Ket || '',
                    'Ekskul 2 Nama': eks.e2Nama || '',
                    'Ekskul 2 Nilai': eks.e2Nilai || '',
                    'Ekskul 2 Ket': eks.e2Ket || '',
                    'Sakit': abs.sakit || 0,
                    'Izin': abs.izin || 0,
                    'Alpa': abs.alpa || 0
                };
            });

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(exportData);
            ws['!cols'] = [
                {wch: 25}, {wch: 15}, {wch: 12}, {wch: 25}, 
                {wch: 15}, {wch: 12}, {wch: 25}, 
                {wch: 8}, {wch: 8}, {wch: 8}
            ];

            XLSX.utils.book_append_sheet(wb, ws, "Ekskul & Absensi");
            XLSX.writeFile(wb, `Ekskul_Absen_Kelas_${state.activeKelas}_Smt_${state.activeSemester}.xlsx`);
            showToast('Ekspor Berhasil', 'Template Ekskul & Absensi terunduh.');
        }

        function importEkskulExcel(e) {
            const files = e.target.files;
            if (files.length === 0) return;

            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

                    let updated = 0;

                    jsonData.forEach(row => {
                        let nama = "";
                        let e1Nama = "", e1Nilai = "", e1Ket = "";
                        let e2Nama = "", e2Nilai = "", e2Ket = "";
                        let sakit = 0, izin = 0, alpa = 0;

                        for (let key in row) {
                            const cleanKey = key.trim().toLowerCase();
                            if (cleanKey === 'nama') nama = String(row[key]).trim();
                            if (cleanKey === 'ekskul 1 nama') e1Nama = String(row[key]).trim();
                            if (cleanKey === 'ekskul 1 nilai') e1Nilai = String(row[key]).trim();
                            if (cleanKey === 'ekskul 1 ket') e1Ket = String(row[key]).trim();
                            if (cleanKey === 'ekskul 2 nama') e2Nama = String(row[key]).trim();
                            if (cleanKey === 'ekskul 2 nilai') e2Nilai = String(row[key]).trim();
                            if (cleanKey === 'ekskul 2 ket') e2Ket = String(row[key]).trim();
                            if (cleanKey === 'sakit') sakit = parseInt(row[key]) || 0;
                            if (cleanKey === 'izin') izin = parseInt(row[key]) || 0;
                            if (cleanKey === 'alpa') alpa = parseInt(row[key]) || 0;
                        }

                        if (!nama) return;

                        let student = state.siswa[state.activeKelas].find(s => s.nama.toLowerCase() === nama.toLowerCase());

                        if (student) {
                            const key = `${student.id}_${state.activeSemester}`;
                            state.ekskul[key] = { e1Nama, e1Nilai, e1Ket, e2Nama, e2Nilai, e2Ket };
                            state.absensi[key] = { sakit, izin, alpa };
                            updated++;
                        }
                    });

                    saveAllData();
                    showToast('Impor Ekskul Selesai', `Berhasil menyinkronkan data Ekskul & Absensi untuk ${updated} siswa.`);
                    loadStudentGrades();
                    renderDashboard();
                } catch (err) {
                    console.error(err);
                    showToast('Gagal Impor', 'Format berkas Excel tidak dapat diproses.', 'error');
                }
            };
            reader.readAsArrayBuffer(files[0]);
            document.getElementById('excel-ekskul-input').value = '';
        }

        // --- GLOBAL JSON BACKUP ---
        function exportJSONBackup() {
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(state));
            const dlAnchorElem = document.createElement('a');
            dlAnchorElem.setAttribute("href", dataStr);
            dlAnchorElem.setAttribute("download", `AlFalah_ERapor_Backup_${Date.now()}.json`);
            dlAnchorElem.click();
            showToast('Cadangan Selesai', 'Backup JSON berhasil diunduh.');
        }

        function triggerJSONImport() {
            const input = document.getElementById('json-backup-input');
            if (input) input.click();
        }

        function importJSONBackup(e) {
            const files = e.target.files;
            if (files.length === 0) return;

            const file = files[0];
            const reader = new FileReader();

            reader.onload = function(evt) {
                try {
                    const importedState = JSON.parse(evt.target.result);
                    
                    // Validasi dasar untuk memastikan file adalah backup e-raport yang valid
                    if (importedState && importedState.siswa && importedState.mapel) {
                        state = importedState; // Timpa state saat ini dengan backup
                        saveAllData(); // Simpan ke Cloud & LocalStorage
                        syncConfigUI(); // Update UI form konfigurasi
                        
                        // Segarkan antarmuka
                        populateStudentSelectors();
                        renderDashboard();
                        
                        // Render ulang tab yang mungkin sedang aktif
                        const activeTab = document.querySelector('.tab-content:not(.hidden)');
                        if (activeTab) {
                            if (activeTab.id === 'tab-siswa') renderSiswa();
                            if (activeTab.id === 'tab-nilai') loadStudentGrades();
                            if (activeTab.id === 'tab-raport') renderRaportDoc();
                            if (activeTab.id === 'tab-skhu') renderSKHUDoc();
                        }
                        
                        showToast('Pemulihan Berhasil', 'Data cadangan JSON berhasil diimpor dan disinkronkan.');
                    } else {
                        throw new Error("Struktur JSON tidak dikenali.");
                    }
                } catch (err) {
                    console.error(err);
                    showToast('Gagal Impor', 'Berkas JSON rusak atau bukan file cadangan yang valid.', 'error');
                }
            };
            
            reader.readAsText(file);
            document.getElementById('json-backup-input').value = ''; // Reset input agar bisa import file yang sama berulang kali
        }

        // --- INPUT GRADES CONTROLLER ---
        function populateStudentSelectors() {
            const list = state.siswa[state.activeKelas] || [];
            const ids = ['nilai-student-selector', 'raport-student-selector', 'skhu-student-selector'];
            
            ids.forEach(id => {
                const el = document.getElementById(id);
                if (!el) return;
                const prevVal = el.value;
                el.innerHTML = '';
                
                if (list.length === 0) {
                    el.innerHTML = '<option value="">(Belum Ada Siswa)</option>';
                    return;
                }
                
                list.forEach(s => {
                    el.innerHTML += `<option value="${s.id}">${s.nama}</option>`;
                });
                
                if (prevVal && list.some(item => item.id === prevVal)) {
                    el.value = prevVal;
                }
            });
        }

        function loadStudentGrades() {
            const selector = document.getElementById('nilai-student-selector');
            const area = document.getElementById('nilai-input-area');
            if (!selector || !area) return;

            const studentId = selector.value;
            if (!studentId) {
                area.classList.add('hidden');
                return;
            }

            area.classList.remove('hidden');
            area.classList.add('flex');

            const grid = document.getElementById('grades-input-grid');
            grid.innerHTML = '';

            state.mapel.forEach(m => {
                const sKey = `${studentId}_${state.activeSemester}_${m.id}`;
                const scores = state.nilai[sKey] || { p: '', k: '' };

                grid.innerHTML += `
                    <div class="p-4 bg-slate-50 rounded-2xl border border-slate-100 flex flex-col gap-3">
                        <div class="flex items-center justify-between">
                            <span class="text-xs font-extrabold text-slate-800">${m.nama}</span>
                            <span class="text-[9px] font-bold text-slate-400 bg-white px-2 py-0.5 rounded-full border border-slate-100">${m.kode}</span>
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-[9px] font-extrabold text-slate-400 uppercase mb-1">Pengetahuan (P)</label>
                                <input type="number" min="0" max="100" id="p-${m.id}" value="${scores.p}" placeholder="0-100" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                            </div>
                            <div>
                                <label class="block text-[9px] font-extrabold text-slate-400 uppercase mb-1">Keterampilan (K)</label>
                                <input type="number" min="0" max="100" id="k-${m.id}" value="${scores.k}" placeholder="0-100" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-brand-500">
                            </div>
                        </div>
                    </div>
                `;
            });

            // Load Ekskul & Absensi values
            const key = `${studentId}_${state.activeSemester}`;
            const eks = state.ekskul[key] || { e1Nama: '', e1Nilai: '', e1Ket: '', e2Nama: '', e2Nilai: '', e2Ket: '' };
            const abs = state.absensi[key] || { sakit: 0, izin: 0, alpa: 0 };

            document.getElementById('ekskul1-nama').value = eks.e1Nama || '';
            document.getElementById('ekskul1-nilai').value = eks.e1Nilai || '';
            document.getElementById('ekskul1-ket').value = eks.e1Ket || '';
            document.getElementById('ekskul2-nama').value = eks.e2Nama || '';
            document.getElementById('ekskul2-nilai').value = eks.e2Nilai || '';
            document.getElementById('ekskul2-ket').value = eks.e2Ket || '';

            document.getElementById('absen-sakit').value = abs.sakit !== undefined ? abs.sakit : 0;
            document.getElementById('absen-izin').value = abs.izin !== undefined ? abs.izin : 0;
            document.getElementById('absen-alpa').value = abs.alpa !== undefined ? abs.alpa : 0;

            const catKey = `${studentId}_${state.activeSemester}`;
            const meta = state.catatan[catKey] || { teks: '', keputusan: '' };
            document.getElementById('catatan-keputusan').value = meta.keputusan || '';
            document.getElementById('catatan-teks').value = meta.teks || '';
        }

        function saveAllGrades() {
            const studentId = document.getElementById('nilai-student-selector').value;
            if (!studentId) return;

            // Save Grades
            state.mapel.forEach(m => {
                const pVal = parseFloat(document.getElementById(`p-${m.id}`).value);
                const kVal = parseFloat(document.getElementById(`k-${m.id}`).value);
                const sKey = `${studentId}_${state.activeSemester}_${m.id}`;
                state.nilai[sKey] = {
                    p: isNaN(pVal) ? 0 : pVal,
                    k: isNaN(kVal) ? 0 : kVal
                };
            });

            // Save Ekskul & Absensi
            const key = `${studentId}_${state.activeSemester}`;
            state.ekskul[key] = {
                e1Nama: document.getElementById('ekskul1-nama').value.trim(),
                e1Nilai: document.getElementById('ekskul1-nilai').value.trim(),
                e1Ket: document.getElementById('ekskul1-ket').value.trim(),
                e2Nama: document.getElementById('ekskul2-nama').value.trim(),
                e2Nilai: document.getElementById('ekskul2-nilai').value.trim(),
                e2Ket: document.getElementById('ekskul2-ket').value.trim()
            };

            state.absensi[key] = {
                sakit: parseInt(document.getElementById('absen-sakit').value) || 0,
                izin: parseInt(document.getElementById('absen-izin').value) || 0,
                alpa: parseInt(document.getElementById('absen-alpa').value) || 0
            };

            const catKey = `${studentId}_${state.activeSemester}`;
            state.catatan[catKey] = {
                teks: document.getElementById('catatan-teks').value.trim(),
                keputusan: document.getElementById('catatan-keputusan').value.trim()
            };

            saveAllData();
            showToast('Tersimpan', 'Nilai akademik, ekskul, absensi & catatan motivasi berhasil disimpan.');
        }

        // --- AUTOMATIC DESCRIPTIONS GENERATOR (Kemendikbud Standard) ---
        function generateDescription(subject, score, aspect) {
            if (!score || score < 1) return '-';
            const subjectDesc = aspect === 'p' ? subject.descP : subject.descK;
            if (score >= 93) {
                return `Sangat baik dan kompeten dalam ${subjectDesc}`;
            } else if (score >= 84) {
                return `Baik dalam menunjukkan penguasaan materi ${subjectDesc}`;
            } else if (score >= 75) {
                return `Cukup mumpuni serta menguasai kompetensi dasar ${subjectDesc}`;
            } else {
                return `Perlu dibimbing dan dimotivasi lebih giat terkait kompetensi ${subjectDesc}`;
            }
        }

        // --- RAPORT DOCUMENT GENERATION ---
        function renderRaportDoc() {
            const sId = document.getElementById('raport-student-selector').value;
            const doc = document.getElementById('raport-document');
            if (!doc) return;

            if (!sId) {
                doc.innerHTML = '<div class="text-center py-20 text-slate-400 font-bold">Pilih nama siswa di atas untuk melihat rapor resmi.</div>';
                return;
            }

            const s = (state.siswa[state.activeKelas] || []).find(item => item.id === sId);
            if (!s) return;

            const catKey = `${sId}_${state.activeSemester}`;
            const catObj = state.catatan[catKey] || { teks: '-', keputusan: '-' };
            const eksObj = state.ekskul[catKey] || { e1Nama: '', e1Nilai: '', e1Ket: '', e2Nama: '', e2Nilai: '', e2Ket: '' };
            const absObj = state.absensi[catKey] || { sakit: 0, izin: 0, alpa: 0 };

            let tableRows = '';
            let descRows = '';

            state.mapel.forEach((m, idx) => {
                const sKey = `${sId}_${state.activeSemester}_${m.id}`;
                const score = state.nilai[sKey] || { p: 0, k: 0 };
                const avg = ((score.p + score.k) / 2).toFixed(0);
                const passStatus = score.p >= 75 ? 'Lulus' : 'Perlu Bimbingan';

                tableRows += `
                    <tr class="border-b border-black text-center">
                        <td class="py-2 px-3 text-left border-r border-black font-semibold">${idx + 1}. ${m.nama}</td>
                        <td class="py-2 px-3 border-r border-black">75</td>
                        <td class="py-2 px-3 border-r border-black font-semibold">${score.p || '-'}</td>
                        <td class="py-2 px-3 border-r border-black font-semibold">${score.k || '-'}</td>
                        <td class="py-2 px-3 border-r border-black font-bold">${avg > 0 ? avg : '-'}</td>
                        <td class="py-2 px-3 font-semibold text-xs">${passStatus}</td>
                    </tr>
                `;

                descRows += `
                    <tr class="border-b border-black align-top">
                        <td class="py-3 px-3 border-r border-black font-bold text-center">${idx + 1}</td>
                        <td class="py-3 px-3 border-r border-black font-bold">${m.nama}</td>
                        <td class="py-3 px-3 border-r border-black text-xs leading-relaxed">
                            <strong>Pengetahuan:</strong> ${generateDescription(m, score.p, 'p')}<br>
                            <strong class="block mt-1">Keterampilan:</strong> ${generateDescription(m, score.k, 'k')}
                        </td>
                    </tr>
                `;
            });

            // Render Ekskul rows
            let ekskulRows = '';
            let hasEkskul = false;
            if (eksObj.e1Nama) {
                hasEkskul = true;
                ekskulRows += `
                    <tr class="border-b border-black text-center text-xs">
                        <td class="py-2 px-3 border-r border-black">1</td>
                        <td class="py-2 px-3 border-r border-black text-left font-semibold">${eksObj.e1Nama}</td>
                        <td class="py-2 px-3 border-r border-black font-bold">${eksObj.e1Nilai || '-'}</td>
                        <td class="py-2 px-3 text-left italic">${eksObj.e1Ket || '-'}</td>
                    </tr>
                `;
            }
            if (eksObj.e2Nama) {
                hasEkskul = true;
                ekskulRows += `
                    <tr class="border-b border-black text-center text-xs">
                        <td class="py-2 px-3 border-r border-black">2</td>
                        <td class="py-2 px-3 border-r border-black text-left font-semibold">${eksObj.e2Nama}</td>
                        <td class="py-2 px-3 border-r border-black font-bold">${eksObj.e2Nilai || '-'}</td>
                        <td class="py-2 px-3 text-left italic">${eksObj.e2Ket || '-'}</td>
                    </tr>
                `;
            }
            if (!hasEkskul) {
                ekskulRows = `
                    <tr class="border-b border-black text-center text-xs">
                        <td colspan="4" class="py-3 text-slate-400 italic">Tidak mengikuti kegiatan ekstrakurikuler</td>
                    </tr>
                `;
            }

            doc.innerHTML = `
                <!-- Kop Surat -->
                <div class="text-center border-b-[3px] border-black pb-4 mb-6">
                    <h1 class="text-lg font-bold uppercase tracking-wide">Yayasan Pendidikan Nurhyat Iman Al-Falah</h1>
                    <h2 class="text-xl font-extrabold uppercase tracking-tight">SMP IT NURHYAT IMAN AL-FALAH</h2>
                    <p class="text-xs italic font-medium">Izin Operasional No: 421.3/2883/PPB-3/2019 - NPSN: 69992520</p>
                    <p class="text-[11px] font-semibold">Tanjung Siram, Kec. Bilah Barat, Kabupaten Labuhanbatu, Sumatera Utara</p>
                </div>

                <div class="text-center mb-6">
                    <h3 class="text-md font-bold uppercase tracking-wider">Laporan Hasil Belajar Siswa (Raport)</h3>
                    <p class="text-xs font-bold">Tahun Pelajaran: 2025/2026 | Semester: ${state.activeSemester === 1 ? '1 (Ganjil)' : '2 (Genap)'}</p>
                </div>

                <!-- Identitas -->
                <table class="w-full text-xs font-semibold mb-6 border-none" style="border: none !important;">
                    <tr style="border: none !important;"><td class="py-1 px-0 w-1/4" style="border: none;">Nama Peserta Didik</td><td class="py-1 px-0 w-4" style="border: none;">:</td><td class="py-1 px-0 font-bold" style="border: none;">${s.nama}</td></tr>
                    <tr style="border: none !important;"><td class="py-1 px-0" style="border: none;">NISN</td><td class="py-1 px-0" style="border: none;">:</td><td class="py-1 px-0 font-bold tracking-wider" style="border: none;">${s.nisn || '-'}</td></tr>
                    <tr style="border: none !important;"><td class="py-1 px-0" style="border: none;">Kelas</td><td class="py-1 px-0" style="border: none;">:</td><td class="py-1 px-0" style="border: none;">${state.activeKelas}</td></tr>
                </table>

                <!-- Tabel Capaian Akademik -->
                <div class="mb-6">
                    <h4 class="font-bold text-xs mb-2">A. Nilai Hasil Belajar</h4>
                    <table class="w-full text-xs border border-black border-collapse">
                        <thead>
                            <tr class="bg-slate-100 border-b border-black text-center font-bold">
                                <th class="py-2 px-3 text-left border-r border-black w-2/5">Mata Pelajaran</th>
                                <th class="py-2 px-3 border-r border-black w-16">KKM</th>
                                <th class="py-2 px-3 border-r border-black w-16">P</th>
                                <th class="py-2 px-3 border-r border-black w-16">K</th>
                                <th class="py-2 px-3 border-r border-black w-16">Rerata</th>
                                <th class="py-2 px-3">Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${tableRows}
                        </tbody>
                    </table>
                </div>

                <!-- Lembar Deskripsi Capaian Kompetensi -->
                <div class="mb-6 page-break-before">
                    <h4 class="font-bold text-xs mb-2">B. Deskripsi Capaian Kompetensi</h4>
                    <table class="w-full text-xs border border-black border-collapse">
                        <thead>
                            <tr class="bg-slate-100 border-b border-black font-bold">
                                <th class="py-2 px-3 border-r border-black w-12 text-center">No</th>
                                <th class="py-2 px-3 border-r border-black w-1/4 text-left">Mata Pelajaran</th>
                                <th class="py-2 px-3 text-left">Deskripsi Kemajuan Belajar</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${descRows}
                        </tbody>
                    </table>
                </div>

                <!-- Ekskul & Absensi -->
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
                    <!-- Ekskul Table -->
                    <div>
                        <h4 class="font-bold text-xs mb-2">C. Kegiatan Ekstrakurikuler</h4>
                        <table class="w-full text-xs border border-black border-collapse">
                            <thead>
                                <tr class="bg-slate-100 border-b border-black font-bold text-center">
                                    <th class="py-2 px-2 border-r border-black w-8">No</th>
                                    <th class="py-2 px-3 border-r border-black text-left">Kegiatan Ekstrakurikuler</th>
                                    <th class="py-2 px-2 border-r border-black w-12">Nilai</th>
                                    <th class="py-2 px-3 text-left">Keterangan</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${ekskulRows}
                            </tbody>
                        </table>
                    </div>

                    <!-- Absensi Table -->
                    <div>
                        <h4 class="font-bold text-xs mb-2">D. Ketidakhadiran (Absensi)</h4>
                        <table class="w-full text-xs border border-black border-collapse">
                            <thead>
                                <tr class="bg-slate-100 border-b border-black font-bold">
                                    <th class="py-2 px-4 text-left border-r border-black">Keadaan Absensi</th>
                                    <th class="py-2 px-4 text-center w-24">Jumlah Hari</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr class="border-b border-black">
                                    <td class="py-2 px-4 border-r border-black font-semibold">1. Sakit</td>
                                    <td class="py-2 px-4 text-center font-bold">${absObj.sakit || 0} Hari</td>
                                </tr>
                                <tr class="border-b border-black">
                                    <td class="py-2 px-4 border-r border-black font-semibold">2. Izin</td>
                                    <td class="py-2 px-4 text-center font-bold">${absObj.izin || 0} Hari</td>
                                </tr>
                                <tr class="border-b border-black">
                                    <td class="py-2 px-4 border-r border-black font-semibold">3. Tanpa Keterangan (Alpa)</td>
                                    <td class="py-2 px-4 text-center font-bold">${absObj.alpa || 0} Hari</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Catatan Akhir & Keputusan -->
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8 pt-4">
                    <div class="border border-black p-4 rounded-xl">
                        <h4 class="font-bold text-xs uppercase mb-1">E. Catatan & Motivasi Wali Kelas:</h4>
                        <p class="text-xs italic font-medium">"${catObj.teks}"</p>
                    </div>
                    <div class="border border-black p-4 rounded-xl flex flex-col justify-center items-center text-center">
                        <h4 class="font-bold text-xs uppercase mb-1">Keputusan Akhir:</h4>
                        <div class="text-sm font-extrabold text-brand-900 bg-brand-50 border border-brand-200 rounded-lg px-4 py-2 mt-1">
                            ${catObj.keputusan || 'BELUM DITENTUKAN'}
                        </div>
                    </div>
                </div>

                <!-- Tanda Tangan -->
                <div class="grid grid-cols-3 text-center text-xs font-semibold mt-10">
                    <div>
                        <p class="mb-16">Orang Tua / Wali Murid,</p>
                        <p class="font-bold underline">_______________________</p>
                    </div>
                    <div>
                        <p class="mb-16">Wali Kelas ${state.activeKelas},</p>
                        <p class="font-bold underline">${state.config.waliNama}</p>
                    </div>
                    <div>
                        <p class="mb-16">${state.config.tempat}, ${state.config.tglRapor}</p>
                        <p class="font-bold underline">${state.config.kepsekNama}</p>
                        <p class="text-[10px]">Kepala Sekolah</p>
                    </div>
                </div>
            `;
        }

        // --- SKHU DOCUMENT GENERATION ---
        function renderSKHUDoc() {
            const sId = document.getElementById('skhu-student-selector').value;
            const doc = document.getElementById('skhu-document');
            if (!doc) return;

            if (!sId) {
                doc.innerHTML = '<div class="text-center py-20 text-slate-400 font-bold">Pilih nama siswa di atas untuk merender surat kelulusan.</div>';
                return;
            }

            const s = (state.siswa[state.activeKelas] || []).find(item => item.id === sId);
            if (!s) return;

            let tableRows = '';
            let totalScore = 0;
            let countMapel = 0;

            state.mapel.forEach((m, idx) => {
                const sKey = `${sId}_${state.activeSemester}_${m.id}`;
                const scoreObj = state.nilai[sKey] || { p: 0, k: 0 };
                const avg = scoreObj.p > 0 ? scoreObj.p : 0;
                totalScore += avg;
                if (avg > 0) countMapel++;

                tableRows += `
                    <tr class="border-b border-black text-center font-semibold">
                        <td class="py-2 px-3 border-r border-black">${idx + 1}</td>
                        <td class="py-2 px-3 border-r border-black text-left">${m.nama}</td>
                        <td class="py-2 px-3 font-bold">${avg > 0 ? avg : '-'}</td>
                    </tr>
                `;
            });

            const finalAvg = countMapel > 0 ? (totalScore / countMapel).toFixed(1) : '-';

            doc.innerHTML = `
                <!-- Kop Resmi -->
                <div class="text-center border-b-[3px] border-black pb-4 mb-6">
                    <h1 class="text-lg font-bold uppercase tracking-wide">Yayasan Pendidikan Nurhyat Iman Al-Falah</h1>
                    <h2 class="text-xl font-extrabold uppercase tracking-tight">SMP IT NURHYAT IMAN AL-FALAH</h2>
                    <p class="text-xs italic font-medium">Izin Operasional No: 421.3/2883/PPB-3/2019 - NPSN: 69992520</p>
                    <p class="text-[11px] font-semibold">Tanjung Siram, Kec. Bilah Barat, Kabupaten Labuhanbatu, Sumatera Utara</p>
                </div>

                <div class="text-center mb-8">
                    <h3 class="text-sm font-bold uppercase tracking-widest underline">SURAT KETERANGAN HASIL UJIAN</h3>
                    <p class="text-xs font-bold">Nomor: ${state.config.noSurat || '-'}</p>
                </div>

                <p class="text-xs mb-4">Yang bertanda tangan di bawah ini, Kepala Sekolah Menengah Pertama Islam Terpadu (SMP IT) Al-Falah Tanjung Siram, dengan ini menerangkan bahwa:</p>

                <!-- Identitas Murid -->
                <table class="w-full text-xs font-semibold mb-6 border-none" style="border: none !important;">
                    <tr style="border: none !important;"><td class="py-1.5 px-0 w-1/4" style="border: none;">Nama Lengkap</td><td class="py-1.5 px-0 w-4" style="border: none;">:</td><td class="py-1.5 px-0 font-bold uppercase" style="border: none;">${s.nama}</td></tr>
                    <tr style="border: none !important;"><td class="py-1.5 px-0" style="border: none;">Tempat, Tanggal Lahir</td><td class="py-1.5 px-0" style="border: none;">:</td><td class="py-1.5 px-0" style="border: none;">${s.tempatLahir || '-'}, ${s.tglLahir || '-'}</td></tr>
                    <tr style="border: none !important;"><td class="py-1.5 px-0" style="border: none;">NISN</td><td class="py-1.5 px-0" style="border: none;">:</td><td class="py-1.5 px-0 font-bold tracking-widest" style="border: none;">${s.nisn || '-'}</td></tr>
                    <tr style="border: none !important;"><td class="py-1.5 px-0" style="border: none;">Nama Orang Tua / Wali</td><td class="py-1.5 px-0" style="border: none;">:</td><td class="py-1.5 px-0" style="border: none;">${s.ortu || '-'}</td></tr>
                    <tr style="border: none !important;"><td class="py-1.5 px-0" style="border: none;">Sekolah Asal</td><td class="py-1.5 px-0" style="border: none;">:</td><td class="py-1.5 px-0" style="border: none;">SMP IT Nurhyat Iman Al-Falah</td></tr>
                </table>

                <p class="text-xs mb-6">Telah mengikuti seluruh rangkaian ujian kelulusan akademis dan dinyatakan lulus dari satuan pendidikan berdasarkan hasil penilaian evaluasi belajar dengan pencapaian nilai sebagai berikut:</p>

                <!-- Nilai SKHU -->
                <table class="w-full text-xs border border-black border-collapse mb-8">
                    <thead>
                        <tr class="bg-slate-100 border-b border-black font-bold">
                            <th class="py-2 px-3 border-r border-black w-16 text-center">No</th>
                            <th class="py-2 px-3 border-r border-black text-left">Mata Pelajaran Kurikulum</th>
                            <th class="py-2 px-3 w-32 text-center">Nilai Ujian Kelulusan</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRows}
                        <tr class="bg-slate-50 border-t border-black font-bold text-center">
                            <td colspan="2" class="py-2 px-3 border-r border-black text-right uppercase">Rata-Rata Nilai Kelulusan</td>
                            <td class="py-2 px-3 text-brand-950 font-extrabold text-sm">${finalAvg}</td>
                        </tr>
                    </tbody>
                </table>

                <!-- Bagian Penutup & Tanda Tangan -->
                <div class="flex justify-between items-start text-xs mt-12 font-semibold">
                    <div class="w-32 h-40 border border-slate-300 rounded flex items-center justify-center text-center p-2 text-[9px] italic text-slate-400">
                        Pas Foto<br>3x4<br>Siswa
                    </div>
                    <div class="text-center">
                        <p class="mb-2">${state.config.tempat}, ${state.config.tglSkhu}</p>
                        <p class="mb-20">Kepala Sekolah SMP IT Al-Falah,</p>
                        <p class="font-bold underline uppercase">${state.config.kepsekNama}</p>
                        <p class="text-[10px]">NIP / NUPTK. -</p>
                    </div>
                </div>
            `;
        }

        // --- GLOBAL TOAST CONTROLLER ---
        function showToast(title, desc, type = 'success') {
            const toast = document.getElementById('toast-notif');
            const icon = document.getElementById('toast-icon');
            const titleEl = document.getElementById('toast-title');
            const descEl = document.getElementById('toast-desc');

            if (!toast) return;

            if (type === 'success') {
                icon.className = 'fa-solid fa-circle-check text-brand-400 text-lg';
            } else {
                icon.className = 'fa-solid fa-triangle-exclamation text-rose-500 text-lg';
            }

            titleEl.innerText = title;
            descEl.innerText = desc;

            toast.classList.replace('hidden', 'flex');
            setTimeout(() => {
                toast.classList.replace('flex', 'hidden');
            }, 3500);
        }

        // Startup actions
        window.onload = function() {
            connectToDatabase();
        }
    </script>
</body>
</html>
