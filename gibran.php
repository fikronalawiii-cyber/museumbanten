<?php
$error = '';
$data_alternatif = [];
$kriteria_terdeteksi = [];
$step1_bobot = [];
$step2_vektor_s = [];
$step3_vektor_v = [];
$s_values_list = []; 

// Fungsi pintar untuk membaca XLSX asli maupun CSV (dengan pembersihan kolom kosong)
function baca_excel_or_csv($file_input_name)
{
    if (!isset($_FILES[$file_input_name])) return false;
    
    $filename = $_FILES[$file_input_name]['tmp_name'];
    $real_name = $_FILES[$file_input_name]['name'];
    $ext = strtolower(pathinfo($real_name, PATHINFO_EXTENSION));
    
    $data = [];

    // 1. JIKA FILE ADALAH CSV
    if ($ext === 'csv') {
        if (($handle = fopen($filename, "r")) !== FALSE) {
            $first_line = fgets($handle);
            $separator = (strpos($first_line, ';') !== false) ? ';' : ',';
            rewind($handle);

            while (($rowData = fgetcsv($handle, 1000, $separator)) !== FALSE) {
                $rowData = array_map('trim', $rowData);
                
                // Hapus kolom pertama jika kosong (efek samping ekspor csv tertentu)
                if (isset($rowData[0]) && $rowData[0] === '') {
                    array_shift($rowData);
                }
                
                if (!empty($rowData) && implode('', $rowData) !== '') {
                    $data[] = $rowData;
                }
            }
            fclose($handle);
        }
        return $data;
    }

    // 2. JIKA FILE ADALAH XLSX ASLI
    if (!class_exists('ZipArchive')) return false;

    $zip = new ZipArchive;
    if ($zip->open($filename) === TRUE) {
        $sharedStrings = [];
        $stringsXml = $zip->getFromName('xl/sharedStrings.xml');
        if ($stringsXml) {
            $xml = simplexml_load_string($stringsXml);
            foreach ($xml->si as $val) $sharedStrings[] = (string)$val->t;
        }

        $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
        if ($sheetXml) {
            $xml = simplexml_load_string($sheetXml);
            foreach ($xml->sheetData->row as $row) {
                $rowData = [];
                foreach ($row->c as $c) {
                    $val = (string)$c->v;
                    $type = (string)$c['t'];
                    if ($type == 's' && isset($sharedStrings[$val])) {
                        $rowData[] = $sharedStrings[$val];
                    } else {
                        $rowData[] = $val;
                    }
                }
                
                // Hapus kolom pertama jika kosong
                if (isset($rowData[0]) && $rowData[0] === '') {
                    array_shift($rowData);
                }
                
                if (!empty($rowData)) $data[] = $rowData;
            }
        }
        $zip->close();
    }
    return $data;
}

// Proses Form
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['file_alternatif']) && isset($_FILES['file_bobot'])) {
    
    // Panggil fungsi menggunakan nama input HTML-nya
    $rows_alt = baca_excel_or_csv('file_alternatif');
    $rows_bobot = baca_excel_or_csv('file_bobot');

    if ($rows_alt === false || $rows_bobot === false) {
        $error = "Gagal memproses file. Pastikan ekstensi ZipArchive aktif di server.";
    } elseif (empty($rows_alt) || empty($rows_bobot)) {
        $error = "File Excel/CSV kosong atau tidak sesuai format.";
    } else {
        
        // 1. Baca Data Bobot (Abaikan header jika berisi teks non-kode)
        $start_row_bobot = (strpos(strtolower($rows_bobot[0][0]), 'kriteria') !== false) ? 1 : 0;
        for ($i = $start_row_bobot; $i < count($rows_bobot); $i++) {
            $r = $rows_bobot[$i];
            if (!empty($r[0])) {
                $kode = trim($r[0]);
                $kriteria_terdeteksi[$kode] = [
                    'nama' => $r[1] ?? $kode,
                    'bobot_awal' => (float)($r[2] ?? 0),
                    'tipe' => strtolower(trim($r[3] ?? 'benefit'))
                ];
            }
        }

        // 2. Baca Data Alternatif
        $header_alt = $rows_alt[0];
        for ($i = 1; $i < count($rows_alt); $i++) {
            $r = $rows_alt[$i];
            if (!empty($r[0])) {
                $item = ['nama' => $r[0]];
                for ($j = 1; $j < count($header_alt); $j++) {
                    $k_kode = trim($header_alt[$j]);
                    if (isset($kriteria_terdeteksi[$k_kode])) {
                        $item[$k_kode] = (float)($r[$j] ?? 0);
                    }
                }
                $data_alternatif[] = $item;
            }
        }

        // --- PERHITUNGAN WP ---
        if (empty($kriteria_terdeteksi) || empty($data_alternatif)) {
            $error = "Struktur data kriteria atau alternatif tidak sinkron. Periksa baris header file Anda.";
        } else {
            $total_w = array_sum(array_column($kriteria_terdeteksi, 'bobot_awal'));
            foreach ($kriteria_terdeteksi as $k => $v) {
                $step1_bobot[$k] = $total_w > 0 ? ($v['bobot_awal'] / $total_w) : 0;
            }

            $total_s_all = 0;
            foreach ($data_alternatif as $index => $alt) {
                $s_val = 1;
                $formula_parts = [];
                foreach ($kriteria_terdeteksi as $k => $v) {
                    $wj = number_format($step1_bobot[$k], 3, '.', '');
                    $pangkat = ($v['tipe'] == 'cost') ? -$wj : $wj;
                    
                    // Validasi pengaman jika nilai alternatif bernilai 0 agar tidak crash di pow()
                    $base_val = isset($alt[$k]) ? $alt[$k] : 1;
                    $s_val *= pow($base_val, $pangkat);
                    
                    $formatted_pangkat = ($pangkat < 0 ? '-' : '') . number_format(abs($pangkat), 3);
                    $formula_parts[] = "(" . $base_val . "<sup>" . $formatted_pangkat . "</sup>)";
                }

                $hasil_s_formatted = number_format($s_val, 3);
                $s_values_list[] = $hasil_s_formatted;

                $step2_vektor_s[] = [
                    'nama' => $alt['nama'],
                    'kode' => 'S' . ($index + 1),
                    'rumus' => implode("", $formula_parts),
                    'hasil' => $s_val
                ];
                $total_s_all += $s_val;
            }

            foreach ($step2_vektor_s as $index => $data_s) {
                $vi_val = $total_s_all > 0 ? ($data_s['hasil'] / $total_s_all) : 0;
                $step3_vektor_v[] = [
                    'nama' => $data_s['nama'],
                    'kode' => 'V' . ($index + 1),
                    'si_val' => $data_s['hasil'],
                    'total_s' => $total_s_all,
                    'hasil' => $vi_val
                ];
            }

            $step3_vektor_v_ranked = $step3_vektor_v;
            usort($step3_vektor_v_ranked, function ($a, $b) {
                return $b['hasil'] <=> $a['hasil'];
            });
        }
    }
}
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SPK WP - Dashboard Perhitungan</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto+Mono:wght@400;700&family=Inter:wght@400;900&display=swap');
        body { font-family: 'Inter', sans-serif; }
        .math { font-family: 'Roboto Mono', monospace; }
        .custom-scroll::-webkit-scrollbar { width: 6px; }
        .custom-scroll::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 10px; }
        .custom-scroll::-webkit-scrollbar-thumb { background: #c7d2fe; border-radius: 10px; }
        .custom-scroll::-webkit-scrollbar-thumb:hover { background: #818cf8; }
    </style>
</head>
<body class="bg-gray-100 text-slate-900">

    <nav class="bg-indigo-950 text-white p-8 shadow-2xl text-center">
        <h1 class="text-3xl font-black uppercase italic tracking-tighter">Weighted Product Engine</h1>
        <p class="text-indigo-400 text-sm mt-1 font-bold">Laporan Perhitungan Lengkap & Presisi</p>
    </nav>

    <main class="max-w-6xl mx-auto p-6">
        
        <section class="bg-white rounded-3xl shadow-xl p-8 mb-12 border border-gray-200">
            <h2 class="text-xl font-black mb-6 text-slate-700 uppercase tracking-tight">📁 Import Berkas Excel</h2>
            <form action="" method="POST" enctype="multipart/form-data" class="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div class="bg-slate-50 p-6 rounded-2xl border-2 border-dashed border-slate-200 hover:border-indigo-400 transition-colors">
                    <label class="block text-xs font-black text-slate-400 uppercase mb-3">1. Bobot Kriteria (.xlsx)</label>
                    <input type="file" name="file_bobot" accept=".xlsx" required class="w-full text-sm">
                </div>
                <div class="bg-slate-50 p-6 rounded-2xl border-2 border-dashed border-slate-200 hover:border-indigo-400 transition-colors">
                    <label class="block text-xs font-black text-slate-400 uppercase mb-3">2. Data Alternatif (.xlsx)</label>
                    <input type="file" name="file_alternatif" accept=".xlsx" required class="w-full text-sm">
                </div>
                <button type="submit" class="md:col-span-2 bg-indigo-600 hover:bg-indigo-800 text-white font-black py-4 rounded-2xl transition-all shadow-xl shadow-indigo-200 uppercase tracking-widest text-lg">Mulai Kalkulasi</button>
            </form>
        </section>

        <?php if (!empty($data_alternatif)): ?>
            
            <div class="mb-12">
                <h2 class="text-2xl font-black mb-6 text-indigo-900 border-l-8 border-indigo-600 pl-4 uppercase italic">Langkah 1: Perbaikan Bobot ($W_j$)</h2>
                <div class="bg-white rounded-3xl shadow-lg border overflow-hidden">
                    <table class="w-full text-sm text-left">
                        <thead class="bg-slate-900 text-white">
                            <tr>
                                <th class="p-5">Kriteria</th>
                                <th class="p-5">Tipe</th>
                                <th class="p-5 text-right text-indigo-400">Normalisasi ($W_j$)</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-100">
                            <?php foreach($kriteria_terdeteksi as $k => $v): ?>
                            <tr>
                                <td class="p-5 font-bold text-slate-700"><?= $k ?> - <?= $v['nama'] ?></td>
                                <td class="p-5"><span class="px-3 py-1 rounded-full text-[10px] font-black uppercase <?= $v['tipe'] == 'cost' ? 'bg-red-100 text-red-600' : 'bg-emerald-100 text-emerald-600' ?>"><?= $v['tipe'] ?></span></td>
                                <td class="p-5 text-right font-mono font-black text-indigo-600"><?= number_format($step1_bobot[$k], 3) ?></td>
                            </tr>
                            <?php endforeach; ?>
                        </tbody>
                    </table>
                </div>
            </div>

            
            <div class="mb-12">
                <h2 class="text-2xl font-black mb-6 text-indigo-900 border-l-8 border-indigo-600 pl-4 uppercase italic">Data Awal Alternatif & Nilai Kriteria</h2>
                <div class="bg-white rounded-3xl shadow-lg border overflow-hidden">
                    <div class="overflow-x-auto overflow-y-auto max-h-[500px] custom-scroll">
                        <table class="w-full min-w-[800px] text-sm text-left">
                            <thead class="bg-slate-900 text-white">
                                <tr>
                                    <th class="p-5 w-1/4">Alternatif</th>
                                    <?php foreach($kriteria_terdeteksi as $k => $v): ?>
                                        <th class="p-5 text-center bg-slate-800">
                                            <?= $k ?><br>
                                            <span class="text-[10px] text-slate-400 font-normal capitalize">(<?= $v['tipe'] ?>)</span>
                                        </th>
                                    <?php endforeach; ?>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100">
                                <?php foreach($data_alternatif as $alt): ?>
                                <tr>
                                    <td class="p-5 font-bold text-slate-700 bg-slate-50/50"><?= $alt['nama'] ?></td>
                                    <?php foreach($kriteria_terdeteksi as $k => $v): ?>
                                        <td class="p-5 text-center font-mono font-bold text-slate-600">
                                            <?= isset($alt[$k]) ? $alt[$k] : '-' ?>
                                        </td>
                                    <?php endforeach; ?>
                                </tr>
                                <?php endforeach; ?>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>




            <div class="mb-12">
                <h2 class="text-2xl font-black mb-6 text-indigo-900 border-l-8 border-indigo-600 pl-4 uppercase italic">Langkah 2: Menghitung Vektor $S_i$</h2>
                <div class="max-h-[500px] overflow-y-auto pr-2 custom-scroll space-y-4">
                    <?php foreach($step2_vektor_s as $s): ?>
                    <div class="bg-white p-6 rounded-3xl shadow-sm border border-gray-200">
                        <div class="mb-2 flex items-center gap-3">
                            <span class="bg-slate-800 text-white text-[10px] font-black px-2 py-1 rounded uppercase"><?= $s['kode'] ?></span>
                            <span class="font-black text-slate-600 uppercase text-xs tracking-tight"><?= $s['nama'] ?></span>
                        </div>
                        <div class="math text-sm bg-slate-50 p-4 rounded-xl border leading-relaxed text-slate-500 italic">
                            <span class="font-black text-indigo-600 not-italic"><?= $s['kode'] ?></span> = <?= $s['rumus'] ?>
                        </div>
                        <div class="mt-2 text-right">
                            <span class="text-lg font-black text-slate-900 font-mono"><?= $s['kode'] ?> = <?= number_format($s['hasil'], 3) ?></span>
                        </div>
                    </div>
                    <?php endforeach; ?>
                </div>
            </div>

            <div class="mb-12">
                <h2 class="text-2xl font-black mb-6 text-indigo-900 border-l-8 border-indigo-600 pl-4 uppercase italic">Total Vektor $S$ ($\sum S_i$)</h2>
                <div class="bg-white p-8 rounded-3xl shadow-md border-2 border-indigo-600">
                    <p class="text-xs font-black text-indigo-400 uppercase mb-4 tracking-widest">Rumus Penjumlahan Seluruh S</p>
                    <div class="math text-lg bg-slate-50 p-6 rounded-2xl border text-slate-700 leading-relaxed custom-scroll overflow-x-auto whitespace-nowrap">
                        <span class="font-black text-indigo-600">Total S</span> = 
                        <?php 
                            // Tampilkan deretan hasil Si
                            echo implode(" + ", $s_values_list); 
                        ?>
                    </div>
                    <div class="mt-4 text-right">
                        <span class="text-3xl font-black text-indigo-950 font-mono italic">Total S = <?= number_format($total_s_all, 3) ?></span>
                    </div>
                </div>
            </div>

            <div class="mb-12">
                <h2 class="text-2xl font-black mb-6 text-indigo-900 border-l-8 border-indigo-600 pl-4 uppercase italic">Langkah 3: Vektor $V_i$ & Ranking</h2>
                
                <div class="grid grid-cols-1 lg:grid-cols-3 gap-8">
                    <div class="lg:col-span-1">
                        <div class="bg-indigo-600 p-4 text-white rounded-t-2xl font-black text-center text-xs tracking-widest uppercase italic">Penjabaran $V_i = S_i / \sum S$</div>
                        <div class="bg-white p-6 rounded-b-2xl shadow border space-y-4 overflow-y-auto max-h-[400px] custom-scroll">
                            <?php foreach($step3_vektor_v as $v): ?>
                            <div class="border-b pb-3 last:border-0">
                                <p class="text-[10px] font-black text-slate-400 mb-1 uppercase"><?= $v['nama'] ?></p>
                                <div class="math text-xs font-bold">
                                    <?= strtolower($v['kode']) ?> = <?= number_format($v['si_val'], 3) ?> / <span class="text-indigo-400"><?= number_format($v['total_s'], 3) ?></span>
                                    <div class="text-indigo-600 mt-1"><?= strtolower($v['kode']) ?> = <?= number_format($v['hasil'], 3) ?></div>
                                </div>
                            </div>
                            <?php endforeach; ?>
                        </div>
                    </div>

                    <div class="lg:col-span-2 bg-white rounded-3xl shadow-2xl border border-gray-200 overflow-hidden flex flex-col">
                        <div class="bg-slate-900 text-white p-4 font-black text-center text-xs tracking-widest uppercase sticky top-0 italic">Tabel Peringkat Akhir</div>
                        <div class="overflow-y-auto max-h-[400px] custom-scroll">
                            <table class="w-full text-left">
                                <thead class="bg-slate-50 text-slate-400 text-[10px] sticky top-0 uppercase border-b">
                                    <tr>
                                        <th class="p-4 text-center w-20">RANK</th>
                                        <th class="p-4">ALTERNATIF</th>
                                        <th class="p-4 text-right">SKOR ($V_i$)</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y">
                                    <?php $rank=1; foreach($step3_vektor_v as $v): ?>
                                    <tr class="<?= $rank == 1 ? 'bg-yellow-50/50 border-l-4 border-yellow-400' : 'hover:bg-slate-50' ?>">
                                        <td class="p-4 text-center">
                                            <?php if($rank == 1): ?>
                                                <span class="bg-yellow-400 text-yellow-900 text-[10px] px-2 py-1 rounded font-black shadow-sm tracking-tighter">JUARA 1</span>
                                            <?php else: ?>
                                                <span class="text-slate-300 font-black text-xs">#<?= $rank ?></span>
                                            <?php endif; ?>
                                        </td>
                                        <td class="p-4 font-black text-slate-700 text-sm italic"><?= $v['nama'] ?></td>
                                        <td class="p-4 text-right font-mono font-black text-indigo-700 text-sm"><?= number_format($v['hasil'], 3) ?></td>
                                    </tr>
                                    <?php $rank++; endforeach; ?>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <?php $win = $step3_vektor_v[0]; ?>
            <section class="bg-slate-900 text-white p-10 rounded-[3rem] shadow-2xl relative overflow-hidden mb-10 border-b-8 border-indigo-600">
                <div class="relative z-10 text-center">
                    <h2 class="text-indigo-400 font-black tracking-widest text-[10px] mb-4 uppercase italic">Rekomendasi Keputusan Akhir</h2>
                    <p class="text-3xl leading-snug font-black italic">
                         Sistem merekomendasikan <span class="text-yellow-400 underline decoration-indigo-500 underline-offset-8"><?= $win['nama'] ?></span> sebagai pilihan terbaik dengan perolehan skor preferensi <?= number_format($win['hasil'], 3) ?>.
                    </p>
                </div>
                <div class="absolute -right-20 -bottom-20 w-64 h-64 bg-indigo-600 rounded-full opacity-10 blur-3xl"></div>
            </section>

        <?php endif; ?>
    </main>

    <footer class="py-12 text-center text-slate-400 text-[10px] font-black uppercase tracking-[0.3em]">
        Weighted Product Dashboard &bull; Precision Calculation &bull; Native PHP
    </footer>

</body>
</html>