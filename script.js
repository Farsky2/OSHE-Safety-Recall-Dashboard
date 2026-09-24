// GLOBAL STATE VARIABLES
let rawDataset = [];
let filteredDataset = [];
let chartInstanceSite = null;
let chartInstanceRemidi = null;

// PAGINATION STATE
let currentPage = 1;
const pageSize = 10;

// ON PAGE LOAD INITIALIZATION
window.onload = function() {
    // Bind listeners for Step 1 Configuration Inputs
    const inputOSHE = document.getElementById('inputPenandatanganOSHE');
    const inputGM = document.getElementById('inputPenandatanganGM');
    
    if (inputOSHE) inputOSHE.addEventListener('input', updateSignatures);
    if (inputGM) inputGM.addEventListener('input', updateSignatures);
    
    // Bind Drag & Drop Events
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    if (dropZone && fileInput) {
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('bg-amber-100', 'border-amber-500');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('bg-amber-100', 'border-amber-500');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('bg-amber-100', 'border-amber-500');
            if (e.dataTransfer.files.length > 0) {
                fileInput.files = e.dataTransfer.files;
                handleFileUpload(e.dataTransfer.files[0]);
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFileUpload(e.target.files[0]);
            }
        });
    }

    updateSignatures();
};

function updateSignatures() {
    const osheElem = document.getElementById('signNameOSHE');
    const gmElem = document.getElementById('signNameGM');
    const inputOSHE = document.getElementById('inputPenandatanganOSHE');
    const inputGM = document.getElementById('inputPenandatanganGM');

    if (osheElem && inputOSHE) osheElem.innerText = inputOSHE.value || 'M. Harris Domili';
    if (gmElem && inputGM) gmElem.innerText = inputGM.value || 'Nama General Manager';
}

function prosesSimulasiDemo() {
    const demoSites = ['Site BAS-A (Batu Kajang)', 'Site BAS-B (Tabang)', 'Site BAS-C (IKN)', 'Head Office Balikpapan', 'Site BAS-D (Lahat)'];
    const demoDepts = ['Operation', 'Engineering', 'OSHE', 'Plant / Maintenance', 'HRGA & Safety'];
    const demoTopics = ['Fatigue Management (BAS.OSHE.PR.012)', 'Prosedur Dumping (BAS.OPR.PR.003)', 'Standar Jalan Tambang (BAS.OPR.PR.005)'];
    const demoNames = [
        'Ahmad Fauzi', 'Budi Santoso', 'Chandra Wijaya', 'Dedi Kurniawan', 'Eko Prasetyo',
        'Fajar Nugraha', 'Gita Gutawa', 'Hendra Setiawan', 'Indra Lesmana', 'Joko Widodo',
        'Kiki Fatmala', 'Lukman Hakim', 'M. Harris Domili', 'Nurdin Abdullah', 'Oki Setiana',
        'Putri Indonesia', 'Qori Sandioriva', 'Rizky Febian', 'Siti Badriah', 'Taufik Hidayat',
        'Umar Amir', 'Vina Panduwinata', 'Wahyu Hidayat', 'Xaverius Roy', 'Yusuf Mansur',
        'Zulkifli Hasan', 'Andi Pratama', 'Bagas Kaffa', 'Cinta Laura', 'Doni Monardo'
    ];

    const generated = [];

    demoNames.forEach((nama, idx) => {
        const nik = 'NIK-' + (1000 + idx);
        const site = demoSites[idx % demoSites.length];
        const dept = demoDepts[idx % demoDepts.length];
        const topik = demoTopics[idx % demoTopics.length];
        
        const attempts = Math.floor(Math.random() * 3) + 1; 
        let score = 0;

        if (attempts === 1) {
            score = Math.floor(Math.random() * 20) + 80; 
        } else if (attempts === 2) {
            score = Math.random() > 0.2 ? Math.floor(Math.random() * 15) + 80 : Math.floor(Math.random() * 20) + 55;
        } else {
            score = Math.random() > 0.4 ? Math.floor(Math.random() * 15) + 80 : Math.floor(Math.random() * 25) + 50;
        }

        const status = score >= 80 ? 'LULUS' : 'REMIDI';

        generated.push({
            nik: nik,
            nama: nama,
            site: site,
            dept: dept,
            topik: topik,
            skor: score,
            attempts: attempts,
            status: status
        });
    });

    rawDataset = generated;
    showAlert('Data simulasi berhasil dimuat! ' + generated.length + ' data peserta disiapkan.', 'success');
    processDataset();
}

function handleFileUpload(file) {
    showAlert('Membaca dan memproses file Excel: ' + file.name + '...', 'info');

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Otomatis pilih sheet 'Log Peserta' jika ada, jika tidak ambil sheet pertama
            let targetSheetName = workbook.SheetNames.find(name => name.toLowerCase().includes('log peserta')) || workbook.SheetNames[0];
            const worksheet = workbook.Sheets[targetSheetName];
            const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            if (!jsonRows || jsonRows.length === 0) {
                showAlert('File Excel kosong atau format tidak sesuai.', 'error');
                return;
            }

            const mappedData = jsonRows.map((row, index) => {
                const keys = Object.keys(row);

                const findVal = (possibleKeywords, defaultVal) => {
                    for (let key of keys) {
                        const cleanKey = key.toLowerCase().trim();
                        for (let kw of possibleKeywords) {
                            if (cleanKey.includes(kw)) {
                                return row[key];
                            }
                        }
                    }
                    return defaultVal;
                };

                const nama = findVal(['nama', 'employee', 'karyawan', 'name'], 'Peserta ' + (index + 1));
                const nik = findVal(['nik', 'id', 'no', 'nrp'], 'NIK-' + (2000 + index));
                const site = findVal(['site', 'lokasi', 'project', 'cabang'], 'Site Utama');
                const dept = findVal(['dept', 'departemen', 'divisi', 'department'], 'OSHE');
                const topik = findVal(['topik', 'materi', 'evaluasi', 'modul', 'quiz'], 'Safety Recall');
                
                let skor = parseFloat(findVal(['skor', 'nilai', 'score', 'hasil', 'percentage'], 0));
                if (isNaN(skor)) skor = 0;

                let attempts = parseInt(findVal(['percobaan', 'remidi', 'ulang', 'attempt', 'kali'], 1));
                if (isNaN(attempts) || attempts < 1) attempts = 1;

                let statusStr = findVal(['status', 'kelulusan'], '');
                let status = 'REMIDI';

                if (statusStr.toString().toUpperCase().includes('LULUS') || skor >= 80) {
                    status = 'LULUS';
                }

                return {
                    nik: String(nik),
                    nama: String(nama),
                    site: String(site),
                    dept: String(dept),
                    topik: String(topik),
                    skor: skor,
                    attempts: attempts,
                    status: status
                };
            });

            // Filter baris kosong
            const filteredMappedData = mappedData.filter(d => d.nama && d.nama.trim() !== '' && !d.nama.startsWith('Peserta '));

            rawDataset = filteredMappedData;
            showAlert('File Excel ' + file.name + ' berhasil di-parsing (' + filteredMappedData.length + ' baris data dari sheet "' + targetSheetName + '")!', 'success');
            processDataset();

        } catch (err) {
            console.error(err);
            showAlert('Gagal membaca file Excel. Pastikan format file valid.', 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function showAlert(msg, type) {
    const alertBox = document.getElementById('statusAlert');
    if (!alertBox) return;

    alertBox.classList.remove('hidden', 'bg-emerald-100', 'text-emerald-800', 'bg-rose-100', 'text-rose-800', 'bg-amber-100', 'text-amber-800');
    
    if (type === 'success') {
        alertBox.classList.add('bg-emerald-100', 'text-emerald-800');
        alertBox.innerHTML = '<i class="fa-solid fa-circle-check text-base"></i><span>' + msg + '</span>';
    } else if (type === 'error') {
        alertBox.classList.add('bg-rose-100', 'text-rose-800');
        alertBox.innerHTML = '<i class="fa-solid fa-circle-xmark text-base"></i><span>' + msg + '</span>';
    } else {
        alertBox.classList.add('bg-amber-100', 'text-amber-800');
        alertBox.innerHTML = '<i class="fa-solid fa-spinner fa-spin text-base"></i><span>' + msg + '</span>';
    }
}

function processDataset() {
    populateDropdownFilters();
    applyFilters();
}

function populateDropdownFilters() {
    const siteSelect = document.getElementById('filterSite');
    const deptSelect = document.getElementById('filterDept');
    if (!siteSelect || !deptSelect) return;

    const uniqueSites = [...new Set(rawDataset.map(d => d.site))].filter(Boolean).sort();
    const uniqueDepts = [...new Set(rawDataset.map(d => d.dept))].filter(Boolean).sort();

    siteSelect.innerHTML = '<option value="ALL">Semua Site (' + uniqueSites.length + ')</option>';
    uniqueSites.forEach(s => {
        siteSelect.innerHTML += `<option value="${s}">${s}</option>`;
    });

    deptSelect.innerHTML = '<option value="ALL">Semua Departemen (' + uniqueDepts.length + ')</option>';
    uniqueDepts.forEach(d => {
        deptSelect.innerHTML += `<option value="${d}">${d}</option>`;
    });
}

function applyFilters() {
    const searchInput = document.getElementById('filterSearch');
    const siteInput = document.getElementById('filterSite');
    const deptInput = document.getElementById('filterDept');
    const statusInput = document.getElementById('filterStatus');

    const searchText = searchInput ? searchInput.value.toLowerCase() : '';
    const selectedSite = siteInput ? siteInput.value : 'ALL';
    const selectedDept = deptInput ? deptInput.value : 'ALL';
    const selectedStatus = statusInput ? statusInput.value : 'ALL';

    filteredDataset = rawDataset.filter(item => {
        const namaStr = String(item.nama || '').toLowerCase();
        const nikStr = String(item.nik || '').toLowerCase();
        
        const matchSearch = namaStr.includes(searchText) || nikStr.includes(searchText);
        const matchSite = selectedSite === 'ALL' || item.site === selectedSite;
        const matchDept = selectedDept === 'ALL' || item.dept === selectedDept;

        let matchStatus = true;
        if (selectedStatus === 'Lulus Langsung') {
            matchStatus = item.status === 'LULUS' && item.attempts === 1;
        } else if (selectedStatus === 'Lulus Setelah Remidi') {
            matchStatus = item.status === 'LULUS' && item.attempts > 1;
        } else if (selectedStatus === 'Belum Lulus / Remidi') {
            matchStatus = item.status === 'REMIDI';
        }

        return matchSearch && matchSite && matchDept && matchStatus;
    });

    currentPage = 1;
    updateMetricsScorecard();
    renderCharts();
    renderTable();
}

function updateMetricsScorecard() {
    const total = filteredDataset.length;
    const lulusCount = filteredDataset.filter(d => d.status === 'LULUS').length;
    const remidiCount = filteredDataset.filter(d => d.attempts > 1 || d.status === 'REMIDI').length;
    const passRate = total > 0 ? ((lulusCount / total) * 100).toFixed(1) : 0;

    const siteRemidiMap = {};
    filteredDataset.forEach(d => {
        if (d.attempts > 1 || d.status === 'REMIDI') {
            siteRemidiMap[d.site] = (siteRemidiMap[d.site] || 0) + 1;
        }
    });

    let topRemidiSite = '-';
    let maxRemidi = 0;
    for (let s in siteRemidiMap) {
        if (siteRemidiMap[s] > maxRemidi) {
            maxRemidi = siteRemidiMap[s];
            topRemidiSite = s;
        }
    }

    const setVal = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.innerText = val;
    };

    setVal('statTotalPeserta', total);
    setVal('statTotalLulus', lulusCount);
    setVal('statTotalRemidi', remidiCount);
    setVal('statPassRate', passRate + '%');
    setVal('statSiteRemidiTop', topRemidiSite);
    setVal('statSiteRemidiCount', maxRemidi + ' Peserta Mengulang');
    setVal('labelTotalRecords', 'Data Terisi: ' + total + ' Peserta');
}

function renderCharts() {
    const siteMap = {};
    filteredDataset.forEach(d => {
        if (!siteMap[d.site]) {
            siteMap[d.site] = { lulus: 0, remidi: 0 };
        }
        if (d.status === 'LULUS') {
            siteMap[d.site].lulus++;
        } else {
            siteMap[d.site].remidi++;
        }
    });

    const siteLabels = Object.keys(siteMap);
    const lulusData = siteLabels.map(s => siteMap[s].lulus);
    const remidiData = siteLabels.map(s => siteMap[s].remidi);

    const canvasSite = document.getElementById('chartSiteComparison');
    if (canvasSite) {
        const ctx1 = canvasSite.getContext('2d');
        if (chartInstanceSite) chartInstanceSite.destroy();

        chartInstanceSite = new Chart(ctx1, {
            type: 'bar',
            data: {
                labels: siteLabels,
                datasets: [
                    {
                        label: 'Lulus (≥80%)',
                        data: lulusData,
                        backgroundColor: '#10B981'
                    },
                    {
                        label: 'Remidi (<80%)',
                        data: remidiData,
                        backgroundColor: '#EF4444'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'top' }
                },
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 } }
                }
            }
        });
    }

    let attempt1Count = 0;
    let attempt2Count = 0;
    let attempt3PlusCount = 0;

    filteredDataset.forEach(d => {
        if (d.attempts === 1) attempt1Count++;
        else if (d.attempts === 2) attempt2Count++;
        else attempt3PlusCount++;
    });

    const canvasRemidi = document.getElementById('chartRemidiCount');
    if (canvasRemidi) {
        const ctx2 = canvasRemidi.getContext('2d');
        if (chartInstanceRemidi) chartInstanceRemidi.destroy();

        chartInstanceRemidi = new Chart(ctx2, {
            type: 'doughnut',
            data: {
                labels: ['Percobaan 1x (Langsung Lulus)', 'Percobaan 2x (Remidi 1x)', 'Percobaan 3x+ (Remidi ≥2x)'],
                datasets: [{
                    data: [attempt1Count, attempt2Count, attempt3PlusCount],
                    backgroundColor: ['#10B981', '#F59E0B', '#EF4444']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'bottom' }
                }
            }
        });
    }
}

function renderTable() {
    const tableBody = document.getElementById('tableBody');
    if (!tableBody) return;
    tableBody.innerHTML = '';

    if (filteredDataset.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="9" class="py-8 text-center text-slate-400">
                    Tidak ada data yang sesuai dengan filter pencarian.
                </td>
            </tr>`;
        const pagInfo = document.getElementById('paginationInfo');
        const pagCtrl = document.getElementById('paginationControls');
        if (pagInfo) pagInfo.innerText = 'Menampilkan 0 data';
        if (pagCtrl) pagCtrl.innerHTML = '';
        return;
    }

    const totalPages = Math.ceil(filteredDataset.length / pageSize);
    if (currentPage > totalPages) currentPage = totalPages;

    const startIndex = (currentPage - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, filteredDataset.length);
    const pageData = filteredDataset.slice(startIndex, endIndex);

    pageData.forEach((row, idx) => {
        const globalIndex = startIndex + idx + 1;
        
        let statusBadge = '';
        if (row.status === 'LULUS') {
            if (row.attempts === 1) {
                statusBadge = `<span class="bg-emerald-100 text-emerald-800 border border-emerald-300 font-bold px-2.5 py-1 rounded-full text-[10px]">
                    <i class="fa-solid fa-check"></i> Lulus Direct
                </span>`;
            } else {
                statusBadge = `<span class="bg-blue-100 text-blue-800 border border-blue-300 font-bold px-2.5 py-1 rounded-full text-[10px]">
                    <i class="fa-solid fa-check-double"></i> Lulus (Remidi ${row.attempts - 1}x)
                </span>`;
            }
        } else {
            statusBadge = `<span class="bg-rose-100 text-rose-800 border border-rose-300 font-bold px-2.5 py-1 rounded-full text-[10px]">
                <i class="fa-solid fa-triangle-exclamation"></i> Belum Lulus (Remidi)
            </span>`;
        }

        let attemptText = `<span class="font-semibold text-slate-700">${row.attempts}x Ujian</span>`;
        if (row.attempts > 1) {
            attemptText += ` <span class="text-[10px] text-amber-600 font-bold">(${row.attempts - 1}x Remidi)</span>`;
        }

        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition border-b border-slate-200';
        tr.innerHTML = `
            <td class="py-3 px-4 text-slate-500">${globalIndex}</td>
            <td class="py-3 px-4 font-mono text-slate-600">${row.nik}</td>
            <td class="py-3 px-4 font-bold text-slate-900">${row.nama}</td>
            <td class="py-3 px-4 text-slate-700">${row.site}</td>
            <td class="py-3 px-4 text-slate-600">${row.dept}</td>
            <td class="py-3 px-4 text-slate-600 max-w-xs truncate" title="${row.topik}">${row.topik}</td>
            <td class="py-3 px-4 text-center font-black ${row.skor >= 80 ? 'text-emerald-600' : 'text-rose-600'}">${row.skor}</td>
            <td class="py-3 px-4 text-center">${attemptText}</td>
            <td class="py-3 px-4 text-center">${statusBadge}</td>
        `;
        tableBody.appendChild(tr);
    });

    const pagInfo = document.getElementById('paginationInfo');
    if (pagInfo) pagInfo.innerText = `Menampilkan ${startIndex + 1} - ${endIndex} dari ${filteredDataset.length} data`;
    
    let navHtml = '';
    navHtml += `<button onclick="changePage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''} class="px-2.5 py-1 bg-white border border-slate-300 rounded text-xs disabled:opacity-50">Prev</button>`;
    
    for (let p = 1; p <= totalPages; p++) {
        if (p === currentPage) {
            navHtml += `<button class="px-2.5 py-1 bg-amber-500 text-slate-950 font-bold rounded text-xs">${p}</button>`;
        } else if (p === 1 || p === totalPages || Math.abs(p - currentPage) <= 1) {
            navHtml += `<button onclick="changePage(${p})" class="px-2.5 py-1 bg-white border border-slate-300 rounded text-xs">${p}</button>`;
        }
    }
    
    navHtml += `<button onclick="changePage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''} class="px-2.5 py-1 bg-white border border-slate-300 rounded text-xs disabled:opacity-50">Next</button>`;
    
    const pagCtrl = document.getElementById('paginationControls');
    if (pagCtrl) pagCtrl.innerHTML = navHtml;
}

function changePage(page) {
    const totalPages = Math.ceil(filteredDataset.length / pageSize);
    if (page >= 1 && page <= totalPages) {
        currentPage = page;
        renderTable();
    }
}

function downloadRekapCSV() {
    if (filteredDataset.length === 0) {
        showAlert('Tidak ada data untuk di-export.', 'error');
        return;
    }

    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += "NIK,Nama Karyawan,Site,Departemen,Topik Evaluasi,Skor Akhir,Jumlah Percobaan,Jumlah Remidi,Status Kelulusan\n";

    filteredDataset.forEach(d => {
        const remidiCount = d.attempts > 1 ? d.attempts - 1 : 0;
        const rowStr = `"${d.nik}","${d.nama}","${d.site}","${d.dept}","${d.topik}",${d.skor},${d.attempts},${remidiCount},"${d.status}"`;
        csvContent += rowStr + "\n";
    });

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Rekap_OSHE_Safety_Recall_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function setActiveNav(element) {
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('bg-amber-500', 'text-slate-950', 'font-semibold', 'shadow-sm');
        item.classList.add('hover:bg-slate-800', 'hover:text-white');
    });
    element.classList.add('bg-amber-500', 'text-slate-950', 'font-semibold', 'shadow-sm');
    element.classList.remove('hover:bg-slate-800', 'hover:text-white');
}
