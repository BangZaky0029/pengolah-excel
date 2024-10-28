// Proses File Shopee
document.getElementById('processButton').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Silakan upload file Shopee terlebih dahulu!');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);

        const selectedColumns = json.map(row => ({
            "Waktu Pembayaran Dilakukan": row["Waktu Pembayaran Dilakukan"],
            "SKU Induk": row["SKU Induk"],
            "Nomor Referensi SKU": row["Nomor Referensi SKU"],
            "Catatan dari Pembeli": row["Catatan dari Pembeli"]
        }));

        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(selectedColumns);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Hasil Shopee');

        const downloadLink = document.getElementById('downloadLink');
        downloadLink.href = URL.createObjectURL(new Blob([XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' })]));
        downloadLink.download = 'hasil_shopee.xlsx';
        downloadLink.innerText = 'Download Hasil Shopee';
        document.getElementById('result').classList.remove('hidden');
    };

    reader.readAsArrayBuffer(file);
});

// Proses File TikTok
document.getElementById('tiktokProcessButton').addEventListener('click', function() {
    const tiktokFileInput = document.getElementById('tiktokFileInput');
    const file = tiktokFileInput.files[0];

    if (!file) {
        alert('Silakan upload file TikTok terlebih dahulu!');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        let json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Hapus baris kedua
        if (json.length > 1) {
            json.splice(1, 1);
        }

        const processedData = json.slice(1).map(row => ({
            "SKU ID": row[5],
            "Created Time": row[25],
            "Buyer Message": row[38]
        }));

        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(processedData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'TikTok Hasil');

        const tiktokDownloadLink = document.getElementById('tiktokDownloadLink');
        tiktokDownloadLink.href = URL.createObjectURL(new Blob([XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' })]));
        tiktokDownloadLink.download = 'hasil_tiktok.xlsx';
        tiktokDownloadLink.innerText = 'Download Hasil TikTok';
        document.getElementById('tiktokResult').classList.remove('hidden');
    };

    reader.readAsArrayBuffer(file);
});
