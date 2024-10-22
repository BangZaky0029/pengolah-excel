document.getElementById('processButton').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Silakan upload file Excel terlebih dahulu!');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});

        // Ambil data dari sheet pertama
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Konversi ke JSON
        const json = XLSX.utils.sheet_to_json(worksheet);

        // Ambil kolom yang diinginkan
        const selectedColumns = json.map(row => ({
            "No. Pesanan": row["No. Pesanan"],
            "Status Pesanan": row["Status Pesanan"],
            "Pesanan Harus Dikirimkan Sebelum (Menghindari keterlambatan)": row["Pesanan Harus Dikirimkan Sebelum (Menghindari keterlambatan)"],
            "Waktu Pesanan Dibuat": row["Waktu Pesanan Dibuat"],
            "Waktu Pembayaran Dilakukan": row["Waktu Pembayaran Dilakukan"],
            "SKU Induk": row["SKU Induk"],
            "Nomor Referensi SKU": row["Nomor Referensi SKU"],
            "Nama Variasi": row["Nama Variasi"],
            "Catatan dari Pembeli": row["Catatan dari Pembeli"]
        }));

        // Buat file baru
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(selectedColumns);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Hasil');

        // Buat download link
        const downloadLink = document.getElementById('downloadLink');
        downloadLink.href = URL.createObjectURL(new Blob([XLSX.write(newWorkbook, {bookType:'xlsx', type: 'array'})]));
        downloadLink.download = 'hasil_olah.xlsx';
        downloadLink.innerText = 'Download Hasil';
        document.getElementById('result').classList.remove('hidden');
    };

    reader.readAsArrayBuffer(file);
});
