document.addEventListener('DOMContentLoaded', function () {
    const form = document.querySelector('.form-komponen');
    initializeCamera(form);

    document.querySelector('#addKomponenBtn').addEventListener('click', () => {
        const komponen = document.querySelector('[name="id_komponen[]"]').value;
        const komponenText = document.querySelector('[name="id_komponen[]"] option:checked').text;
        const temuan = document.querySelector('[name="id_temuan[]"]').value;
        const temuanText = document.querySelector('[name="id_temuan[]"] option:checked').text;
        const solusi = document.querySelector('[name="id_solusi[]"]').value;
        const solusiText = document.querySelector('[name="id_solusi[]"] option:checked').text;
        const prioritas = document.querySelector('[name="prioritas[]"]').value;
        const prioritasText = document.querySelector('[name="prioritas[]"] option:checked').text;
        const keterangan = document.querySelector('[name="keterangan[]"]').value;
        const fotoInput = document.querySelector('[name="foto_bukti[]"]');
        const fotoFile = fotoInput.files[0];
    
        // Validasi input
        if (!komponen || !temuan || !solusi || !prioritas || !fotoFile) {
            alert('Harap lengkapi semua field yang bertanda * sebelum menambahkan.');
            return;
        }
    
        const reader = new FileReader();
        reader.onload = function (e) {
            const tableBody = document.querySelector('#komponenTable tbody');
            const newRow = document.createElement('tr');
            newRow.innerHTML = `
                <td>${komponenText}<input type="hidden" name="id_komponen[]" value="${komponen}"></td>
                <td>${temuanText}<input type="hidden" name="id_temuan[]" value="${temuan}"></td>
                <td>${solusiText}<input type="hidden" name="id_solusi[]" value="${solusi}"></td>
                <td>${prioritasText}<input type="hidden" name="prioritas[]" value="${prioritas}"></td>
                <td>${keterangan}<input type="hidden" name="keterangan[]" value="${keterangan}"></td>
                <td><img src="${e.target.result}" class="img-thumbnail" style="max-width: 100px;"><input type="hidden" name="foto_bukti_base64[]" value="${e.target.result}"></td>
                <td><button type="button" class="btn btn-danger btn-delete-row">Hapus</button></td>
            `;
            tableBody.appendChild(newRow);
    
            // Event untuk menghapus baris
            newRow.querySelector('.btn-delete-row').addEventListener('click', () => {
                newRow.remove();
            });
    
            // Reset form setelah menambahkan ke tabel
            form.reset();
        };
        reader.readAsDataURL(fotoFile);
    });
});
