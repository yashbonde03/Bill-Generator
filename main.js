document.getElementById('excelInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const workbook = XLSX.read(evt.target.result, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet);

        const container = document.getElementById('billsContainer');
        container.innerHTML = '';

        const template = document.getElementById('billTemplate');

        data.forEach((row, index) => {
            const clone = template.content.cloneNode(true);

            clone.querySelector('.site').textContent = row["SITE"] || '';
            clone.querySelector('.address').textContent = row["ADDRESS"] || '';
            clone.querySelector('.tel').textContent = row["TEL NO"] || '';
            
            const dateElem = clone.querySelector('.date');
            const dateInput = clone.querySelector('.date-input');
            if (!isNaN(row["DATE"])) {
                const excelEpoch = new Date((row["DATE"] - 25569) * 86400 * 1000);
                const formattedDate = excelEpoch.toLocaleString();
                dateElem.textContent = formattedDate;
                dateInput.value = formattedDate;
            } else {
                dateElem.textContent = row["DATE"] || '';
                dateInput.value = row["DATE"] || '';
            }

            dateInput.addEventListener('input', () => {
                dateElem.textContent = dateInput.value;
            });

            clone.querySelector('.vehicle').textContent = row["VEHICLE"] || 'NOT ENTERED';
            clone.querySelector('.bsn').textContent = row["BSN"] || '';
            clone.querySelector('.hose').textContent = row["HOSE ID"] || '';
            clone.querySelector('.density').textContent = row["DENSITY"] || '';
            clone.querySelector('.rate').textContent = row["RATE"] || '';
            clone.querySelector('.volume').textContent = row["VOLUME"] || '';
            clone.querySelector('.amount').textContent = row["AMOUNT"] || '';

            container.appendChild(clone);
        });
    };
    reader.readAsBinaryString(file);
});
