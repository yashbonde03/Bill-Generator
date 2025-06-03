<script>
    // Convert Excel serial date to "YYYY-MM-DD HH:MM"
    function excelDateToJSDate(serial) {
        const utc_days = Math.floor(serial - 25569);
        const utc_value = utc_days * 86400;
        const date_info = new Date(utc_value * 1000);

        const fractional_day = serial - Math.floor(serial) + 0.0000001;
        let totalSeconds = Math.floor(86400 * fractional_day);

        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);

        const dateStr = `${date_info.getFullYear()}-${(date_info.getMonth() + 1).toString().padStart(2, '0')}-${date_info.getDate().toString().padStart(2, '0')}`;
        const timeStr = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        return `${dateStr} ${timeStr}`;
    }

    document.getElementById('excelInput').addEventListener('change', function (e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function (evt) {
            const workbook = XLSX.read(evt.target.result, { type: 'binary' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet);

            const container = document.getElementById('billsContainer');
            container.innerHTML = '';

            const template = document.getElementById('billTemplate');

            data.forEach(row => {
                const clone = template.content.cloneNode(true);

                clone.querySelector('.site').textContent = row["SITE"] || '';
                clone.querySelector('.address').textContent = row["ADDRESS"] || '';
                clone.querySelector('.tel').textContent = row["TEL NO"] || '';

                // ✅ Handle both manual date string and serial number
                const rawDate = row["DATE"];
                let formattedDate = '';
                if (typeof rawDate === 'number') {
                    formattedDate = excelDateToJSDate(rawDate);
                } else {
                    formattedDate = rawDate || '';
                }
                clone.querySelector('.date').textContent = formattedDate;

                clone.querySelector('.vehicle').textContent = row["VEHICLE"] || '';
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
</script>
