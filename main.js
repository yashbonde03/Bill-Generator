function excelDateToJSDate(input) {
  // Case 1: If it's a number (Excel serial)
  if (typeof input === 'number') {
    const utc_days = Math.floor(input - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);

    const fractional_day = input % 1;
    const totalSeconds = Math.round(fractional_day * 86400);
    const hours = String(Math.floor(totalSeconds / 3600)).padStart(2, '0');
    const minutes = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0');

    const dd = String(date_info.getDate()).padStart(2, '0');
    const mm = String(date_info.getMonth() + 1).padStart(2, '0');
    const yyyy = date_info.getFullYear();

    return `${dd}-${mm}-${yyyy} ${hours}:${minutes}`;
  }

  // Case 2: If it's a string like "29/01/25 16:45"
  if (typeof input === 'string') {
    const [datePart, timePart] = input.split(' ');
    const [day, month, year] = datePart.split('/');

    const fullYear = Number(year) < 100 ? `20${year}` : year;
    return `${day.padStart(2, '0')}-${month.padStart(2, '0')}-${fullYear} ${timePart}`;
  }

  return 'Invalid Date';
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
      clone.querySelector('.site').textContent = row["SITE"];
      clone.querySelector('.address').textContent = row["ADDRESS"];
      clone.querySelector('.tel').textContent = row["TEL NO"];
    //   clone.querySelector('.date').textContent = row["DATE"];
    clone.querySelector('.date').textContent = excelDateToJSDate(row["DATE"]);

//     const serial = row["DATE"];
// clone.querySelector('.date').textContent = excelDateToJSDate(serial);
      clone.querySelector('.vehicle').textContent = row["VEHICLE"];
      clone.querySelector('.bsn').textContent = row["BSN"];
      clone.querySelector('.hose').textContent = row["HOSE ID"];
      clone.querySelector('.density').textContent = row["DENSITY"];
      clone.querySelector('.rate').textContent = row["RATE"];
      clone.querySelector('.volume').textContent = row["VOLUME"];
      clone.querySelector('.amount').textContent = row["AMOUNT"];
      container.appendChild(clone);
    });
  };
  reader.readAsBinaryString(file);
});

document.getElementById('savePdfBtn').addEventListener('click', async () => {
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ unit: 'pt', format: 'a4' });

  const bills = document.querySelectorAll('.bill');
  if (bills.length === 0) return alert('No bills to save.');

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  for (let i = 0; i < bills.length; i++) {
    const bill = bills[i];
    bill.scrollIntoView();

    try {
      const canvas = await html2canvas(bill, {
        scale: 2,
        useCORS: true,
      });

      const imgData = canvas.toDataURL('image/png');
      const imgProps = pdf.getImageProperties(imgData);
      const imgWidth = pageWidth * 0.9;
      const imgHeight = (imgProps.height * imgWidth) / imgProps.width;
      const x = (pageWidth - imgWidth) / 2;
      const y = (pageHeight - imgHeight) / 2;

      if (i > 0) pdf.addPage();
      pdf.addImage(imgData, 'PNG', x, y, imgWidth, imgHeight);
    } catch (err) {
      console.error('PDF generation error:', err);
      alert('PDF generation failed. Check console for details.');
      return;
    }
  }

  pdf.save('IndianOil_Bills.pdf');
});
