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
      clone.querySelector('.date').textContent = row["DATE"];
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
