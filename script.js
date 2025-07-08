let pileData = [];

document.getElementById('excelFile').addEventListener('change', function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    pileData = jsonData;
    populatePileOptions();
  };
  reader.readAsArrayBuffer(file);
});

function populatePileOptions() {
  const select = document.getElementById('pileSelect');
  select.innerHTML = '<option value="">-- เลือก --</option>';
  pileData.forEach(row => {
    if (row['Pile No']) {
      const option = document.createElement('option');
      option.value = row['Pile No'];
      option.textContent = row['Pile No'];
      select.appendChild(option);
    }
  });
}

document.getElementById('pileSelect').addEventListener('change', function() {
  const selected = this.value;
  const pile = pileData.find(p => p['Pile No'] == selected);
  if (pile) {
    document.getElementById('nDesign').value = pile['N'] || '';
    document.getElementById('eDesign').value = pile['E'] || '';
  } else {
    document.getElementById('nDesign').value = '';
    document.getElementById('eDesign').value = '';
  }
});

function calculate() {
  const nDesign = parseFloat(document.getElementById('nDesign').value);
  const eDesign = parseFloat(document.getElementById('eDesign').value);
  const nActual = parseFloat(document.getElementById('nActual').value);
  const eActual = parseFloat(document.getElementById('eActual').value);

  if (isNaN(nDesign) || isNaN(eDesign) || isNaN(nActual) || isNaN(eActual)) {
    document.getElementById('result').innerText = 'กรุณากรอกค่าทั้งหมด';
    return;
  }

  const dN = nActual - nDesign;
  const dE = eActual - eDesign;
  const deviation = Math.sqrt(dN * dN + dE * dE).toFixed(3);

  document.getElementById('result').innerHTML = `
     ค่าคลาดเคลื่อน (Deviation): <b>${deviation}</b> เมตร<br>
     ค่าต่าง N: ${dN.toFixed(3)}, E: ${dE.toFixed(3)}
  `;
}

function clearForm() {
  document.getElementById('pileSelect').value = '';
  document.getElementById('nDesign').value = '';
  document.getElementById('eDesign').value = '';
  document.getElementById('nActual').value = '';
  document.getElementById('eActual').value = '';
  document.getElementById('result').innerText = '-';
}
