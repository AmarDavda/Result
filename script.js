document.getElementById('searchForm').addEventListener('submit', function(event) {
  event.preventDefault();

  const enroll = document.getElementById('enroll').value.trim();
  const grno = document.getElementById('grno').value.trim();

  // Assume we have the Excel file uploaded or read
  const fileInput = 'data.xlsx';

  fetchExcelData(fileInput, enroll, grno);
});

function fetchExcelData(filePath, enroll, grno) {
  const reader = new XMLHttpRequest();
  reader.open("GET", filePath, true);
  reader.responseType = "arraybuffer";

  reader.onload = function(e) {
    const data = new Uint8Array(reader.response);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const matchedRecord = searchForRecord(jsonData, enroll, grno);

    displayResult(matchedRecord);
  };

  reader.send();
}

function searchForRecord(data, enroll, grno) {
  for (let i = 1; i < data.length; i++) { //skips header line
    const row = data[i];
    if (row[0] == enroll && row[1] == grno) {
      return { name: row[2], marks: row[3] };
    }
  }
  return null;
}

function displayResult(record) {
  const resultDiv = document.getElementById('result');
  resultDiv.innerHTML = '';

  if (record) {
    resultDiv.innerHTML = `<h5>Result Found</h5><br>
                           <h6>Name: ${record.name}</h6>
                           <h6>Marks: ${record.marks}</h6>`;
  } else {
    resultDiv.innerHTML = `<p class="text-danger">No matching record found!!</p>`;
  }
}
