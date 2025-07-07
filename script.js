document.getElementById('searchForm').addEventListener('submit', function(event) {
  event.preventDefault();

  const enroll = document.getElementById('enroll').value.trim();
  const grno = document.getElementById('grno').value.trim();

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
  const subject = data[0][1];  
  const sem = data[1][1];      
  const className = data[2][1]; 

  for (let i = 4; i < data.length; i++) { 
    const row = data[i];
    if (row[0] == enroll && row[1] == grno) {
      return {
        name: row[2],
        marks: row[3],
        subject: subject,
        sem: sem,
        class: className
      };
    }
  }
  return null;
}

function displayResult(record) {
  const resultDiv = document.getElementById('result');
  resultDiv.innerHTML = '';

  if (record) {
    resultDiv.innerHTML = `<h4>Result Found!</h4><br>
                           <h5>Name: ${record.name}</h5>
                           <h6>Sem: ${record.sem} | Class: ${record.class}</h6>
                           <h5>Subject: ${record.subject}</h5>
                           <h5>Marks: ${record.marks}</h5>`;
  } else {
    resultDiv.innerHTML = `<p class="text-danger">No matching record found!!</p>`;
  }
}
