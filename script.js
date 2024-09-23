document.getElementById('searchForm').addEventListener('submit', function(event) {
  event.preventDefault();

  // Get user input
  const enroll = document.getElementById('enroll').value.trim();
  const grno = document.getElementById('grno').value.trim();

  // Assume we have the Excel file uploaded or read
  // For simplicity, this example uses a pre-loaded file path
  const fileInput = 'data.xlsx'; // You can change this to actual file input logic if required

  fetchExcelData(fileInput, enroll, grno);
});

// Function to fetch data from the Excel sheet and match
function fetchExcelData(filePath, enroll, grno) {
  const reader = new XMLHttpRequest();
  reader.open("GET", filePath, true);
  reader.responseType = "arraybuffer";

  reader.onload = function(e) {
    const data = new Uint8Array(reader.response);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Convert worksheet data to JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Search for the record
    const matchedRecord = searchForRecord(jsonData, enroll, grno);

    // Display the result
    displayResult(matchedRecord);
  };

  reader.send();
}

// Function to search for a match in the Excel data
function searchForRecord(data, enroll, grno) {
  for (let i = 1; i < data.length; i++) { // Skip header row
    const row = data[i];
    if (row[0] == enroll && row[1] == grno) {
      return { name: row[2], marks: row[3] };
    }
  }
  return null;
}

// Function to display the result on the webpage
function displayResult(record) {
  const resultDiv = document.getElementById('result');
  resultDiv.innerHTML = ''; // Clear previous result

  if (record) {
    resultDiv.innerHTML = `<h5>Result Found</h5><br>
                           <h6>Name: ${record.name}</h6>
                           <h6>Marks: ${record.marks}</h6>`;
  } else {
    resultDiv.innerHTML = `<p class="text-danger">No matching record found!!</p>`;
  }
}

