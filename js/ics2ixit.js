function naturalSort(a, b) {
  const regex = /(\d+)/g;
  const aNumbers = a.match(regex) || [];
  const bNumbers = b.match(regex) || [];

  for (let i = 0; i < Math.max(aNumbers.length, bNumbers.length); i++) {
    const aNum = parseInt(aNumbers[i] || 0, 10);
    const bNum = parseInt(bNumbers[i] || 0, 10);

    if (aNum !== bNum) {
      return aNum - bNum;
    }
  }

  return a.localeCompare(b);
}


function consolidateIXITs() {
  const sectionIXIT = document.getElementById('consolidatedIXITs');

  // Clear the existing content
  sectionIXIT.innerHTML = '';

  const ixitMap = new Map();

  const rows = document.querySelectorAll('tr');
  rows.forEach(row => {
    const selectElement = row.querySelector('.ixittodo');

    if (selectElement && selectElement.value === 'Yes') {
      const ics = row.id;
      const ixitElements = row.querySelectorAll('[class^="IXIT_"]');

      ixitElements.forEach(ixitElement => {
        const className = ixitElement.className;
        const parts = ixitElement.textContent.split(': ')[1].split(', ');

        if (!ixitMap.has(className)) {
          ixitMap.set(className, { ics: new Set(), parts: new Set(), ixitElements: [] });
        }

        parts.forEach(part => {
          ixitMap.get(className).parts.add(part.trim());
        });
        ixitMap.get(className).ics.add(ics);
        ixitMap.get(className).ixitElements.push(ixitElement);
      });
    }
  });

  // Construct the consolidated table
  const table = document.createElement('table');
  table.classList.add('consolidated-table');

  // Create table header row
  const headerRow = document.createElement('tr');
  const headerCell1 = document.createElement('th');
  const headerCell2 = document.createElement('th');
  const headerCell3 = document.createElement('th');
  headerCell1.textContent = 'ID';
  headerCell2.textContent = 'Required Documentation';
  headerCell3.textContent = 'ICS';
  headerRow.appendChild(headerCell1);
  headerRow.appendChild(headerCell2);
  headerRow.appendChild(headerCell3);
  table.appendChild(headerRow);

  // Add table body rows
  ixitMap.forEach(({ parts, ics, ixitElements }, className) => {
    const consolidatedPart = Array.from(parts).join(', ');

    const row = document.createElement('tr');
    const cell1 = document.createElement('td');
    const cell2 = document.createElement('td');
    const cell3 = document.createElement('td');

    cell1.textContent = className.replace('_', ' ');
    cell2.textContent = consolidatedPart;

    // Create a span for each ICS and set the tooltip
    ics.forEach(ic => {
      const span = document.createElement('span');
      span.textContent = ic;

      // Set the tooltip with relevant IXIT information for each individual ICS
      if (ic) {
        const block = document.getElementById(ic);
        if (block && className) {
            span.title = block.querySelector('.' + className).innerHTML.split(': ')[1];
        }
      }

      cell3.appendChild(span);
      cell3.appendChild(document.createTextNode(', '));
    });

    cell3.removeChild(cell3.lastChild);

    row.appendChild(cell1);
    row.appendChild(cell2);
    row.appendChild(cell3);
    table.appendChild(row);
  });

  sectionIXIT.appendChild(table);

  // Sort the table rows using natural sorting
  const tableRows = table.querySelectorAll('tr');
  const sortedRows = Array.from(tableRows).sort((a, b) => {
    const aText = a.cells[0].textContent.trim();
    const bText = b.cells[0].textContent.trim();
    return naturalSort(aText, bText);
  });

  sortedRows.forEach(row => table.appendChild(row));
}

function generateExcel() {
  const workbook = XLSX.utils.book_new();
  const summarySheetData = [];

  const rows = document.querySelectorAll('.consolidated-table tr:not(:first-child)');

  rows.forEach(row => {
    const cells = row.querySelectorAll('td');
    if (cells.length >= 2) {
      const id = cells[0].textContent;
      const documentation = cells[1].textContent.split(', ');
      const ics = cells[2].textContent;


      summarySheetData.push({ 'IXIT ID': id, 'Required Documentation': documentation.join(', '), 'Associated ICS': ics });
    }
  });

  
  const summarySheet = XLSX.utils.json_to_sheet(summarySheetData);
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');  

  
  rows.forEach(row => {
    const cells = row.querySelectorAll('td');
    if (cells.length >= 2) {
      const id = cells[0].textContent;
      const documentation = cells[1].textContent.split(', ');

      const sheetName = id;
      const sheetData = [{}]; // Create an empty object for the first row

      documentation.forEach((part, index) => {
        sheetData[0][`A${index + 1}`] = part;
      });

      const sheet = XLSX.utils.json_to_sheet(sheetData, {skipHeader:true});
      XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
    }
  });


  XLSX.writeFile(workbook, 'cetome - ICS2IXIT.xlsx');
}

/* Handle XLSX import*/
async function handleFile(file) {
  try {
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);

    reader.onload = async (event) => {
      const arrayBuffer = event.target.result;
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = 'ICS';
      const worksheet = workbook.Sheets[sheetName];

      // Adjust the range to start from row 2 (index 1) to row 85 (index 84). ONLY FOR BSI FILE!
      const range = { s: { r: 1, c: 0 }, e: { r: 84, c: worksheet['!ref'].split(':')[1].charCodeAt(0) - 65 } };
      const data = XLSX.utils.sheet_to_json(worksheet, { range });

      data.forEach(row => {
        const icsValue = row['Provision'].replace(/^6\./, '6-');;
        const supportValue = row['Support'];
        const detailValue = row['Detail'];//\n(In case of "Yes" in the support column, the corresponding IXIT entries are used for details)'];

        const rowElement = document.getElementById(icsValue);
        if (rowElement) {
          const selectElement = rowElement.querySelector('select.ixittodo');
          selectElement.value = supportValue;
        
          if (detailValue !== undefined && detailValue !== '') {
              const detailElement = rowElement.querySelector('textarea');
              detailElement.value = detailValue.trim();
            }
          
          colorize({ target: selectElement });
        }
      });

      consolidateIXITs();
      runColorize(); 
      
      scrollDownTo('consolidatedIXITs');

    };
  } catch (error) {
    console.error('Error importing Excel file:', error);
    alert('Error importing Excel file: ' + error.message);
  }
}

function scrollDownTo(id) {
    const goTo = document.getElementById(id);
    goTo.scrollIntoView({ behavior: 'smooth' });
}

// Event listeners for drag and drop
const dropContainer = document.getElementById("dropcontainer");

dropContainer.addEventListener("dragover", (e) => {
  e.preventDefault();
}, false);

dropContainer.addEventListener("dragenter", () => {
  dropContainer.classList.add("drag-active");
});

dropContainer.addEventListener("dragleave", () => {
  dropContainer.classList.remove("drag-active");
});

dropContainer.addEventListener("drop", (e) => {
  e.preventDefault();
  dropContainer.classList.remove("drag-active");
  handleFile(e.dataTransfer.files[0]);
});

// Event listener for browse button
document.getElementById('ICSfile').onchange = function(event) {
  handleFile(event.target.files[0]);
};


// Add a button to trigger Excel generation
const generateExcelButton = document.getElementById('generate-excel');
generateExcelButton.addEventListener('click', generateExcel);

function colorize(event) {
  const selectedOption = event.target.options[event.target.selectedIndex];
  const action = selectedOption.value;
  const originalBg = event.target.parentNode.parentNode.style.backgroundColor;

  // event.target.style.backgroundColor = "white";
  // event.target.style.color = "black";

  let bgColor;

  switch (action) {
    case "Yes":
      bgColor = "#aea";
      break;
    case "No":
      bgColor = "#eaa";
      break;
    case "N/A":
      bgColor = "#aaa";
      break;
    default:
      bgColor = originalBg;
  }

  event.target.parentNode.parentNode.style.backgroundColor = bgColor;
}

const selects = document.querySelectorAll('select');
selects.forEach(select => {
  select.addEventListener('change', colorize);
});

// Add event listeners to all select elements
const selectElements = document.querySelectorAll('.ixittodo');

selectElements.forEach(selectElement => {
  selectElement.addEventListener('change', consolidateIXITs);
});

function runColorize() {
const selects = document.querySelectorAll('select');
  selects.forEach(select => {
    select.addEventListener('change', colorize);
    // Trigger the colorize function initially
    colorize({ target: select });
  });
}

// Call the function on page load
document.addEventListener('DOMContentLoaded', () => {
  consolidateIXITs();
  runColorize(); 
});