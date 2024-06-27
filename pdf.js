const fs = require('fs');
const pdf = require('pdf-parse');
const XLSX = require('xlsx');

async function pdfToExcel(pdfPath, excelPath) {
  try {
    
    const pdfBuffer = fs.readFileSync(pdfPath);
    const data = await pdf(pdfBuffer);

    
    const lines = data.text.split('\n').filter(line => line.trim() !== '');

    
    const headers = ['Sr. No.', 'AIR', 'NEET Roll No.', 'CET Form No.', 'Reg. Sr No.', 'Name', 'G', 'Cat', 'Quota', 'Code College'];
    const rows = [];

    let isTableData = false;

    lines.forEach(line => {
      const columns = line.trim().split(/\s+/);

      
      if (columns.length > 0 && columns[0].match(/^\d+$/)) {
        isTableData = true;
      }

      if (isTableData && columns.length >= 10) {
        const sr_no = columns[0];
        const air = columns[1];
        const neet_roll_no = columns[2];
        const cet_form_no = columns[3];
        const reg_sr_no = columns[4];

        
        let name = '';
        let nameEndIndex = 5;
        while (nameEndIndex < columns.length && columns[nameEndIndex].length > 1) {
          name += columns[nameEndIndex] + ' ';
          nameEndIndex++;
          if (nameEndIndex - 5 > 4) break; 
        }
        name = name.trim();

        const g = columns[nameEndIndex] || ''; 
        const cat = columns[nameEndIndex + 1] || ''; 
        const quota = columns[nameEndIndex + 2] || ''; 
        const code_college = columns[nameEndIndex + 3] || ''; 

        
        if ([sr_no, air, neet_roll_no, cet_form_no, reg_sr_no, name, g, cat, quota, code_college].length === 10) {
          rows.push([sr_no, air, neet_roll_no, cet_form_no, reg_sr_no, name, g, cat, quota, code_college]);
        }
      }

      
      if (isTableData && columns.length < 10) {
        isTableData = false;
      }
    });

    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

    
    const columnWidths = headers.map(header => ({ wch: header.length + 5 }));
    ws['!cols'] = columnWidths;

    
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    
    XLSX.writeFile(wb, excelPath);

    console.log("PDF data has been converted to Excel successfully");
  } catch (error) {
    console.error('Error converting PDF to Excel:', error);
  }
}


const pdfPath = 'NEET-Selection-MOP2.pdf';
const excelPath = 'NEET-Selection-MOP2.xlsx';
pdfToExcel(pdfPath, excelPath);