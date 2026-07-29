const PDFDocument = require('pdfkit-table');
const fs = require('fs');

async function testPdf() {
  try {
    const doc = new PDFDocument({ margin: 30, size: 'A4' });
    doc.pipe(fs.createWriteStream('test.pdf'));
    
    doc.fontSize(16).text('Header Text', { align: 'center' });
    doc.moveDown(2);

    const tableRows = [];
    tableRows.push([
      '', 
      '', 
      '', 
      'Saldo Awal', 
      '', 
      '', 
      '', 
      'Rp 0'
    ]);
    
    tableRows.push([
      '1',
      '12/12/2026',
      'Test',
      'Test',
      'Test',
      'Rp 100.000',
      '',
      'Rp 100.000'
    ]);

    tableRows.push([
      '', 
      '', 
      '', 
      'TOTAL', 
      '', 
      'Rp 100.000', 
      'Rp 0', 
      'Rp 100.000'
    ]);

    const table = {
      headers: [
        { label: "No", width: 25 },
        { label: "Tanggal", width: 55 },
        { label: "Kategori", width: 70 },
        { label: "Uraian", width: 100 },
        { label: "Penerima", width: 65 },
        { label: "Pemasukan", width: 75 },
        { label: "Pengeluaran", width: 75 },
        { label: "Saldo", width: 70 },
      ],
      rows: tableRows,
    };

    await doc.table(table, {
      prepareHeader: () => doc.font("Helvetica-Bold").fontSize(9),
      prepareRow: (row, indexColumn, indexRow, rectRow, rectCell) => {
        doc.font("Helvetica").fontSize(8);
      },
    });

    doc.end();
    console.log("PDF generated successfully");
  } catch (error) {
    console.error("Error generating PDF:", error);
  }
}

testPdf();
