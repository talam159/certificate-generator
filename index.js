const { signaturee } = require('./signature.js') //SVG code of the two signatures.

const fs = require('fs');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const SVGtoPDF = require('svg-to-pdfkit');

const data = [];
const workbook = XLSX.readFile('data.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
XLSX.utils.sheet_to_json(worksheet).forEach(row => {
    data.push({
        name: row.Name,
        rollNumber: row['RollNumber'],
        gpa: row.GPA,
        department: row.department,
        degree: row.Degree,
        grade: row.letterGrade,
        id: row.RollNumber
    });
});

function generateCertificate(student){

  const doc = new PDFDocument({ size: 'A4', layout: 'landscape' }); // Customize size and margin

    // Resize the document canvas
    // doc.scale(0.75); // Scale down by 75% (adjust as needed)
    // const outputFolderPath = 'certificates'; // Folder path to save certificates
    // const pdfOutputPath = path.join(outputFolderPath, `certificate_${student.name}.pdf`);
    const pdfOutputPath = `certificate_${student.name}_${student.rollNumber}.pdf`;
    doc.pipe(fs.createWriteStream(pdfOutputPath));
    if (fs.existsSync(pdfOutputPath)) {
        fs.unlinkSync(pdfOutputPath); // Delete existing certificate file
    }
  
  // Helper to move to next line
  function jumpLine(doc, lines) {
    for (let index = 0; index < lines; index++) {
      doc.moveDown();
    }
  }
  

  // doc.image('bg.png', 0, 0); //##### for background. Waiting for the raw file for the design
  
  
  // doc.rect(0, 0, doc.page.width, doc.page.height).fill('#fff');
  
  doc.fontSize(10);
  
  // Margin
  const distanceMargin = 18;
  
  // doc
  //   .fillAndStroke('#1162a0')
  //   .lineWidth(5)
  //   .lineJoin('round')
  //   .rect(
  //     distanceMargin,
  //     distanceMargin,
  //     doc.page.width - distanceMargin * 2,
  //     doc.page.height - distanceMargin * 2,
  //   )
  //   .stroke();
  
  // Header
  const maxWidth = 140;
  const maxHeight = 70;
  
  
  
    jumpLine(doc,0);
  
    
  doc.moveDown(11.7);
  
  doc
    .font('fonts/monotype-corsiva-bold.otf')
    .fontSize(27)
    .fill('#021c27')
    .text(`${student.name}`, {//##Name##
      align: 'center',
    });
    doc
    .font('fonts/monotype-corsiva-bold.otf')
    .fontSize(16)
    .fill('#021c27')
    .text(`ID NO. ${student.id}`, {//##Name##
      align: 'center',
    });
    
    doc.moveDown(1.2);

    

  doc
    .font('fonts/Monotype-Corsiva.ttf')
    .fontSize(24)
    .fill('#021c27')
    .text(`${student.degree}`, {//##degree##
      align: 'center',
    });
    doc.moveDown(0.5);


    
    doc
    .font('fonts/Times-New-Roman-Bold.ttf')
    .fontSize(13)
    .fill('#021c27')
    .text(`with a CGPA ${student.gpa} in a scale of 4.00`, {
      align: 'center',
    });
    doc.moveDown(1);
    doc
    .font('fonts/TIMES.ttf')
    .fontSize(10)
    .fill('#021c27')
    .text('WITH ALL RIGHTS, PRIVILEGES AND OBLIGATIONS APPERTAINING THERETO', {
      align: 'center',
    });
    doc.moveDown(0.4);
    doc
    .font('fonts/TIMES.ttf')
    .fontSize(10)
    .fill('#021c27')
    .text('AWARDED UNDER THE SEAL OF CITY UNIVERSITY, DHAKA, BANGLADESH,', {
      align: 'center',
    });
    doc.moveDown(0.4);
    doc
    .font('fonts/TIMES.ttf')
    .fontSize(10)
    .fill('#021c27')
    .text('ON THIS ELEVENTH DAY OF MAY, TWO THOUSAND AND TWENTY FOUR.', {
      align: 'center',
    });
  
 
  
  
  SVGtoPDF(doc, signature, 0, 205);
  
  doc.lineWidth(1);
  
  
  doc
  .font('fonts/NotoSansJP-Light.otf')
  .fontSize(9)
  .text('This certificate is digitally signed through RVLCA, licensed Certifying Auhority of  Government of Bangladesh', 200, doc.page.height - 20, {
    lineBreak: false
  });
  

  const bottomHeight = doc.page.height - 500;
  
  
  doc.end();
  console.log(`Certificate generated successfully: ${pdfOutputPath}`);
  
}
data.forEach(student => {
  generateCertificate(student);
});
