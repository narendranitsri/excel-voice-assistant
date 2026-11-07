import fs from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

const samplesDir = path.resolve('./samples');
const photosDir = path.join(samplesDir, 'photos');
fs.mkdirSync(photosDir, { recursive: true });

// 1. Create samples/sample_applicants.xlsx with 2 sheets
const registrationRows = [
  {
    Serial_No: '101',
    Application_No: 'RAJ-2025-101',
    Applicant_Name: 'Ramesh Kumar',
    Father_Name: 'Mohan Lal',
    DOB: '15/08/1992',
    Gender: 'M',
    Mobile: '9876543210',
    Aadhaar: '456789012345',
    Category: 'OBC',
    PIN_Code: '302029',
    District: 'Jaipur',
    Block: 'Sanganer',
    Village: 'Muhana',
    Annual_Income: 45000,
    Photo_File: '101_Ramesh_Kumar.jpg',
    Address: 'Ward 4, Muhana Mandi Road, Sanganer',
    Status: 'Submitted',
    Amount: '12500'
  },
  {
    Serial_No: '102',
    Application_No: 'RAJ-2025-102',
    Applicant_Name: 'Sunita Devi',
    Father_Name: 'Sohan Lal',
    DOB: '20/11/1995',
    Gender: 'Female',
    Mobile: '9123456789',
    Aadhaar: '789012345678',
    Category: 'General',
    PIN_Code: '342001',
    District: 'Jodhpur',
    Block: 'Mandore',
    Village: 'Boranada',
    Annual_Income: 38000,
    Photo_File: '102_Sunita_Devi.jpg',
    Address: 'Plot 18, Boranada Phase 2',
    Status: 'Approved',
    Amount: '15000'
  },
  {
    Serial_No: '103',
    Application_No: 'RAJ-2025-103',
    Applicant_Name: 'Mahesh Sharma',
    Father_Name: 'Kailash Sharma',
    DOB: '05/03/1989',
    Gender: 'Male',
    Mobile: '9988776655',
    Aadhaar: '654321098765',
    Category: 'General',
    PIN_Code: '313001',
    District: 'Udaipur',
    Block: 'Girwa',
    Village: 'Bedla',
    Annual_Income: 52000,
    Photo_File: '103_Mahesh_Sharma.jpg',
    Address: 'Near Govt School, Bedla',
    Status: 'Pending Verification',
    Amount: '10000'
  },
  {
    Serial_No: '104',
    Application_No: 'RAJ-2025-104',
    Applicant_Name: 'Kavita Meena',
    Father_Name: 'Ramकिशन मीणा',
    DOB: '1997-12-01',
    Gender: 'F',
    Mobile: '98765', // Intentional invalid 5-digit mobile to test validator.py
    Aadhaar: '12345678', // Intentional invalid 8-digit Aadhaar to test validator.py
    Category: 'ST',
    PIN_Code: '3020', // Intentional invalid 4-digit PIN to test validator.py
    District: 'Jaipur',
    Block: 'Amer',
    Village: 'Achrol',
    Annual_Income: 40000,
    Photo_File: '104_Kavita_Meena.jpg',
    Address: 'Main Bazaar, Achrol',
    Status: 'Pending',
    Amount: '12500'
  }
];

const statusSearchRows = [
  { Application_No: 'RAJ-2025-101', Applicant_Name: 'Ramesh Kumar', District: 'Jaipur' },
  { Application_No: 'RAJ-2025-102', Applicant_Name: 'Sunita Devi', District: 'Jodhpur' },
  { Application_No: 'RAJ-2025-103', Applicant_Name: 'Mahesh Sharma', District: 'Udaipur' },
  { Application_No: 'RAJ-2025-104', Applicant_Name: 'Kavita Meena', District: 'Jaipur' }
];

const wb = XLSX.utils.book_new();
const ws1 = XLSX.utils.json_to_sheet(registrationRows);
const ws2 = XLSX.utils.json_to_sheet(statusSearchRows);
XLSX.utils.book_append_sheet(wb, ws1, 'Applicants_Registration');
XLSX.utils.book_append_sheet(wb, ws2, 'Status_Search_IDs');
XLSX.writeFile(wb, path.join(samplesDir, 'sample_applicants.xlsx'));

// 2. Create a valid Word (.docx) file containing both "Label: Value" paragraphs and a Table
async function createSampleDocx() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.folder('_rels').file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.folder('word').file(
    'document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Applicant_Name: Ramesh Kumar</w:t></w:r></w:p>
    <w:p><w:r><w:t>Father_Name: Mohan Lal</w:t></w:r></w:p>
    <w:p><w:r><w:t>DOB: 15/08/1992</w:t></w:r></w:p>
    <w:p><w:r><w:t>Gender: Male</w:t></w:r></w:p>
    <w:p><w:r><w:t>Mobile: 9876543210</w:t></w:r></w:p>
    <w:p><w:r><w:t>Aadhaar: 456789012345</w:t></w:r></w:p>
    <w:p><w:r><w:t>District: Jaipur</w:t></w:r></w:p>
    <w:p><w:r><w:t>Block: Sanganer</w:t></w:r></w:p>
    <w:p><w:r><w:t>Village: Muhana</w:t></w:r></w:p>
    <w:p><w:r><w:t>PIN_Code: 302029</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Applicant_Name</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Mobile</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>District</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Block</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Ramesh Kumar</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>9876543210</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Jaipur</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Sanganer</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Sunita Devi</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>9123456789</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Jodhpur</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Mandore</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`
  );
  const content = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path.join(samplesDir, 'sample_applicants.docx'), content);
}

// 3. Create valid minimal JPEG files for sample applicant photos
const minimalJpegHex =
  'ffd8ffe000104a46494600010101006000600000ffdb004300080606070605080707070909080a0c140d0c0b0b0c1912130f141d1a1f1e1d1a1c1c20242e2720222c231c1c2837292c30313434341f27393d38323c2e333432ffc0000b080001000101011100ffc4001f0000010501010101010100000000000000000102030405060708090a0bffe40014100100000000000000000000000000000000ffda0008010100003f00fbfdffd9';
const jpegBuffer = Buffer.from(minimalJpegHex, 'hex');
['101_Ramesh_Kumar.jpg', '102_Sunita_Devi.jpg', '103_Mahesh_Sharma.jpg', '456789012345.jpg'].forEach((fname) => {
  fs.writeFileSync(path.join(photosDir, fname), jpegBuffer);
});

await createSampleDocx();
console.log('Generated sample_applicants.xlsx, sample_applicants.docx, and sample photos successfully!');
