const express = require('express');
const multer = require('multer');
const mammoth = require('mammoth');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
app.use(cors());

const upload = multer();

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const result = await mammoth.extractRawText({ buffer: req.file.buffer });
    const text = result.value;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');
    const rows = text.split('\n');

    rows.forEach((row) => {
      const columns = row.split('\t');
      worksheet.addRow(columns);
    });

    const excelBuffer = await workbook.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=data.xlsx');
    res.send(excelBuffer);
  } catch (error) {
    console.error(error);
    res.status(500).send('เกิดข้อผิดพลาดในการอัปโหลดไฟล์');
  }
});

app.listen(4000, () => {
  console.log('เซิร์ฟเวอร์เริ่มต้นที่พอร์ต 4000');
});

module.exports = app;
