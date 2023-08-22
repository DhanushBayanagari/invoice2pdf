const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const pug=require('pug')
const fs = require('fs');
const path = require('path');
const mysql=require('mysql');
// const fs = require('fs');
const xlsx = require('xlsx');
const { execFile } = require('child_process');
let excelFilePath;

const app = express();
const upload = multer({ dest: 'uploads/' });
app.use(express.static(path.join(__dirname,'assets')));
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'pug');


app.get('/', (req, res) => {
  res.render('index');
});

app.get('/login', (req, res) => {
  res.render('login');
});

// Load the XLSX file
// const workbook = xlsx.readFile('i2p.xlsx');
// const worksheet = workbook.Sheets[workbook.SheetNames[0]];
// const data = xlsx.utils.sheet_to_json(worksheet);

// Function to read data from Excel file
function readExcelFile(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(worksheet, { header: 1 });
}


//connect to database
const conn=mysql.createConnection(
  {
    host:"localhost",
    user:"root",
    pass:"",
    database:"i2p"

  }
)

conn.connect((err)=>{
  if(err) throw err;
  else console.log("database connected")
 
  app.post('/generate-invoice', upload.single('file'), (req, res) => {
      if (!req.file) {
        return res.status(400).send('No file uploaded.');
      }
    
      // const workbook = new ExcelJS.Workbook();
      // workbook.xlsx
      //   .readFile(req.file.path)
      //   .then(() => {
        excelFilePath=req.file.path;

        const data = readExcelFile(excelFilePath);

  // Assuming the first row in the Excel file contains column names
  const columns = data[0];

  // Remove the header row from the data array
  data.shift();

  // Insert each row into the database
  for (const row of data) {
    const rowData = {};

    for (let i = 0; i < columns.length; i++) {
      // Assuming the columns in the Excel file match the database table column names
      rowData[columns[i]] = row[i];
    }

    // Assuming your_table_name is the name of your database table
    const query = 'INSERT INTO excel SET ?';

    conn.query(query, rowData, (err, result) => {
      if (err) {
        console.error('Error inserting row:', err.message);
        return;
      }

      console.log('Row inserted:', result.insertId);
    });
  }


  // Close the database connection
  conn.end((err) => {
    if (err) {
      console.error('Error closing database connection:', err.message);
    } else {
      console.log('Database connection closed');
    }
        
  console.log('Connected to the database');
});
  })
  //  excelFilePath ='i2p.xlsx';
  // console.log(excelFilePath)
  // const data = readExcelFile(excelFilePath);

  // // Assuming the first row in the Excel file contains column names
  // const columns = data[0];

  // // Remove the header row from the data array
  // data.shift();

  // // Insert each row into the database
  // for (const row of data) {
  //   const rowData = {};

  //   for (let i = 0; i < columns.length; i++) {
  //     // Assuming the columns in the Excel file match the database table column names
  //     rowData[columns[i]] = row[i];
  //   }

  //   // Assuming your_table_name is the name of your database table
  //   const query = 'INSERT INTO excel SET ?';

  //   conn.query(query, rowData, (err, result) => {
  //     if (err) {
  //       console.error('Error inserting row:', err.message);
  //       return;
  //     }

  //     console.log('Row inserted:', result.insertId);
  //   });
  // }


  // // Close the database connection
  // conn.end((err) => {
  //   if (err) {
  //     console.error('Error closing database connection:', err.message);
  //   } else {
  //     console.log('Database connection closed');
  //   }
  });
// })
    // Insert data into the database table
    // const tableName = 'excel';

    // const insertQuery = `INSERT INTO ${tableName} `;
    // const values = data.map((row) => [row.column1, row.column2]);
  
    // conn.query(insertQuery, [values], (err, result) => {
    //   if (err) throw err;
    //   console.log(`Inserted ${result.affectedRows} rows into the database.`);
      
      // Close the database connection
      //conn.end();
  //   });

  


// var sql = "CREATE TABLE excel (name VARCHAR(255), amount VARCHAR(255))";
// conn.query(sql, function (err, result) {
//   if (err) throw err;
//   console.log("Table created");
// });







// app.post('/generate-invoice', upload.single('file'), (req, res) => {
//   if (!req.file) {
//     return res.status(400).send('No file uploaded.');
//   }

//   const workbook = new ExcelJS.Workbook();
//   workbook.xlsx
//     .readFile(req.file.path)
//     .then(() => {
//       const invoices = [];

//       for (let i = 2; i < 7; i++) {
//         const worksheet = workbook.worksheets[0];
//         const invoiceNumber = worksheet.getCell(`A${i}`).value;
//         const customerName = worksheet.getCell(`B${i}`).value;
//         //const totalAmount = worksheet.getCell(`C${i}`).value;

//         const doc = new PDFDocument();
//         doc.pipe(fs.createWriteStream(`invoice-${invoiceNumber}.pdf`));

//         // doc.fontSize(20).text('Invoice', { align: 'center' });
//         // doc.fontSize(14).text(`Invoice Number: ${invoiceNumber}`);
//         // doc.fontSize(14).text(`Customer Name: ${customerName}`);
//         // doc.fontSize(14).text(`Total Amount: ${totalAmount}`);

//         const fontSize = 15;
//         let position = 20;
      
      
//         doc.text(`INVOICE`, 200,( position=+10));
//         doc.text(`Invoice Date: 18 May 2023`, 180, (position += 20));
//         doc.text(`Due on Receipt: 18 May 2023`, 180, (position += 20));

//         doc.text(`Zylker Electronics Hub`, 20, (position+=30));
//         doc.text(`141, Northern Street`, 20, (position += 20));
//         doc.text(`Greater South Avenue`, 20, (position += 20));
//         doc.text(`New York, New York 10001`, 20, (position += 20));
//         doc.text(`USA`, 20, (position += 20));
      
//         // Add the bill-to details
//         position += 60;
//         doc.text(`Bill To`, 20, (position+=30));
//         doc.text(`Ship To`, 400, position);
//         doc.text(`Ms. Mary D. Dunton`, 20, (position += 20));
//         doc.text(`1324 Hinkle Lake Road`, 20, (position += 20));
//         doc.text(`Needham 02192 Maine`, 20, (position += 20));
//         doc.text(`Ms. Mary D. Dunton`, 370, (position-=40));
//         doc.text(`1324 Hinkle Lake Road`, 370, (position += 20));
//         doc.text(`Needham 02192 Maine`, 370, (position += 20));
      
//         // Add the item table
//         position += 50;
//         doc.text(`Item & Description`, 20, position);
//         doc.text(`Qty`, 380, position);
//         doc.text(`Rate`, 420, position);
//         doc.text(`Amount`, 470, position);
        
//         position += 10;
//         doc.text("OSLR camera with advanced shooting capabilities", 20, (position+=20));
//         doc.text("1", 390, position);
//         doc.text("899.00", 410, position);
//         doc.text("899.00",480, position);
      
//         position += 20;
//         doc.text("Activity tracker with heart rate monitoring", 20, position);
//         doc.text("1", 390, position);
//         doc.text("129.00", 410, position);
//         doc.text("129.00", 480, position);
      
//         position += 20;
//         doc.text("Lightweight laptop with a powerful processor", 20, position);
//         doc.text("1", 390, position);
//         doc.text("999.00", 410, position);
//         doc.text("999.00", 480, position);
      
//         // Add the subtotal, tax rate, and total balance due
//         const subtotal = 2027.00;
//         const taxRate = 5.00;
//         const totalAmount = 2128.35;
      
//         doc.text("Sub Total", 400, (position+=20));
//         doc.text(subtotal.toFixed(2), 480, position);
      
//         doc.text("Tax Rate", 400, (position += 20));
//         doc.text(taxRate.toFixed(2), 480, position);
      
//         doc.text("Total Balance Due", 340, (position += 20));
//         doc.text(totalAmount.toFixed(2),480, position);
      
//         doc.end();

//         invoices.push(`invoice-${invoiceNumber}.pdf`);
//       }

//       res.send('Invoices generated successfully: ' + invoices.join(', '));
//     })
//     .catch((error) => {
//       console.log('Error reading Excel file:', error);
//       res.status(500).send('Error generating invoices.');
//     });
// });

app.listen(8000, () => {
  console.log('Server is listening on port 8000.');
});
