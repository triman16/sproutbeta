const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const xlsx = require('xlsx');

const app = express();
const port = 3000;
const excelFilePath = 'leads.xlsx';

app.use(cors());
app.use(bodyParser.json());

// Function to read emails from Excel file
const readEmailsFromExcel = () => {
    if (fs.existsSync(excelFilePath)) {
        const workbook = xlsx.readFile(excelFilePath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        return data.map(row => row[0]);
    }
    return [];
};

// Function to write emails to Excel file
const writeEmailsToExcel = (emails) => {
    const workbook = xlsx.utils.book_new();
    const sheet = xlsx.utils.aoa_to_sheet(emails.map(email => [email]));
    xlsx.utils.book_append_sheet(workbook, sheet, 'Leads');
    xlsx.writeFile(workbook, excelFilePath);
    console.log('Excel file updated:', emails);
};

let emails = readEmailsFromExcel();
console.log('Initial emails:', emails);

app.post('/subscribe', (req, res) => {
    const email = req.body.email;
    console.log('Received email:', email);
    if (email) {
        if (!emails.includes(email)) {
            emails.push(email);
            writeEmailsToExcel(emails);
            res.status(200).send({ message: 'Congrats! You are in, first newsletter coming soon.' });
        } else {
            res.status(400).send({ message: 'Email is already subscribed!' });
        }
    } else {
        res.status(400).send({ message: 'Email is required!' });
    }
});

app.get('/subscribers', (req, res) => {
    res.status(200).json(emails);
});

app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
