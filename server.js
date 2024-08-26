const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const moment = require('moment-timezone');
const excelJS = require('exceljs');
const nodemailer = require('nodemailer');
const schedule = require('node-schedule');

const app = express();
const PORT = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));

// Utility function to get the current time in BST and format it for filenames
function getCurrentBSTDateTime(format = 'YYYY-MM-DD_HH-mm') {
    return moment.tz('Asia/Dhaka').format(format);
}

// Utility function to create or load the Excel file
function getExcelFilePath() {
    const date = getCurrentBSTDateTime('YYYY-MM-DD');
    return path.join(__dirname, `${date}_orders.xlsx`);
}

function createOrUpdateExcel(order) {
    const filePath = getExcelFilePath();
    const workbook = new excelJS.Workbook();

    if (fs.existsSync(filePath)) {
        // Load existing workbook if the file already exists
        return workbook.xlsx.readFile(filePath).then(() => {
            const worksheet = workbook.getWorksheet(1);
            worksheet.addRow(order).commit();
            return workbook.xlsx.writeFile(filePath);
        });
    } else {
        // Create a new workbook if the file does not exist
        const worksheet = workbook.addWorksheet('Orders');
        worksheet.columns = [
            { header: 'Name', key: 'name' },
            { header: 'Mobile', key: 'mobile' },
            { header: 'Address', key: 'address' },
            { header: 'Product', key: 'product' },
            { header: 'Quantity', key: 'quantity' },
            { header: 'Order Time', key: 'order_time' },
        ];
        worksheet.addRow(order).commit();
        return workbook.xlsx.writeFile(filePath);
    }
}

// Route to handle form submission
app.post('/submit-order', (req, res) => {
    try {
        const { name, mobile, address, product, quantity } = req.body;
        const orderTime = getCurrentBSTDateTime('YYYY-MM-DD HH:mm:ss');

        const order = {
            name,
            mobile,
            address,
            product: product || 'Not specified',
            quantity: quantity || 1,
            order_time: orderTime,
        };

        createOrUpdateExcel(order).then(() => {
            res.send('Order has been recorded!');
        }).catch(error => {
            console.error('Error updating Excel file:', error);
            res.status(500).send('An error occurred while processing your order.');
        });
    } catch (error) {
        console.error('Error processing order:', error);
        res.status(500).send('An error occurred while processing your order.');
    }
});

// Email configuration
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: 'your-email@gmail.com',
        pass: 'your-email-password',
    },
});

// Function to send the Excel file via email
function sendOrderFile() {
    const filePath = getExcelFilePath();
    const mailOptions = {
        from: 'your-email@gmail.com',
        to: 'recipient-email@gmail.com',
        subject: `Order Report - ${getCurrentBSTDateTime('YYYY-MM-DD')}`,
        text: 'Please find the attached order report.',
        attachments: [{ filename: path.basename(filePath), path: filePath }],
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            return console.error('Error sending email:', error);
        }
        console.log('Email sent:', info.response);

        // Delete the file immediately after sending the email
        fs.unlink(filePath, err => {
            if (err) {
                return console.error('Error deleting file:', err);
            }
            console.log('File deleted:', filePath);
        });
    });
}

// Schedule the email to be sent every day at 12:00 PM BST
schedule.scheduleJob({ hour: 12, minute: 0, tz: 'Asia/Dhaka' }, sendOrderFile);

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
