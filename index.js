"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const nodemailer = __importStar(require("nodemailer"));
const xlsx = __importStar(require("xlsx"));
const fs = __importStar(require("fs"));
// Sample data to be written to the Excel file
const recruiters = [
    { Name: "John Doe", Email: "exwhyzed884@gmail.com", Company: "Company A" },
];
// Function to create an XLSX file
const createExcelFile = (data, filePath) => {
    // Convert the data to a worksheet
    const worksheet = xlsx.utils.json_to_sheet(data);
    // Create a new workbook and append the worksheet
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, "Recruiters");
    // Write the workbook to a file
    xlsx.writeFile(workbook, filePath);
};
// Specify the file path for the new Excel file
const filePath = "./recruiters.xlsx";
// Create the Excel file with the sample data if it doesn't exist
if (!fs.existsSync(filePath)) {
    createExcelFile(recruiters, filePath);
    console.log(`Excel file created at ${filePath}`);
}
else {
    console.log(`Excel file already exists at ${filePath}`);
}
// Email configuration
const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
        user: "vs630916@gmail.com",
        pass: "vzpw zwjg moou kvcy",
    },
});
// Function to read Excel file
const readExcel = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet);
};
// Function to send email
const sendEmail = (to, subject, text) => {
    const mailOptions = {
        from: "vs630916@gmail.com",
        to,
        subject,
        text,
    };
    return new Promise((resolve, reject) => {
        transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
                return reject(error);
            }
            resolve(info);
        });
    });
};
// Main function to automate the process
const automateEmails = () => __awaiter(void 0, void 0, void 0, function* () {
    if (!fs.existsSync(filePath)) {
        console.error("Excel file not found.");
        return;
    }
    const recruiters = readExcel(filePath);
    for (const recruiter of recruiters) {
        const { Email, Name, Company } = recruiter;
        const subject = "Job Application for [Position]";
        const text = `Hi ${Name},

I hope this email finds you well. I am writing to express my interest in a potential role at ${Company}. 

[Include more details about your qualifications and interest in the company]

Thank you for considering my application.

Best regards,
[Your Name]`;
        try {
            const info = yield sendEmail(Email, subject, text);
            console.log(`Email sent to ${Email}: ${info.response}`);
        }
        catch (error) {
            console.error(`Failed to send email to ${Email}:`, error);
        }
    }
});
automateEmails();
