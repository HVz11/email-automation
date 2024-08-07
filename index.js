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
// Specify the file path for the Excel file
const filePath = "./sample.xlsx";
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
        const subject = "Software Engineering Student Interested in SDE/Full Stack roles";
        const text = `Dear ${Name},

I hope this message finds you well. I am Vaibhav Singh, a UG software engineering student at IIIT Bhopal (2024).

I am writing to express my strong interest in joining ${Company} as an SDE/Full Stack Intern and to inquire about any relevant opportunities within your organization.

I am proficient in JavaScript, C/C++, Python, SQL, and Typescript, and I have experience with Node, Express, Next.js, MongoDB, PostgreSQL, Django, and various Python libraries. I possess strong skills in MERN stack, Data Structures, and Algorithms, and I am currently expanding my knowledge in Next.js and DevOps.

Here are a few highlights of my  Internship Experience:
I am currently doing an internship at KaamBack as a Full Stack Intern. My day-to-day tasks mainly involve working on the backend of the website. I develop the service to match clients with the talent they require using a rating-based system. I am also working on the DevOps part, such as implementing the CI/CD pipeline and Dockerizing the application.

I am especially interested in ${Company}'s vision and I believe that my skills and enthusiasm are well-aligned with your organization's goals. I am eager to contribute my technical skills, collaborate with your team, and continue to grow as a software engineer.

I have attached my resume to provide you with a more detailed overview of my qualifications.

Resume: https://drive.google.com/file/d/1ldF0vvzma6vXnxk2Al9N4IgIFO8SpWR8/view?usp=sharing
Github: https://github.com/HVz11

Thank you for considering my application. I look forward to speaking with you about potential opportunities. Please feel free to reach out to me via LinkedIn or email at vs630916@gmail.com  to schedule a conversation at your convenience.

Warm regards,

Vaibhav Singh 
Linkedin: https://www.linkedin.com/in/vaibhav-singh-11vs/
Mobile no.: +91 8010875037
`;
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
