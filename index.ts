import * as nodemailer from "nodemailer";
import * as xlsx from "xlsx";
import * as fs from "fs";

// Sample data to be written to the Excel file
const recruiters = [
  { Name: "John Doe", Email: "exwhyzed884@gmail.com", Company: "Company A" },
];

// Function to create an XLSX file
const createExcelFile = (data: any[], filePath: string) => {
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
} else {
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
const readExcel = (filePath: string) => {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet);
};

// Function to send email
const sendEmail = (
  to: string,
  subject: string,
  text: string
): Promise<nodemailer.SentMessageInfo> => {
  const mailOptions = {
    from: "vs630916@gmail.com",
    to,
    subject,
    text,
  };

  return new Promise((resolve, reject) => {
    transporter.sendMail(
      mailOptions,
      (error, info: nodemailer.SentMessageInfo) => {
        if (error) {
          return reject(error);
        }
        resolve(info);
      }
    );
  });
};

// Main function to automate the process
const automateEmails = async () => {
  if (!fs.existsSync(filePath)) {
    console.error("Excel file not found.");
    return;
  }

  const recruiters = readExcel(filePath) as {
    Email: string;
    Name: string;
    Company: string;
  }[];
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
      const info = await sendEmail(Email, subject, text);
      console.log(`Email sent to ${Email}: ${info.response}`);
    } catch (error) {
      console.error(`Failed to send email to ${Email}:`, error);
    }
  }
};

automateEmails();
