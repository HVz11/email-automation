import * as nodemailer from "nodemailer";
import * as xlsx from "xlsx";
import * as fs from "fs";

// Specify the file path for the Excel file
const filePath = "./recruiters.xlsx";

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
    const subject =
      "Software Engineering Student Interested in SDE/Full Stack roles";
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
      const info = await sendEmail(Email, subject, text);
      console.log(`Email sent to ${Email}: ${info.response}`);
    } catch (error) {
      console.error(`Failed to send email to ${Email}:`, error);
    }
  }
};

automateEmails();
