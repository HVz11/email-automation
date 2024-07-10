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
    const subject =
      "Software Engineering Student Interested in SDE/Full Stack Roles";
    const text = `Dear ${Name},

I hope this message finds you well. My name is Vaibhav Singh, and I am a UG software engineering student at IIIT Bhopal (2024).

I am writing to express my strong interest in joining ${Company} as an SDE/Full Stack Intern and to inquire about any relevant opportunities within your organization.

I am passionate about software development and have a solid foundation in programming languages such as JavaScript, C/C++, Python, SQL, and Typescript as well as experience in Node, Express, NextJs, tailwind, MongoDB, PostgreSQL, Django, and other Python libraries, etc 

Here are a few highlights of my qualifications:
I am proficient in MERN stack and Data Structures and Algorithms and a beginner in Nextjs, and DevOps. I have also been doing competitive programming since last year. I have much experience in developing Full Stack projects which are listed on my Github. 

I am currently doing an internship at KaamBack as a Full Stack Intern. My day-to-day tasks mainly involve working on the backend of the website. I develop the service to match clients with the talent they require using a rating-based system. I am also working on the DevOps part, such as implementing the CI/CD pipeline and Dockerizing the application.

I worked on a research paper titled "Design and Modelling of Machine Learning Based Photonic Sensor for Different Disease Detections", which has been accepted for presentation at the IEEE ICDV 2024 conference. We used pre-tuned machine learning models and Python, along with other libraries, for improved visualization and analysis.


I am particularly drawn to ${Company}'s vision and believe that my skills and enthusiasm align well with your organization's goals. I am eager to contribute my technical skills, collaborate with your team, and continue my growth as a software engineer.

I have attached my resume to provide you with a more detailed overview of my qualifications.

Resume: https://drive.google.com/file/d/1ldF0vvzma6vXnxk2Al9N4IgIFO8SpWR8/view?usp=sharing
Github: https://github.com/HVz11

Thank you for considering my application. I look forward to the possibility of speaking with you about potential opportunities. Please feel free to reach out to me via LinkedIn or email at vs630916@gmail.com  to schedule a conversation at your convenience.

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
