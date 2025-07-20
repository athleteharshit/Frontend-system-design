const fs = require("fs");
const xlsx = require("xlsx");
const nodemailer = require("nodemailer");
const { exit } = require("process");

// Load Excel file
const workbook = xlsx.readFile("./list.xlsx");
const sheetName = "harshit";
const worksheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(worksheet);

// Email configuration
const newTransporter = () => {
  return nodemailer.createTransport({
    pool: true,
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    auth: {
      user: "guptaharshit545@gmail.com",
      pass: process.env.MAIL_KEY,
    },
  });
};
const transporter = newTransporter();

const sendEmail = async (row) => {
  const { Name, Company, Email, Role, Link } = row;
  const nameParts = Name.split(" ");
  const name = nameParts[0];

  const mailOptions = {
    from: "Harshit Gupta <guptaharshit545@gmail.com>",
    to: Email,
    subject: `Request for an Interview Opportunity - ${Role} at ${Company}`,
    html: `
<p>Greetings ${name},</p>

<p>I'm Harshit Gupta, a Frontend Developer | Specialized in Building Scalable React.js Applications. I got to know through your LinkedIn post that <b>${Company}</b> is looking for a <b>${Role}</b>, therefore, I'm reaching out to tell you a bit about myself. I have:</p>

<ul>
  <li><b>4.2 years</b> of hands-on experience in <b>Frontend Development</b></li>
  <li>Skilled in <b>JavaScript, TypeScript, React.js, Next.js, Tailwind CSS, Redux Toolkit, Micro-Frontend architecture, monorepo codebases</b></li>
  <li>Built a custom UI library and integrated real-time features like SSE & Webhooks</li>
  <li>Optimized app performance and implemented scalable architectures across multiple internal projects</li>
  <li>Proficient in <b>Testing (Jest, React Testing Library)</b> and performance tools like Lighthouse, DevTools</li>
</ul>

<p>My notice period is <b>45 days</b>, however, I am open to <b>early joining</b> if required and would be happy to discuss a suitable arrangement.</p>

<p>PS: I have attached my <b><a href="https://docs.google.com/document/d/1RWgPEc2r3iKbSULV_w9XeMoEWGDlS0-i63MiPq2zVjc/edit">Resume</a></b>, <b><a href="https://www.linkedin.com/in/athleteharshit/">LinkedIn</a></b>, ${
      Link ? `<b><a href=${Link}>Job Posting</a></b>` : ""
    } and would be very grateful if you could consider me for an interview opportunity at <b>${Company}</b>.</p>

<p>
Thanking you<br>
Regards,<br>
<b>Harshit Gupta</b><br>
<b>Email:</b> guptaharshit545@gmail.com<br>
<b>Contact:</b> +91 7388544174<br>
<b>Portfolio:</b>https://portofolio-frontend.vercel.app/<br>
<b>GitHub:</b>https://github.com/athleteharshit<br>
<b>LeetCode:</b>https://leetcode.com/u/athleteharshit/<br>
</p>
`,
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log("✅ Email sent to", Email);
  } catch (error) {
    console.error("❌ Error sending email to", Email, error.message);
  }
};

const sendEmailsSynchronously = async () => {
  for (const row of data) {
    await sendEmail(row);
    await new Promise((resolve) => setTimeout(resolve, Math.random() * 90000)); // Random delay (up to 90 seconds)
  }
  console.log("🎉 Done Sending all mails");
  exit();
};

sendEmailsSynchronously();
// <li>Currently working as <b>SDE-I</b> at <b>SS Supply Chain Solutions</b> since April 2023</li>
// <li>2 years of experience as a <b>React Developer</b> at <b>Appinventiv</b></li>
// <li>Delivered production-grade projects including onboarding modules, real-time consultations, e-commerce, and healthcare apps</li>
