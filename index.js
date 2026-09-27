require('dotenv').config();

const path = require('path');
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const { subject, buildBody } = require('./email-message');

const emailsFile = path.join(__dirname, 'emails-list.xlsx');
const dryRun = process.argv.includes('--dry-run');

function parseRecipients() {
  const workbook = XLSX.readFile(emailsFile);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

  return rows
    .map((row) => {
      const name = String(row.Name || row.name || '').trim();
      const email = String(row.Email || row.email || '').trim();

      if (!email) {
        return null;
      }

      return { name, email };
    })
    .filter(Boolean);
}

function validateConfig() {
  const required = ['SMTP_HOST', 'SMTP_USER', 'SMTP_PASS'];
  const missing = required.filter((key) => !process.env[key]);

  if (missing.length > 0) {
    throw new Error(
      `Missing environment variables: ${missing.join(', ')}. Copy .env.example to .env and fill in your SMTP credentials.`
    );
  }
}

function createTransporter() {
  validateConfig();

  return nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure: Number(process.env.SMTP_PORT || 587) === 465,
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS,
    },
  });
}

async function sendBulkEmails() {
  const recipients = parseRecipients();

  if (recipients.length === 0) {
    throw new Error('No valid rows found in emails-list.xlsx. Make sure the file has Name and Email columns.');
  }

  if (dryRun) {
    console.log(`Dry run enabled. ${recipients.length} email(s) are ready to send.`);
    recipients.forEach(({ name, email }) => {
      console.log(`- ${email}${name ? ` (${name})` : ''}`);
    });
    return;
  }

  const transporter = createTransporter();

  for (const recipient of recipients) {
    const messageBody = buildBody({
      recipientName: recipient.name,
      recipientEmail: recipient.email,
    });

    const mailOptions = {
      from: process.env.SMTP_FROM || process.env.SMTP_USER,
      to: recipient.email,
      replyTo: process.env.SMTP_REPLY_TO || process.env.SMTP_USER,
      subject,
      text: messageBody,
      html: `<pre style="font-family:Arial, sans-serif; white-space:pre-wrap;">${messageBody.replace(/\n/g, '<br>')}</pre>`,
    };

    await transporter.sendMail(mailOptions);
    console.log(`Sent email to ${recipient.email}`);
  }

  console.log(`Finished. ${recipients.length} email(s) sent.`);
}

sendBulkEmails().catch((error) => {
  console.error('Bulk sender failed:', error.message);
  process.exit(1);
});
