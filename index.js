import xlsx from "xlsx";
import nodemailer from "nodemailer";
import express from "express";
import { MongoClient } from "mongodb";
import dotenv from "dotenv";

dotenv.config();

const MONGO_URI = process.env.MONGO_URI;
const DB_NAME = process.env.DB_NAME;

let dbInstance = null;
const connectDB = async () => {
    if (dbInstance) return dbInstance;
    try {
        const client = new MongoClient(MONGO_URI);
        await client.connect();
        console.log("Connected to MongoDB");
        dbInstance = client.db(DB_NAME);
        return dbInstance;
    } catch (error) {
        console.error("MongoDB connection error:", error);
        process.exit(1);
    }
};

let failedMailLogs = [];
let successMailLogs = [];

const workbook = xlsx.readFile("./company_wise_hr.xlsx");
const sheetName = "Sheet1";
const worksheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(worksheet);

const newTransporter = () => {
    return nodemailer.createTransport({
        pool: true,
        host: "smtp.gmail.com",
        port: 465,
        secure: true,
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS,
        },
    });
};
const transporter = newTransporter();

// Dynamic Template Generator
const generateTemplate = (row, templateType) => {
    const { Name, Company, Email, Role, Link } = row;
    const name = Name.split(" ")[0];
    switch (templateType) {
        case "frontendOpportunity":
            return {
                from: `${process.env.EMAIL_USER}`,
                to: Email,
                subject: `Request for Frontend Developer Opportunity at ${Company}`,
                html: `<p>Greetings ${name},</p>
                    <p>I'm writing to express my keen interest in the frontend opportunity at ${Company}. 
                    I have extensive experience in React.js, Next.js, and other modern frontend technologies.</p>
                    <p>Looking forward to your response.</p>
                    <p>Best regards,<br>Your Name</p>`
            };
        case "position":
            return {
                from: `${process.env.EMAIL_USER}`,
                to: Email,
                subject: `Application for ${Role} Role at ${Company}`,
                html: `<p>Dear ${name},</p>
                    <p>I am interested in the ${Role} role at ${Company}. 
                    With my background in frontend development, I am confident in contributing effectively to your team.</p>
                    <p>Find my LinkedIn profile <a href="${Link}">here</a>.</p>
                    <p>Looking forward to your response.</p>
                    <p>Best regards,<br>Your Name</p>`
            };
        default:
            console.error("Invalid template type provided.");
            return null;
    }
};

// Function to send email
const sendEmail = async (row, templateType) => {
    const mailOptions = generateTemplate(row, templateType);
    if (!mailOptions) return;
    try {
        await transporter.sendMail(mailOptions);
        console.log(`Email sent to ${row.Email}`);
        successMailLogs.push(`Email sent to ${row.Email}`);
    } catch (error) {
        console.error(`Error sending email to ${row.Email}:`, error);
        failedMailLogs.push(row);
    }
};

const sendEmailsSynchronously = async (templateType) => {
    const db = await connectDB();
    const collection = db.collection("mailCount");
    let lastCountDoc = await collection.findOne({ _id: "emailCount" });
    let lastCount = lastCountDoc ? lastCountDoc.count : 0;
    let currentCount = lastCount;

    const batchSize = 400; // assuming 400 emails limit for the day
    const dataSubset = data.slice(lastCount, lastCount + batchSize);

    console.log(">>> Starting Mail Chain");
    if (dataSubset.length === 0) {
        console.log("No records to process for today.");
        return;
    }

    for (const row of dataSubset) {
        currentCount += 1;
        await sendEmail(row, templateType);
        await collection.updateOne({ _id: "emailCount" }, { $set: { count: currentCount } }, { upsert: true });
        await new Promise((resolve) => setTimeout(resolve, Math.random() * 90000));
    }
    console.log("All emails sent successfully.");
    console.log(">>> Mails sent:", successMailLogs);
    console.log(">>> Mails failed:", failedMailLogs);
    successMailLogs = [];
    failedMailLogs = [];
};

// Express Routes
const app = express();
const port = 8080;

app.get("/", (req, res) => {
    res.send({message: "Welcome to the Email Automation API", howToUse: 'Just hit the /send route to initiate the email chain'});

});

app.get("/count", async (req, res) => {
    const db = await connectDB();
    const collection = db.collection("mailCount");
    let lastCountDoc = await collection.findOne({ _id: "emailCount" });
    res.json({ mailCount: lastCountDoc.count });
});

app.get("/send", async (req, res) => {
    const templateType = req.query.template || "frontendOpportunity";
    sendEmailsSynchronously(templateType);
    res.json("Emails initiated successfully");
});

app.listen(port, () => {
    console.log(`Server started`);
});
