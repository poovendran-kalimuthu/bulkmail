const express = require("express");
const cors = require("cors");
const nodemailer = require("nodemailer");
const mongoose = require("mongoose");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");

const app = express();
app.use(cors({
    origin: process.env.FRONTEND_URL || "*",
    credentials: true
}));
app.use(express.json());

const JWT_SECRET = process.env.JWT_SECRET || "bulkmail_secret_key_123";
const MONGODB_URI = "mongodb+srv://poovendranhari_db_user:6Hbd5DOwL2s0y3OY@cluster0.tr8qgih.mongodb.net/bulkmail?appName=Cluster0";

// Connect to MongoDB Atlas
mongoose.connect(MONGODB_URI)
    .then(() => console.log("Connected to MongoDB Atlas successfully."))
    .catch(err => console.error("MongoDB Connection error:", err.message));

// --- SCHEMAS & MODELS ---

const Schema = mongoose.Schema;

// User Model
const userSchema = new Schema({
    email: { type: String, required: true, unique: true, lowercase: true, trim: true },
    password: { type: String, required: true },
    createdAt: { type: Date, default: Date.now }
});
const User = mongoose.model("User", userSchema);

// Template Model
const templateSchema = new Schema({
    userId: { type: Schema.Types.ObjectId, ref: "User", required: true },
    name: { type: String, required: true },
    subject: { type: String, required: true },
    body: { type: String, required: true },
    isHtml: { type: Boolean, default: true },
    createdAt: { type: Date, default: Date.now }
});
const Template = mongoose.model("Template", templateSchema);

// Campaign History Log Model
const campaignLogSchema = new Schema({
    userId: { type: Schema.Types.ObjectId, ref: "User", required: true },
    date: { type: String, required: true },
    subject: { type: String, required: true },
    total: { type: Number, required: true },
    success: { type: Number, required: true },
    failure: { type: Number, required: true },
    status: { type: String, required: true },
    details: [
        {
            email: String,
            success: Boolean,
            error: String,
            messageId: String
        }
    ],
    createdAt: { type: Date, default: Date.now }
});
const CampaignLog = mongoose.model("CampaignLog", campaignLogSchema);

// SMTP Configurations Model
const smtpSettingsSchema = new Schema({
    userId: { type: Schema.Types.ObjectId, ref: "User", required: true, unique: true },
    useCustom: { type: Boolean, default: false },
    host: { type: String, default: "smtp.gmail.com" },
    port: { type: Number, default: 587 },
    secure: { type: Boolean, default: false },
    user: { type: String, default: "" },
    pass: { type: String, default: "" },
    fromName: { type: String, default: "" },
    fromEmail: { type: String, default: "" },
    createdAt: { type: Date, default: Date.now }
});
const SmtpSettings = mongoose.model("SmtpSettings", smtpSettingsSchema);

// Contact Model
const contactSchema = new Schema({
    userId: { type: Schema.Types.ObjectId, ref: "User", required: true },
    email: { type: String, required: true, trim: true },
    valid: { type: Boolean, required: true },
    placeholders: { type: Schema.Types.Mixed, default: {} },
    createdAt: { type: Date, default: Date.now }
});
const Contact = mongoose.model("Contact", contactSchema);

// --- AUTHENTICATION MIDDLEWARE ---

const authMiddleware = (req, res, next) => {
    const token = req.header("Authorization")?.replace("Bearer ", "");
    if (!token) {
        return res.status(401).json({ success: false, error: "Access denied. Login required." });
    }

    try {
        const decoded = jwt.verify(token, JWT_SECRET);
        req.user = decoded;
        next();
    } catch (ex) {
        res.status(401).json({ success: false, error: "Session expired or invalid token. Please log in again." });
    }
};

// --- ENDPOINTS ---

// Basic health check
app.get("/", (req, res) => {
    res.send("SwiftPost Authenticated Backend is running...");
});

// Detailed health check API
app.get("/api/health", (req, res) => {
    const dbStatus = mongoose.connection.readyState === 1 ? "CONNECTED" : "DISCONNECTED";
    res.json({
        success: true,
        status: "UP",
        timestamp: new Date(),
        uptime: process.uptime(),
        database: dbStatus
    });
});

// 1. Auth: Signup
app.post("/api/auth/signup", async (req, res) => {
    const { email, password } = req.body;
    if (!email || !password) {
        return res.status(400).json({ success: false, error: "Email and password are required." });
    }

    try {
        const existingUser = await User.findOne({ email });
        if (existingUser) {
            return res.status(400).json({ success: false, error: "User with this email already exists." });
        }

        const hashedPassword = await bcrypt.hash(password, 10);
        const user = new User({ email, password: hashedPassword });
        await user.save();

        // Create default SMTP configuration
        const defaultSmtp = new SmtpSettings({ userId: user._id });
        await defaultSmtp.save();

        const token = jwt.sign({ id: user._id, email: user.email }, JWT_SECRET, { expiresIn: "24h" });

        res.status(201).json({
            success: true,
            token,
            user: { id: user._id, email: user.email }
        });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 2. Auth: Login
app.post("/api/auth/login", async (req, res) => {
    const { email, password } = req.body;
    if (!email || !password) {
        return res.status(400).json({ success: false, error: "Email and password are required." });
    }

    try {
        const user = await User.findOne({ email });
        if (!user) {
            return res.status(400).json({ success: false, error: "Invalid email or password." });
        }

        const isMatch = await bcrypt.compare(password, user.password);
        if (!isMatch) {
            return res.status(400).json({ success: false, error: "Invalid email or password." });
        }

        const token = jwt.sign({ id: user._id, email: user.email }, JWT_SECRET, { expiresIn: "24h" });

        res.json({
            success: true,
            token,
            user: { id: user._id, email: user.email }
        });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 3. Auth: Verify current token
app.get("/api/auth/me", authMiddleware, async (req, res) => {
    try {
        const user = await User.findById(req.user.id).select("-password");
        if (!user) {
            return res.status(404).json({ success: false, error: "User not found." });
        }
        res.json({ success: true, user });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 4. Templates: Get templates
app.get("/api/templates", authMiddleware, async (req, res) => {
    try {
        const list = await Template.find({ userId: req.user.id }).sort({ createdAt: -1 });
        res.json({ success: true, data: list });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 5. Templates: Add template
app.post("/api/templates", authMiddleware, async (req, res) => {
    const { name, subject, body, isHtml } = req.body;
    if (!name || !subject || !body) {
        return res.status(400).json({ success: false, error: "Name, subject and body are required." });
    }

    try {
        const tpl = new Template({
            userId: req.user.id,
            name,
            subject,
            body,
            isHtml
        });
        await tpl.save();
        res.status(201).json({ success: true, data: tpl });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 6. Templates: Delete template
app.delete("/api/templates/:id", authMiddleware, async (req, res) => {
    try {
        const result = await Template.findOneAndDelete({ _id: req.params.id, userId: req.user.id });
        if (!result) {
            return res.status(404).json({ success: false, error: "Template not found or unauthorized." });
        }
        res.json({ success: true, message: "Template deleted successfully." });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 7. SMTP Settings: Get saved SMTP config
app.get("/api/smtp", authMiddleware, async (req, res) => {
    try {
        let settings = await SmtpSettings.findOne({ userId: req.user.id });
        if (!settings) {
            settings = new SmtpSettings({ userId: req.user.id });
            await settings.save();
        }
        res.json({ success: true, data: settings });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 8. SMTP Settings: Save config
app.post("/api/smtp", authMiddleware, async (req, res) => {
    const { useCustom, host, port, secure, user, pass, fromName, fromEmail } = req.body;
    try {
        let settings = await SmtpSettings.findOne({ userId: req.user.id });
        if (!settings) {
            settings = new SmtpSettings({ userId: req.user.id });
        }

        settings.useCustom = useCustom;
        settings.host = host || "smtp.gmail.com";
        settings.port = parseInt(port) || 587;
        settings.secure = secure ?? false;
        settings.user = user || "";
        if (pass !== undefined) {
            settings.pass = pass;
        }
        settings.fromName = fromName || "";
        settings.fromEmail = fromEmail || "";

        await settings.save();
        res.json({ success: true, data: settings });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 8.5. Contacts: Get mailing list
app.get("/api/contacts", authMiddleware, async (req, res) => {
    try {
        const list = await Contact.find({ userId: req.user.id }).sort({ createdAt: -1 });
        res.json({ success: true, data: list });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 8.6. Contacts: Add single contact manually
app.post("/api/contacts", authMiddleware, async (req, res) => {
    const { email, valid, placeholders } = req.body;
    if (!email) {
        return res.status(400).json({ success: false, error: "Email is required" });
    }

    try {
        const newContact = new Contact({
            userId: req.user.id,
            email,
            valid,
            placeholders: placeholders || {}
        });
        await newContact.save();
        res.status(201).json({ success: true, data: newContact });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 8.7. Contacts: Batch save contacts (for Excel uploads)
app.post("/api/contacts/batch", authMiddleware, async (req, res) => {
    const { contacts, overwrite } = req.body;
    if (!Array.isArray(contacts)) {
        return res.status(400).json({ success: false, error: "Contacts must be an array" });
    }

    try {
        if (overwrite) {
            await Contact.deleteMany({ userId: req.user.id });
        }

        const contactsToInsert = contacts.map(c => ({
            userId: req.user.id,
            email: c.email,
            valid: c.valid,
            placeholders: c.placeholders || {}
        }));

        const inserted = await Contact.insertMany(contactsToInsert);
        res.status(201).json({ success: true, count: inserted.length });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 8.8. Contacts: Update single contact (inline edit)
app.put("/api/contacts/:id", authMiddleware, async (req, res) => {
    const { email, valid, placeholders } = req.body;
    try {
        const contact = await Contact.findOneAndUpdate(
            { _id: req.params.id, userId: req.user.id },
            { email, valid, placeholders },
            { new: true }
        );
        if (!contact) {
            return res.status(404).json({ success: false, error: "Contact not found or unauthorized" });
        }
        res.json({ success: true, data: contact });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 8.9. Contacts: Delete single contact
app.delete("/api/contacts/:id", authMiddleware, async (req, res) => {
    try {
        const result = await Contact.findOneAndDelete({ _id: req.params.id, userId: req.user.id });
        if (!result) {
            return res.status(404).json({ success: false, error: "Contact not found or unauthorized" });
        }
        res.json({ success: true, message: "Contact deleted successfully." });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 8.10. Contacts: Clear all contacts
app.delete("/api/contacts", authMiddleware, async (req, res) => {
    try {
        await Contact.deleteMany({ userId: req.user.id });
        res.json({ success: true, message: "Mailing list cleared successfully." });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 9. Campaign History: Get logs
app.get("/api/logs", authMiddleware, async (req, res) => {
    try {
        const list = await CampaignLog.find({ userId: req.user.id }).sort({ createdAt: -1 });
        res.json({ success: true, data: list });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 10. Campaign History: Save log
app.post("/api/logs", authMiddleware, async (req, res) => {
    const { date, subject, total, success, failure, status, details } = req.body;
    try {
        const log = new CampaignLog({
            userId: req.user.id,
            date,
            subject,
            total,
            success,
            failure,
            status,
            details
        });
        await log.save();
        res.status(201).json({ success: true, data: log });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 11. Campaign History: Clear all logs
app.delete("/api/logs", authMiddleware, async (req, res) => {
    try {
        await CampaignLog.deleteMany({ userId: req.user.id });
        res.json({ success: true, message: "Campaign history logs cleared successfully." });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// 12. Main Dispatcher: Send Mail
app.post("/sendmail", authMiddleware, async function (req, res) {
    const { msg, sub, email, html, smtpConfig } = req.body;

    if (!email || (Array.isArray(email) && email.length === 0)) {
        return res.status(400).json({ success: false, error: "No recipients specified" });
    }

    const emailList = Array.isArray(email) ? email : [email];

    // Determine final SMTP settings
    let finalHost = "";
    let finalPort = 587;
    let finalSecure = false;
    let finalUser = "";
    let finalPass = "";
    let finalFromName = "";
    let finalFromEmail = "";

    // 1. If explicit config passed in request body (custom test connection etc.)
    if (smtpConfig && smtpConfig.host) {
        finalHost = smtpConfig.host;
        finalPort = parseInt(smtpConfig.port) || 587;
        finalSecure = smtpConfig.secure ?? false;
        finalUser = smtpConfig.user || "";
        finalPass = smtpConfig.pass || "";
        finalFromName = smtpConfig.fromName || "";
        finalFromEmail = smtpConfig.fromEmail || finalUser;
    } else {
        // 2. Load user settings from Database
        const dbSettings = await SmtpSettings.findOne({ userId: req.user.id });
        if (!dbSettings || !dbSettings.user || !dbSettings.pass || !dbSettings.host) {
            return res.status(400).json({ 
                success: false, 
                error: "SMTP configuration is incomplete. Please configure your SMTP server host, username, and password in the SMTP Server tab first." 
            });
        }
        finalHost = dbSettings.host;
        finalPort = dbSettings.port;
        finalSecure = dbSettings.secure;
        finalUser = dbSettings.user;
        finalPass = dbSettings.pass;
        finalFromName = dbSettings.fromName;
        finalFromEmail = dbSettings.fromEmail || finalUser;
    }

    if (!finalHost || !finalUser || !finalPass) {
        return res.status(400).json({ 
            success: false, 
            error: "SMTP server credentials (Host, Username, Password) are missing. Please configure them in the SMTP Server tab." 
        });
    }

    const fromAddress = finalFromName ? `"${finalFromName}" <${finalFromEmail}>` : finalFromEmail;

    try {
        const transporter = nodemailer.createTransport({
            host: finalHost,
            port: finalPort,
            secure: finalSecure,
            auth: {
                user: finalUser,
                pass: finalPass,
            },
            tls: {
                rejectUnauthorized: false
            }
        });

        await transporter.verify();

        const results = [];
        for (let i = 0; i < emailList.length; i++) {
            const recipient = emailList[i];
            let recipientEmail = "";
            let emailSubject = sub;
            let emailText = msg;
            let emailHtml = "";
            if (html === true || html === "true") {
                emailHtml = msg || "";
            } else if (typeof html === "string") {
                emailHtml = html;
            }


            if (typeof recipient === "object" && recipient !== null) {
                recipientEmail = recipient.email;
                const placeholders = recipient.placeholders || {};

                Object.keys(placeholders).forEach(key => {
                    const val = placeholders[key] || "";
                    const regex = new RegExp(`{${key}}`, "g");
                    emailSubject = emailSubject.replace(regex, val);
                    emailText = emailText.replace(regex, val);
                    if (emailHtml) {
                        emailHtml = emailHtml.replace(regex, val);
                    }
                });
            } else {
                recipientEmail = recipient;
            }

            try {
                const mailOptions = {
                    from: fromAddress,
                    to: recipientEmail,
                    subject: emailSubject,
                };

                if (emailHtml) {
                    mailOptions.html = emailHtml;
                    mailOptions.text = emailText || "This email is formatted in HTML. Please view it in an HTML-compatible client.";
                } else {
                    mailOptions.text = emailText;
                }

                const info = await transporter.sendMail(mailOptions);
                results.push({ email: recipientEmail, success: true, messageId: info.messageId });
            } catch (sendError) {
                results.push({ email: recipientEmail, success: false, error: sendError.message });
            }
        }

        const allSuccess = results.every(r => r.success);
        res.json({
            success: allSuccess,
            results,
            compatSuccess: allSuccess
        });

    } catch (connectionError) {
        res.json({
            success: false,
            error: `SMTP connection failed: ${connectionError.message}`,
            results: emailList.map(recipient => {
                const email = typeof recipient === "object" ? recipient.email : recipient;
                return { email, success: false, error: `SMTP Connection failed: ${connectionError.message}` };
            })
        });
    }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, function () {
    console.log(`Server started on port ${PORT}....`);
});