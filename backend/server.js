require("dotenv").config();
const express  = require("express");
const mongoose = require("mongoose");
const multer   = require("multer");
const path     = require("path");
const fs       = require("fs");
const cors     = require("cors");
const ExcelJS  = require("exceljs");
const archiver = require("archiver");

const app      = express();
const PORT     = process.env.PORT     || 5000;
const BASE_URL = (process.env.BASE_URL || `http://localhost:${PORT}`).replace(/\/$/, "");
const FRONTEND = path.join(__dirname, "../frontend");
const EXCEL_PATH = path.join(__dirname, "customers.xlsx");
const { google } = require("googleapis");



const SHEET_ID = process.env.GOOGLE_SHEET_ID;

const session      = require("express-session");
const MongoStore = require("connect-mongo")(session);
const bcrypt       = require("bcryptjs");
const User         = require("./models/User");

app.use(cors({

  origin: true,
  credentials: true  
}));


// ==============================27-05-2026====================

// ══════════════════════════════════════════
// SESSION SETUP
// ══════════════════════════════════════════
app.use(session({
  secret: process.env.SESSION_SECRET || "2d1da86182b77da7f41f9b68d9cb92b90aee634fd6ca634a7ced29bb7fcd2af4",
  resave: false,
  saveUninitialized: false,
  store: new MongoStore({
    mongooseConnection: mongoose.connection,
    ttl: 24 * 60 * 60
}),
  cookie: {
    maxAge: 24 * 60 * 60 * 1000,
    httpOnly: true
  }
}));

// ══════════════════════════════════════════
// AUTH MIDDLEWARE
// ══════════════════════════════════════════

// Protect any route — must be logged in
function requireAuth(req, res, next) {
  if (!req.session?.user) {
    // API request → return 401
    if (req.path.startsWith("/api") || req.headers["content-type"]?.includes("application/json")) {
      return res.status(401).json({ error: "Not logged in" });
    }
    // Page request → redirect to login
    return res.redirect("/login");
  }
  next();
}

// Protect route — must be superadmin
function requireSuperAdmin(req, res, next) {
  if (!req.session?.user || req.session.user.role !== "superadmin") {
    if (req.headers["accept"]?.includes("application/json")) {
      return res.status(403).json({ error: "Superadmin access required" });
    }
    return res.redirect("/dashboard");
  }
  next();
}


app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));


// ══════════════════════════════════════════
// UPLOADS ONLY — no auth needed
// ══════════════════════════════════════════
app.use("/uploads", express.static(path.join(__dirname, "uploads")));

// ══════════════════════════════════════════
// PAGE ROUTES — PROTECTED
// ══════════════════════════════════════════
app.get("/",             requireAuth, (_req, res) => res.sendFile(path.join(FRONTEND, "dashboard.html")));
app.get("/dashboard",    requireAuth, (_req, res) => res.sendFile(path.join(FRONTEND, "dashboard.html")));
app.get("/form",         requireAuth, (_req, res) => res.sendFile(path.join(FRONTEND, "pl-bl.html")));
app.get("/edit/:id",     requireAuth, (_req, res) => res.sendFile(path.join(FRONTEND, "pl-bl.html")));
app.get("/admin-manage", requireSuperAdmin, (_req, res) => res.sendFile(path.join(FRONTEND, "admin-manage.html")));

// ══════════════════════════════════════════
// STATIC FILES LAST — css, js, images only
// ══════════════════════════════════════════
app.use(express.static(FRONTEND));





// ══════════════════════════════════════════
// AUTH ROUTES
// ══════════════════════════════════════════

app.get("/login", (req, res) => {
  if (req.session?.user) return res.redirect("/dashboard");
  res.sendFile(path.join(FRONTEND, "login.html"));
});

app.post("/auth/login", async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password)
      return res.status(400).json({ error: "Email and password required" });
    const user = await User.findOne({ email: email.toLowerCase().trim() });
    if (!user)
      return res.status(401).json({ error: "Invalid email or password" });
    if (!user.isActive)
      return res.status(403).json({ error: "Your account has been disabled. Contact superadmin." });
    const isMatch = await user.comparePassword(password);
    if (!isMatch)
      return res.status(401).json({ error: "Invalid email or password" });

    await User.updateOne(
      { _id: user._id }, 
      { lastLogin: new Date() 
      });

    req.session.user = {
      id:    user._id,
      name:  user.name,
      email: user.email,
      role:  user.role
    };
    res.json({ ok: true, role: user.role, name: user.name });
  } catch (err) {
    console.error("Login error:", err.message);
    res.status(500).json({ error: err.message });
  }
});

app.post("/auth/logout", (req, res) => {
  req.session.destroy(() => {
    res.json({ ok: true });
  });
});

app.get("/auth/me", requireAuth, (req, res) => {
  res.json({ user: req.session.user });
});

app.post("/auth/create-admin", requireSuperAdmin, async (req, res) => {
  try {
    const { name, email, password } = req.body;
    if (!name || !email || !password)
      return res.status(400).json({ error: "Name, email and password required" });
    const existing = await User.findOne({ email: email.toLowerCase().trim() });
    if (existing)
      return res.status(400).json({ error: "Email already registered" });
    const admin = new User({
      name, email, password,
      role:          "admin",
      createdBy:     req.session.user.id,
      createdByName: req.session.user.name
    });
    await admin.save();
    res.json({ ok: true, message: `Admin ${name} created successfully` });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/auth/admins", requireSuperAdmin, async (req, res) => {
  try {
    const admins = await User.find({ role: "admin" }, { password: 0 })
      .sort({ createdAt: -1 });
    res.json({ admins });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put("/auth/admin/:id/toggle", requireSuperAdmin, async (req, res) => {
  try {
    const admin = await User.findById(req.params.id);
    if (!admin) return res.status(404).json({ error: "Admin not found" });
    admin.isActive = !admin.isActive;
    await admin.save();
    res.json({ ok: true, isActive: admin.isActive,
      message: `Admin ${admin.isActive ? "enabled" : "disabled"}` });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete("/auth/admin/:id", requireSuperAdmin, async (req, res) => {
  try {
    await User.findByIdAndDelete(req.params.id);
    res.json({ ok: true, message: "Admin deleted" });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put("/auth/admin/:id/reset-password", requireSuperAdmin, async (req, res) => {
  try {
    const { newPassword } = req.body;
    if (!newPassword)
      return res.status(400).json({ error: "New password required" });
    const admin = await User.findById(req.params.id);
    if (!admin)
      return res.status(404).json({ error: "Admin not found" });
    admin.password = newPassword;
    await admin.save();
    res.json({ ok: true, message: "Password reset successfully" });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ⚠️ TEMPORARY — visit once then DELETE this route

// ══════════════════════════════════════════════════════════
// BANKER REQUIREMENT MODEL
// ══════════════════════════════════════════════════════════
const BankerRequirement = mongoose.model("BankerRequirement", new mongoose.Schema({
  customerId:  { type: String, required: true },
  status:      { type: String, enum: ["draft", "submitted"], default: "draft" },
  addedBy:     { type: String, default: "" },
  addedByName: { type: String, default: "" },
  data: {
    employmentType: { type: String, default: "" },
    loanAmount:     { type: String, default: "" },
    loanType:       { type: String, default: "" },
    cibil:          { type: String, default: "" },
    income:         { type: String, default: "" },
    bank:           { type: String, default: "" },
    payslips:       { type: String, default: "3" },
    bankStmt:       { type: String, default: "6" },
    notes:          { type: String, default: "" }
  },
  files: {
    cibilFile:      { type: String, default: "" },
    payslip1:       { type: String, default: "" },
    payslip2:       { type: String, default: "" },
    payslip3:       { type: String, default: "" },
    bankFiles:      { type: [String], default: [] },
    form16:         { type: String, default: "" },
    form26:         { type: String, default: "" },
    itrFiles:       { type: [String], default: [] },
    gstFiles:       { type: String, default: "" },
    offerLetter:    { type: String, default: "" },
    appointmentLetter: { type: String, default: "" },
    relievingLetter:   { type: String, default: "" }
  }
}, { timestamps: true }));

// ── PAGE ROUTE ──
app.get("/banker-requirement", requireAuth, (_req, res) =>
  res.sendFile(path.join(FRONTEND, "banker-requirement.html")));

// ── GET ALL ──
app.get("/api/banker-requirements", requireAuth, async (req, res) => {
  try {
    const records = await BankerRequirement.find().sort({ createdAt: -1 });
    res.json({ records });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── CREATE ──
app.post("/api/banker-requirements", requireAuth, (req, res, next) => {
  uploadFields(req, res, err => { if (err) return res.status(400).json({ error: err.message }); next(); });
}, async (req, res) => {
  try {
    const { customerId, employmentType, loanAmount, loanType,
            cibil, income, bank, payslips, bankStmt, notes } = req.body;

    if (!employmentType || !loanAmount || !loanType)
      return res.status(400).json({ error: "Employment type, loan amount and loan type are required" });

    let cid = (customerId || "").trim();
    if (!cid) cid = await generateCustomerId();

    const uploadedFiles = req.files || {};
    const movedFiles = moveFiles(uploadedFiles, cid);

    const files = {
      cibilFile:      movedFiles.cibilFile || "",
      payslip1:       movedFiles.payslip1 || "",
      payslip2:       movedFiles.payslip2 || "",
      payslip3:       movedFiles.payslip3 || "",
      bankFiles:      movedFiles.bankFiles ? (Array.isArray(movedFiles.bankFiles) ? movedFiles.bankFiles : [movedFiles.bankFiles]) : [],
      form16:         movedFiles.form16File || "",
      form26:         movedFiles.form26File || "",
      itrFiles:       movedFiles.itrFiles ? (Array.isArray(movedFiles.itrFiles) ? movedFiles.itrFiles : [movedFiles.itrFiles]) : [],
      gstFiles:       movedFiles.gstFiles || "",
      offerLetter:    movedFiles.offerLetter || "",
      appointmentLetter: movedFiles.appointmentLetter || "",
      relievingLetter:   movedFiles.relievingLetter || ""
    };

    const record = new BankerRequirement({
      customerId:  cid,
      status:      "submitted",
      addedBy:     req.session.user?.email || "",
      addedByName: req.session.user?.name  || "",
      data: { employmentType, loanAmount, loanType,
              cibil: cibil||"", income: income||"", bank: bank||"",
              payslips: payslips||"3", bankStmt: bankStmt||"6", notes: notes||"" },
      files
    });
    await record.save();
    res.json({ ok: true, record });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── UPDATE ──
app.put("/api/banker-requirements/:id", requireAuth, async (req, res) => {
  try {
    const { employmentType, loanAmount, loanType,
            cibil, income, bank, payslips, bankStmt, notes } = req.body;
    const record = await BankerRequirement.findByIdAndUpdate(
      req.params.id,
      { $set: {
          "data.employmentType": employmentType,
          "data.loanAmount":     loanAmount,
          "data.loanType":       loanType,
          "data.cibil":          cibil    || "",
          "data.income":         income   || "",
          "data.bank":           bank     || "",
          "data.payslips":       payslips || "3",
          "data.bankStmt":       bankStmt || "6",
          "data.notes":          notes    || "",
          status: "submitted"
      }},
      { new: true }
    );
    if (!record) return res.status(404).json({ error: "Not found" });
    res.json({ ok: true, record });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── DELETE ──
app.delete("/api/banker-requirements/:id", requireAuth, async (req, res) => {
  try {
    await BankerRequirement.findByIdAndDelete(req.params.id);
    res.json({ ok: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── GENERATE PDF (no name/email/password) ──
app.get("/api/banker-requirements/:id/pdf", requireAuth, async (req, res) => {
  try {
    const record = await BankerRequirement.findById(req.params.id);
    if (!record) return res.status(404).json({ error: "Not found" });
    const d = record.data || {};

    const loanLabels = {
      "fresh-pl":"Fresh Personal Loan","fresh-bl":"Fresh Business Loan",
      "fresh-hl":"Fresh Home Loan","fresh-lap":"Fresh LAP",
      "fresh-ppl":"Fresh Plot Purchase","fresh-p+c":"Fresh Plot + Construction",
      "renewal":"Renewal"
    };

    const PDFDocument = require("pdfkit");
    const doc = new PDFDocument({ margin: 50, size: "A4" });

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition",
      `inline; filename="banker-${record.customerId}.pdf"`);
    doc.pipe(res);

    // Header bar — blue
    doc.rect(0, 0, 612, 80).fill("#1456a0");
    doc.fillColor("#ffffff").fontSize(22).font("Helvetica-Bold")
       .text("Banker Requirement", 50, 25);
    doc.fontSize(11).font("Helvetica")
       .text(`Customer ID: ${record.customerId}`, 50, 52);

    doc.moveDown(3);
    doc.fillColor("#0a2a55").fontSize(14).font("Helvetica-Bold")
       .text("Requirement Details", 50, 100);
    doc.moveTo(50, 118).lineTo(545, 118).strokeColor("#1456a0").lineWidth(2).stroke();

    const fields = [
      ["Employment Type",    d.employmentType === "salaried" ? "Salaried" : d.employmentType === "selfEmployed" ? "Self Employed" : "—"],
      ["Loan Type",          loanLabels[d.loanType] || d.loanType || "—"],
      ["Loan Amount",        d.loanAmount || "—"],
      ["CIBIL Score",        d.cibil || "—"],
      ["Monthly Income",     d.income || "—"],
      ["Preferred Bank",     d.bank || "—"],
      ["Payslips Required",  d.payslips && d.payslips !== "0" ? d.payslips + " months" : "Not required"],
      ["Bank Statements",    d.bankStmt && d.bankStmt !== "0" ? d.bankStmt + " months" : "Not required"],
    ];

    let y = 130;
    fields.forEach(([label, value]) => {
      doc.fillColor("#666666").fontSize(10).font("Helvetica-Bold").text(label, 50, y);
      doc.fillColor("#1a1a2e").fontSize(11).font("Helvetica").text(value, 200, y);
      y += 28;
      doc.moveTo(50, y - 6).lineTo(545, y - 6).strokeColor("#edf1f5").lineWidth(0.5).stroke();
    });

    if (d.notes) {
      y += 10;
      doc.fillColor("#0a2a55").fontSize(13).font("Helvetica-Bold").text("Additional Notes", 50, y);
      y += 20;
      doc.fillColor("#556274").fontSize(11).font("Helvetica")
         .text(d.notes, 50, y, { width: 495, lineGap: 4 });
    }

    // Footer
    doc.rect(0, 770, 612, 72).fill("#f0f4f8");
    doc.fillColor("#8d99ab").fontSize(9).font("Helvetica")
       .text("Generated by Customer Management System — Confidential",
             50, 782, { align: "center", width: 512 });
    doc.text(`Generated on: ${new Date().toLocaleDateString("en-IN")}`,
             50, 796, { align: "center", width: 512 });

    doc.end();
  } catch (err) {
    console.error("Banker PDF error:", err.message);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ── DOWNLOAD ZIP (requirement PDF + payslips + bank statements from customer record) ──
app.get("/api/banker-requirements/:id/zip", requireAuth, async (req, res) => {
  try {
    const record = await BankerRequirement.findById(req.params.id);
    if (!record) return res.status(404).json({ error: "Not found" });

    const d = record.data || {};
    const loanLabels = {
      "fresh-pl":"Fresh Personal Loan","fresh-bl":"Fresh Business Loan",
      "fresh-hl":"Fresh Home Loan","fresh-lap":"Fresh LAP",
      "fresh-ppl":"Fresh Plot Purchase","fresh-p+c":"Fresh Plot + Construction",
      "renewal":"Renewal"
    };

    // Build the requirement PDF in memory
    const PDFDocument = require("pdfkit");
    const pdfDoc = new PDFDocument({ margin: 50, size: "A4" });
    const pdfChunks = [];
    pdfDoc.on("data", chunk => pdfChunks.push(chunk));

    const pdfReady = new Promise(resolve => pdfDoc.on("end", () => resolve(Buffer.concat(pdfChunks))));

    pdfDoc.rect(0, 0, 612, 80).fill("#1456a0");
    pdfDoc.fillColor("#ffffff").fontSize(22).font("Helvetica-Bold").text("Banker Requirement", 50, 25);
    pdfDoc.fontSize(11).font("Helvetica").text(`Customer ID: ${record.customerId}`, 50, 52);
    pdfDoc.fillColor("#0a2a55").fontSize(14).font("Helvetica-Bold").text("Requirement Details", 50, 100);
    pdfDoc.moveTo(50, 118).lineTo(545, 118).strokeColor("#1456a0").lineWidth(2).stroke();

    const fields = [
      ["Employment Type",   d.employmentType === "salaried" ? "Salaried" : "Self Employed"],
      ["Loan Type",         loanLabels[d.loanType] || d.loanType || "—"],
      ["Loan Amount",       d.loanAmount || "—"],
      ["CIBIL Score",       d.cibil || "—"],
      ["Monthly Income",    d.income || "—"],
      ["Preferred Bank",    d.bank || "—"],
      ["Payslips Required", d.payslips && d.payslips !== "0" ? d.payslips + " months" : "Not required"],
      ["Bank Statements",   d.bankStmt && d.bankStmt !== "0" ? d.bankStmt + " months" : "Not required"],
    ];
    let y = 130;
    fields.forEach(([label, value]) => {
      pdfDoc.fillColor("#666666").fontSize(10).font("Helvetica-Bold").text(label, 50, y);
      pdfDoc.fillColor("#1a1a2e").fontSize(11).font("Helvetica").text(value, 200, y);
      y += 28;
      pdfDoc.moveTo(50, y-6).lineTo(545, y-6).strokeColor("#edf1f5").lineWidth(0.5).stroke();
    });
    if (d.notes) {
      y += 10;
      pdfDoc.fillColor("#0a2a55").fontSize(13).font("Helvetica-Bold").text("Notes", 50, y);
      y += 20;
      pdfDoc.fillColor("#556274").fontSize(11).font("Helvetica").text(d.notes, 50, y, { width: 495 });
    }
    pdfDoc.rect(0, 770, 612, 72).fill("#f0f4f8");
    pdfDoc.fillColor("#8d99ab").fontSize(9).font("Helvetica")
      .text("Generated by Customer Management System — Confidential", 50, 782, { align: "center", width: 512 });
    pdfDoc.end();

    const pdfBuffer = await pdfReady;

    // Set up ZIP stream
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition",
      `attachment; filename="banker-${record.customerId}.zip"`);

    const archive = archiver("zip", { zlib: { level: 9 } });
    archive.on("error", err => { throw err; });
    archive.pipe(res);

    // 1. Add the requirement PDF
    archive.append(pdfBuffer, { name: `${record.customerId}/banker-requirement.pdf` });

    // 2. Pull customer files if the customerId matches a customer record
    const customer = await Customer.findOne({ customerId: record.customerId });
    if (customer) {
      const rawFiles = customer.files?.toObject ? customer.files.toObject() : customer.files;
      const uploadsBase = path.join(__dirname, "uploads");

      // Payslips (offerLetter + salary-related pdfs stored under customer uploads)
      // Bank statements
      const fileGroups = [
        { files: rawFiles.bankFiles,  folder: "bank-statements" },
        { files: rawFiles.itrFiles,   folder: "itr-files" },
      ];
      // Single files that are relevant for banker
      const singleFiles = [
        { file: rawFiles.cibilFile,   name: "cibil-report" },
        { file: rawFiles.offerLetter, name: "offer-letter" },
        { file: rawFiles.form16File,  name: "form16" },
        { file: rawFiles.form26File,  name: "form26" },
      ];

      fileGroups.forEach(({ files, folder }) => {
        const arr = Array.isArray(files) ? files : (files ? [files] : []);
        arr.filter(Boolean).forEach((f, i) => {
          const fullPath = path.join(uploadsBase, f);
          if (fs.existsSync(fullPath)) {
            const ext = path.extname(f);
            archive.file(fullPath, { name: `${record.customerId}/${folder}/${folder}-${i+1}${ext}` });
          }
        });
      });

      singleFiles.forEach(({ file, name }) => {
        if (file) {
          const fullPath = path.join(uploadsBase, file);
          if (fs.existsSync(fullPath)) {
            const ext = path.extname(file);
            archive.file(fullPath, { name: `${record.customerId}/documents/${name}${ext}` });
          }
        }
      });

      // Add full customer details JSON (no sensitive data masking — admin only download)
      archive.append(
        JSON.stringify(customer.latestData, null, 2),
        { name: `${record.customerId}/customer-details.json` }
      );
    }

    await archive.finalize();
  } catch (err) {
    console.error("Banker ZIP error:", err.message);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});





// =====================================================

// Easy logout via browser URL
app.get("/logout", (req, res) => {
  req.session.destroy(() => {
    res.redirect("/login");
  });
});





// ================================27-05-2026====================

mongoose
  .connect(process.env.MONGO_URI)
  .then(() => {
    console.log(" MongoDB connected");
    app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
  })
  .catch(err => { console.error(" MongoDB error:", err.message); process.exit(1); });

const DailyCounter = mongoose.model("DailyCounter", new mongoose.Schema({
  dateKey: { type: String, unique: true },
  count:   { type: Number, default: 0 }
}));

const Customer = mongoose.model("Customer", new mongoose.Schema({
  customerId:   { type: String, required: true, unique: true },

   // ← ADD THESE TWO LINES
  addedBy:      { type: String, default: "" },      // admin email
  addedByName:  { type: String, default: "" },      // admin name


  email:        { type: String, default: "" },
  phone:        { type: String, default: "" },
  version:      { type: Number, default: 1 },
  latestData:   { type: mongoose.Schema.Types.Mixed, default: {} },
  files: {
    summaryPdf:        { type: String, default: "" },
    cibilFile:         { type: String, default: "" },
    offerLetter:       { type: String, default: "" },
    appointmentLetter: { type: String, default: "" },
    relievingLetter:   { type: String, default: "" },
    form16File:        { type: String, default: "" },
    form26File:        { type: String, default: "" },
    itrFiles:          { type: [String], default: [] },
    bankFiles:         { type: [String], default: [] },
    gstFiles:          { type: String, default: "" },
    labourFiles:       { type: String, default: "" }
  },
  history: [{
    version: Number,
    data:    mongoose.Schema.Types.Mixed,
    files:   mongoose.Schema.Types.Mixed,
    savedAt: { type: Date, default: Date.now }
  }]
}, { timestamps: true }));

async function generateCustomerId() {
  const now = new Date();

  // Convert to IST manually
  const ist = new Date(now.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));

  const day = String(ist.getDate()).padStart(2, "0");
  const months = ["JAN","FEB","MAR","APR","MAY","JUN",
                  "JUL","AUG","SEP","OCT","NOV","DEC"];
  const month = months[ist.getMonth()];
  const year  = String(ist.getFullYear()).slice(-2);

  const dateKey = `${day}${month}${year}`;

  const counter = await DailyCounter.findOneAndUpdate(
    { dateKey },
    { $inc: { count: 1 } },
    { upsert: true, new: true }
  );

  return `CUST-${dateKey}-${1000 + counter.count}`;
}

function buildEditLink(id) { return `${BASE_URL}/edit/${id}`; }

function ensureCustomerDir(id) {
  const dir = path.join(__dirname, "uploads", id);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function getDraftId(phone, email) {
  const key = (phone || email || "unknown").replace(/[^a-zA-Z0-9]/g, "");
  return `DRAFT-${key}`;
}

async function deleteDraftRow(provisionalId) {
  if (!fs.existsSync(EXCEL_PATH)) return;
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(EXCEL_PATH);
    const sh = wb.getWorksheet("Customers");
    if (!sh) return;
    let draftRow = -1;
    sh.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      if (String(row.getCell(1).value || "").trim() === provisionalId) draftRow = rowNum;
    });
    if (draftRow > 0) {
      sh.spliceRows(draftRow, 1);
      const tmpPath = EXCEL_PATH + ".tmp";
      await wb.xlsx.writeFile(tmpPath);
      fs.renameSync(tmpPath, EXCEL_PATH);
    }
  } catch (e) { console.error("Draft row cleanup error:", e.message); }
}


//new google sheet code will be added here in future for backup of data to google sheet. For now excel is used for backup and record keeping.
async function getGoogleSheet() {
  const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, "google-credentials.json"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  const client = await auth.getClient();
  return google.sheets({ version: "v4", auth: client });
}

async function ensureSheetHeader(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "Sheet1!A1:BZ1"
  });
  const existing = res.data.values?.[0] || [];
  if (existing.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: "Sheet1!A1",
      valueInputOption: "RAW",
      requestBody: { values: [EXCEL_COLUMNS.map(c => c.header)] }
    });

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 0, endRowIndex: 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.04, green: 0.165, blue: 0.333 },
                  textFormat: {
                    bold: true,
                    foregroundColor: { red: 1, green: 1, blue: 1 },
                    fontSize: 11
                  },
                  horizontalAlignment: "CENTER",
                  verticalAlignment: "MIDDLE"
                }
              },
              fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
          },
          {
            updateSheetProperties: {
              properties: { sheetId: 0, gridProperties: { frozenRowCount: 1 } },
              fields: "gridProperties.frozenRowCount"
            }
          }
        ]
      }
    });
  }
}

async function updateGoogleSheet(customerId, formData, files, version) {
  try {
    const sheets = await getGoogleSheet();
    await ensureSheetHeader(sheets);
    const rowData = {
      customerId, ...formData,
      cibilFile:         files?.cibilFile         ? "Yes" : "No",
      offerLetter:       files?.offerLetter        ? "Yes" : "No",
      appointmentLetter: files?.appointmentLetter  ? "Yes" : "No",
      relievingLetter:   files?.relievingLetter    ? "Yes" : "No",
      form16File:        files?.form16File         ? "Yes" : "No",
      form26File:        files?.form26File         ? "Yes" : "No",
      itrFiles:          (files?.itrFiles?.length) ? "Yes" : "No",
      bankFiles:         (files?.bankFiles?.length) ? "Yes" : "No",
      gstFiles:          files?.gstFiles           ? "Yes" : "No",
      labourFiles:       files?.labourFiles        ? "Yes" : "No",
      summaryPdf:        files?.summaryPdf         ? "Yes" : "No",
      version,
      updatedAt: new Date().toLocaleString("en-IN"),
      editLink:  buildEditLink(customerId)
    };
    const newRow = EXCEL_COLUMNS.map(col => rowData[col.key] ?? "");

    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Sheet1!A:A"
    });
    const rows = existing.data.values || [];
    let foundRowIndex = -1;
    rows.forEach((row, i) => {
      if (i === 0) return;
      if (String(row[0] || "").trim() === customerId) foundRowIndex = i + 1;
    });

    if (foundRowIndex > 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `Sheet1!A${foundRowIndex}`,
        valueInputOption: "RAW",
        requestBody: { values: [newRow] }
      });
    } else {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: "Sheet1!A:A",
        valueInputOption: "RAW",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values: [newRow] }
      });
    }
    console.log(` Google Sheet updated: ${customerId}`);
  } catch (err) {
    console.error(" Google Sheet error:", err.message);
  }
}
// ==================================Google Sheet cleanup for draft rows when lead is submitted.
// This will ensure that there are no stale draft rows in google sheet when lead is submitted.
async function deleteDraftFromGoogleSheet(provisionalId) {
  try {
    const sheets = await getGoogleSheet();
    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Sheet1!A:A"
    });
    const rows = existing.data.values || [];
    let foundRowIndex = -1;
    rows.forEach((row, i) => {
      if (i === 0) return;
      if (String(row[0] || "").trim() === provisionalId) foundRowIndex = i + 1;
    });
    if (foundRowIndex > 0) {
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SHEET_ID,
        range: `Sheet1!A${foundRowIndex}:BZ${foundRowIndex}`
      });
    }
  } catch (e) { console.error("Google Sheet draft cleanup error:", e.message); }
}
//new google sheet code will be added here in future for backup of data to google sheet. For now excel is used for backup and record keeping.



const EXCEL_COLUMNS = [
  { header: "Customer ID",         key: "customerId" },
  { header: "Name",                key: "applicantName" },
  { header: "Phone",               key: "applicantNumber" },
  { header: "Email",               key: "applicantEmail" },
  { header: "DOB",                 key: "dob" },
  { header: "Residential Address", key: "residentialAddress" },
  { header: "Loan Type",           key: "loanType" },
  { header: "Loan Required",       key: "loanRequired" },
  { header: "Bank Credit",         key: "bankCredit" },
  { header: "Employment Type",     key: "employmentType" },
  { header: "Company Name",        key: "companyName" },
  { header: "Designation",         key: "designation" },
  { header: "Company ID Card",     key: "companyIdCard" },
  { header: "Total Experience",    key: "totalExperience" },
  { header: "Present Co. Exp",     key: "presentCompanyExp" },
  { header: "CIBIL Score",         key: "cibil" },
  { header: "Gross Salary",        key: "grossSalary" },
  { header: "Annual Salary",       key: "annualSalary" },
  { header: "Current CTC",         key: "currentCTC" },
  { header: "Latest Net Salary",   key: "latestSalary" },
  { header: "Last Net Salary",     key: "lastSalary" },
  { header: "Before Last Salary",  key: "beforeLastSalary" },
  { header: "Business Name",       key: "businessName" },
  { header: "Office Phone",        key: "officePhone" },
  { header: "Business Type",       key: "businessType" },
  { header: "Monthly Income",      key: "monthlyIncome" },
  { header: "Business Email",      key: "businessEmail" },
  { header: "Business Phone",      key: "businessPhone" },
  { header: "Rental Income",       key: "rentalIncome" },
  { header: "Business Years",      key: "businessYears" },
  { header: "Existing Bank",       key: "existingBank" },
  { header: "Outstanding Amount",  key: "outstandingAmount" },
  { header: "Current EMI",         key: "currentEmi" },
  { header: "Top Up Amount",       key: "topupAmount" },
  { header: "Interest Rate",       key: "interestRate" },
  { header: "Loan Since",          key: "loanSince" },
  { header: "Property Type",       key: "propertyType" },
  { header: "Property Location",   key: "propertyLocation" },
  { header: "Authority Type",      key: "authorityType" },
  { header: "Budget",              key: "budget" },
  { header: "Co-Applicant",        key: "coApplicant" },
  { header: "Co-Applicant Name",   key: "coName" },
  { header: "Co-Applicant Number", key: "coNumber" },
  { header: "Co-Applicant Email",  key: "coEmail" },
  { header: "Co-Applicant Salary", key: "coSalary" },
  { header: "CIBIL File",          key: "cibilFile" },
  { header: "Offer Letter",        key: "offerLetter" },
  { header: "Appointment Letter",  key: "appointmentLetter" },
  { header: "Relieving Letter",    key: "relievingLetter" },
  { header: "Form 16",             key: "form16File" },
  { header: "Form 26AS",           key: "form26File" },
  { header: "ITR Files",           key: "itrFiles" },
  { header: "Bank Statements",     key: "bankFiles" },
  { header: "GST Documents",       key: "gstFiles" },
  { header: "Labour Licence",      key: "labourFiles" },
  { header: "Summary PDF",         key: "summaryPdf" },
  { header: "Version",             key: "version" },
  { header: "Last Updated",        key: "updatedAt" },
  { header: "Edit Link",           key: "editLink" }
];

let excelWriting = false;
const excelQueue = [];

async function updateExcel(customerId, formData, files, version) {
  return new Promise((resolve, reject) => {
    excelQueue.push({ customerId, formData, files, version, resolve, reject });
    if (!excelWriting) processExcelQueue();
  });
}

async function processExcelQueue() {
  if (excelQueue.length === 0) { excelWriting = false; return; }
  excelWriting = true;
  const { customerId, formData, files, version, resolve, reject } = excelQueue.shift();
  try {
    const workbook = new ExcelJS.Workbook();
    if (fs.existsSync(EXCEL_PATH)) await workbook.xlsx.readFile(EXCEL_PATH);
    let sheet = workbook.getWorksheet("Customers");
    if (!sheet) {
      sheet = workbook.addWorksheet("Customers");
      sheet.columns = EXCEL_COLUMNS;
      const hRow = sheet.getRow(1);
      hRow.font      = { bold: true, color: { argb: "FFFFFFFF" } };
      hRow.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0A2A55" } };
      hRow.alignment = { horizontal: "center", vertical: "middle" };
      hRow.height    = 20;
      hRow.commit();
    }
    const rowData = {
      customerId, ...formData,
      cibilFile:         files?.cibilFile         ? "Yes" : "No",
      offerLetter:       files?.offerLetter        ? "Yes" : "No",
      appointmentLetter: files?.appointmentLetter  ? "Yes" : "No",
      relievingLetter:   files?.relievingLetter    ? "Yes" : "No",
      form16File:        files?.form16File         ? "Yes" : "No",
      form26File:        files?.form26File         ? "Yes" : "No",
      itrFiles:          (files?.itrFiles?.length) ? "Yes" : "No",
      bankFiles:         (files?.bankFiles?.length) ? "Yes" : "No",
      gstFiles:          files?.gstFiles           ? "Yes" : "No",
      labourFiles:       files?.labourFiles        ? "Yes" : "No",
      summaryPdf:        files?.summaryPdf         ? "Yes" : "No",
      version,
      updatedAt: new Date().toLocaleString("en-IN"),
      editLink:  buildEditLink(customerId)
    };
    let foundRowNum = -1;
    sheet.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      if (String(row.getCell(1).value || "").trim() === customerId) foundRowNum = rowNum;
    });
    if (foundRowNum > 0) {
      const row = sheet.getRow(foundRowNum);
      EXCEL_COLUMNS.forEach((col, idx) => { row.getCell(idx + 1).value = rowData[col.key] ?? ""; });
      row.commit();
    } else {
      const values = EXCEL_COLUMNS.map(col => rowData[col.key] ?? "");
      const newRow = sheet.addRow(values);
      if (sheet.rowCount % 2 === 0) {
        newRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0F4F8" } };
      }
      newRow.commit();
    }
    const tempPath = EXCEL_PATH + ".tmp";
    await workbook.xlsx.writeFile(tempPath);
    fs.renameSync(tempPath, EXCEL_PATH);
    resolve();
  } catch (err) {
    console.error("Excel write error:", err.message);
    reject(err);
  } finally {
    excelWriting = false;  // ← ADDED: always reset the flag
    processExcelQueue();   // ← already existed, now always runs
  }
}
const storage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    const tempDir = path.join(__dirname, "uploads", "temp");
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });
    cb(null, tempDir);
  },
  filename: (_req, file, cb) => {
    const safe = file.originalname.replace(/[^a-zA-Z0-9_.\-]/g, "_");
    cb(null, Date.now() + "_" + safe);
  }
});
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } });
const uploadFields = upload.fields([
  { name: "summaryPdf", maxCount: 1 }, { name: "cibilFile", maxCount: 1 },
  { name: "payslip1", maxCount: 1 }, { name: "payslip2", maxCount: 1 }, { name: "payslip3", maxCount: 1 },
  { name: "offerLetter", maxCount: 1 }, { name: "appointmentLetter", maxCount: 1 },
  { name: "relievingLetter", maxCount: 1 }, { name: "form16File", maxCount: 1 },
  { name: "form26File", maxCount: 1 }, { name: "itrFiles", maxCount: 5 },
  { name: "bankFiles", maxCount: 6 }, { name: "gstFiles", maxCount: 1 },
  { name: "labourFiles", maxCount: 1 }
]);

function moveFiles(uploadedFiles, customerId) {
  const customerDir = ensureCustomerDir(customerId);
  const result = {};
  Object.entries(uploadedFiles).forEach(([field, files]) => {
    if (!files || files.length === 0) return;
    const moved = files.map(file => {
      const newName = file.originalname.replace(/[^a-zA-Z0-9_.\-]/g, "_");
      const newPath = path.join(customerDir, newName);
      try { fs.renameSync(file.path, newPath); }
      catch (e) { fs.copyFileSync(file.path, newPath); try { fs.unlinkSync(file.path); } catch {} }
      return `${customerId}/${newName}`;
    });
    result[field] = moved.length === 1 ? moved[0] : moved;
  });
  return result;
}

app.get("/generate-id", async (req, res) => {
  try {
    res.json({ customerId: await generateCustomerId() });
  } catch (err) { res.status(500).json({ error: "Could not generate ID" }); }
});

app.post("/auto-save", async (req, res) => {

   console.log("Auto-save hit:", req.body?.email, req.body?.phone);
  try {
    const { email, phone, formData: rawData } = req.body;
    if (!email && !phone) return res.json({ ok: true });

    // let formData = {};
    // try { formData = typeof rawData === "string" ? JSON.parse(rawData) : (rawData || {}); }
    // catch { return res.json({ ok: true }); }

    const formData = rawData && typeof rawData === "object" ? rawData : {}; // ← CHANGED

    let customer = null;
    if (email) customer = await Customer.findOne({ email: email.toLowerCase().trim() });
    if (!customer && phone) customer = await Customer.findOne({ phone: phone.trim() });

    if (customer) {
      customer.latestData = { ...customer.latestData, ...formData };
      if (email) customer.email = email.toLowerCase().trim();
      if (phone) customer.phone = phone.trim();
      await customer.save();
      const rawFiles = customer.files?.toObject ? customer.files.toObject() : customer.files;

  //1.    
      updateExcel(customer.customerId, customer.latestData, rawFiles, customer.version)
        .catch(e => console.error("Auto-save Excel error:", e.message));

        updateGoogleSheet(customer.customerId, customer.latestData, rawFiles, customer.version); // ← ADDED new Google sheet code 

      return res.json({ ok: true, customerId: customer.customerId, isExisting: true });
    }

    // New customer — stable draft ID so same row updates every blur
    const provisionalId = getDraftId(phone, email);
    updateExcel(provisionalId, formData, {}, 0)
      .catch(e => console.error("Draft Excel error:", e.message));

      updateGoogleSheet(provisionalId, formData, {}, 0); // ← ADDED new Google sheet code

    res.json({ ok: true, isExisting: false, provisionalId });

  } catch (err) {
    console.error("Auto-save error:", err.message);
    res.json({ ok: true });
  }
});

app.post("/check-customer", async (req, res) => {
  try {
    const { email, phone } = req.body;
    let customer = null;
    if (email) customer = await Customer.findOne({ email: email.toLowerCase().trim() });
    if (!customer && phone) customer = await Customer.findOne({ phone: phone.trim() });
    if (customer) return res.json({ exists: true, customerId: customer.customerId, version: customer.version, latestData: customer.latestData, files: customer.files });
    res.json({ exists: false });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post("/submit-lead", (req, res, next) => {
  uploadFields(req, res, err => { if (err) return res.status(400).json({ error: err.message }); next(); });
}, async (req, res) => {
  try {
    let { customerId, data: rawData } = req.body;
    let formData = {};
    try { formData = JSON.parse(rawData || "{}"); }
    catch { return res.status(400).json({ error: "Bad form data JSON" }); }

    const email = (formData.applicantEmail || "").toLowerCase().trim();
    const phone = (formData.applicantNumber || "").trim();

    let existing = null;
    if (customerId) existing = await Customer.findOne({ customerId });
    if (!existing && email) existing = await Customer.findOne({ email });
    if (!existing && phone) existing = await Customer.findOne({ phone });

    const uploadedFiles = req.files || {};

    if (existing) {
      const cid = existing.customerId;
      const movedFiles = moveFiles(uploadedFiles, cid);
      const rawExisting = existing.files?.toObject ? existing.files.toObject() : existing.files;
      const newFiles = {
        summaryPdf:        movedFiles.summaryPdf        || rawExisting.summaryPdf        || "",
        cibilFile:         movedFiles.cibilFile         || rawExisting.cibilFile         || "",
        offerLetter:       movedFiles.offerLetter       || rawExisting.offerLetter       || "",
        appointmentLetter: movedFiles.appointmentLetter || rawExisting.appointmentLetter || "",
        relievingLetter:   movedFiles.relievingLetter   || rawExisting.relievingLetter   || "",
        form16File:        movedFiles.form16File        || rawExisting.form16File        || "",
        form26File:        movedFiles.form26File        || rawExisting.form26File        || "",
        itrFiles:  movedFiles.itrFiles  ? (Array.isArray(movedFiles.itrFiles)  ? movedFiles.itrFiles  : [movedFiles.itrFiles])  : (rawExisting.itrFiles  || []),
        bankFiles: movedFiles.bankFiles ? (Array.isArray(movedFiles.bankFiles) ? movedFiles.bankFiles : [movedFiles.bankFiles]) : (rawExisting.bankFiles || []),
        gstFiles:          movedFiles.gstFiles          || rawExisting.gstFiles          || "",
        labourFiles:       movedFiles.labourFiles       || rawExisting.labourFiles       || ""
      };
      existing.history.push({ version: existing.version, data: existing.latestData, files: rawExisting, savedAt: new Date() });
      existing.version   += 1;
      existing.latestData = formData;
      existing.files      = newFiles;
      if (email) existing.email = email;
      if (phone) existing.phone = phone;
      await existing.save();
//3.
      await updateExcel(cid, formData, newFiles, existing.version);
      await updateGoogleSheet(cid, formData, newFiles, existing.version); // ← ADDED new Google sheet code
      await deleteDraftRow(getDraftId(phone, email));
      await deleteDraftFromGoogleSheet(getDraftId(phone, email)).catch(() => {});
      return res.json({ customerId: cid, version: existing.version, hyperlink: buildEditLink(cid), message: "Updated" });
    }

    //const newId = customerId || await generateCustomerId(formData.applicantName || "");

    const newId = customerId || await generateCustomerId();

    ensureCustomerDir(newId);
    const movedFiles = moveFiles(uploadedFiles, newId);
    const newFiles = {
      summaryPdf:        movedFiles.summaryPdf        || "",
      cibilFile:         movedFiles.cibilFile         || "",
      offerLetter:       movedFiles.offerLetter       || "",
      appointmentLetter: movedFiles.appointmentLetter || "",
      relievingLetter:   movedFiles.relievingLetter   || "",
      form16File:        movedFiles.form16File        || "",
      form26File:        movedFiles.form26File        || "",
      itrFiles:  movedFiles.itrFiles  ? (Array.isArray(movedFiles.itrFiles)  ? movedFiles.itrFiles  : [movedFiles.itrFiles])  : [],
      bankFiles: movedFiles.bankFiles ? (Array.isArray(movedFiles.bankFiles) ? movedFiles.bankFiles : [movedFiles.bankFiles]) : [],
      gstFiles:          movedFiles.gstFiles          || "",
      labourFiles:       movedFiles.labourFiles       || ""
    };
    const customer = new Customer({ customerId: newId, email, phone, version: 1, latestData: formData, files: newFiles });
    await customer.save();

    // Auto-create empty PropertyRequirement and BankerRequirement records
    const userEmail = req.session.user?.email || "";
    const userName = req.session.user?.name || "";
    
    await PropertyRequirement.create({
      customerId: newId,
      status: "draft",
      addedBy: userEmail,
      addedByName: userName,
      data: {}
    });
    
    await BankerRequirement.create({
      customerId: newId,
      status: "draft",
      addedBy: userEmail,
      addedByName: userName,
      data: {}
    });

 //3.   
    await updateExcel(newId, formData, newFiles, 1);
    await updateGoogleSheet(newId, formData, newFiles, 1); // ← ADDED new Google sheet code
    await deleteDraftRow(getDraftId(phone, email));
    await deleteDraftFromGoogleSheet(getDraftId(phone, email)).catch(() => {});


    
    // Delete draft from Google Sheet too
   updateGoogleSheet(getDraftId(phone, email) + "_DELETE", {}, {}, 0)
  .catch(() => {}); // optional — or build a deleteGoogleSheetRow function

    return res.json({ customerId: newId, version: 1, hyperlink: buildEditLink(newId), message: "Created" });

  } catch (err) {
    console.error("Submit error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/customer/:id", async (req, res) => {
  try {
    const c = await Customer.findOne({ customerId: req.params.id });
    if (!c) return res.status(404).json({ error: "Not found" });
    const rawFiles = c.files?.toObject ? c.files.toObject() : c.files;
    const fileUrls = {};
    Object.entries(rawFiles).forEach(([key, val]) => {
      fileUrls[key] = Array.isArray(val)
        ? val.filter(Boolean).map(v => `${BASE_URL}/uploads/${v}`)
        : (val ? `${BASE_URL}/uploads/${val}` : "");
    });
    res.json({ customerId: c.customerId, version: c.version, latestData: c.latestData, files: rawFiles, fileUrls, updatedAt: c.updatedAt, createdAt: c.createdAt });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/customers", async (_req, res) => {
  try {
    const list = await Customer.find({}, {
      customerId: 1, version: 1, updatedAt: 1, createdAt: 1,
      "latestData.applicantName": 1, "latestData.applicantNumber": 1,
      "latestData.applicantEmail": 1, "latestData.loanType": 1, "latestData.loanRequired": 1
    }).sort({ updatedAt: -1 });
    res.json({ total: list.length, customers: list });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.delete("/customer/:id", async (req, res) => {
  try {
    const c = await Customer.findOne({ customerId: req.params.id });
    if (!c) return res.status(404).json({ error: "Not found" });
    const dir = path.join(__dirname, "uploads", req.params.id);
    if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
    await Customer.deleteOne({ customerId: req.params.id });
    if (fs.existsSync(EXCEL_PATH)) {
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(EXCEL_PATH);
        const sh = wb.getWorksheet("Customers");
        if (sh) {
          let rowToDel = -1;
          sh.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            if (String(row.getCell(1).value || "").trim() === req.params.id) rowToDel = rowNum;
          });
          if (rowToDel > 0) {
            sh.spliceRows(rowToDel, 1);
            const tmpPath = EXCEL_PATH + ".tmp";
            await wb.xlsx.writeFile(tmpPath);
            fs.renameSync(tmpPath, EXCEL_PATH);
          }
        }
      } catch (e) { console.error("Excel delete row error:", e.message); }
    }
    res.json({ message: `${req.params.id} deleted` });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/download/:id", async (req, res) => {
  try {
    const c = await Customer.findOne({ customerId: req.params.id });
    if (!c) return res.status(404).json({ error: "Not found" });
    const customerDir = path.join(__dirname, "uploads", req.params.id);
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${req.params.id}.zip"`);
    const archive = archiver("zip", { zlib: { level: 9 } });
    archive.on("error", err => { throw err; });
    archive.pipe(res);
    if (fs.existsSync(customerDir)) archive.directory(customerDir, req.params.id);
    archive.append(JSON.stringify(c.latestData, null, 2), { name: `${req.params.id}/customer_details.json` });
    await archive.finalize();
  } catch (err) {
    console.error("Download error:", err.message);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

app.get("/download-excel", (_req, res) => {
  if (!fs.existsSync(EXCEL_PATH)) return res.status(404).json({ error: "Excel not found" });
  res.download(EXCEL_PATH, "customers.xlsx");
});

// Redirect to the Google Sheet online view (preserves live sync). Fall back to Excel if not configured.
app.get("/download-googlesheet", async (_req, res) => {
  console.log("SHEET_ID value:", JSON.stringify(SHEET_ID)); // ← ADD THIS
  try {
    if (SHEET_ID) {
      const viewUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`;
      return res.redirect(viewUrl);
    }
    if (fs.existsSync(EXCEL_PATH)) return res.download(EXCEL_PATH, "customers.xlsx");
    res.status(404).json({ error: "No sheet or excel available" });
  } catch (err) {
    console.error("Download Google Sheet error:", err.message);
    if (fs.existsSync(EXCEL_PATH)) return res.download(EXCEL_PATH, "customers.xlsx");
    res.status(500).json({ error: err.message });
  }
});


// ══════════════════════════════════════════════════════════
// PROPERTY REQUIREMENT MODEL
// ══════════════════════════════════════════════════════════
const PropertyRequirement = mongoose.model("PropertyRequirement", new mongoose.Schema({
  customerId:  { type: String, required: true },
  status:      { type: String, enum: ["draft", "submitted"], default: "draft" },
  addedBy:     { type: String, default: "" },
  addedByName: { type: String, default: "" },
  data: {
    propertyType: { type: String, default: "" },
    sft:          { type: String, default: "" },
    budget:       { type: String, default: "" },
    facing:       { type: String, default: "" },
    location:     { type: String, default: "" },
    amenities:    { type: String, default: "" },
    description:  { type: String, default: "" }
  }
}, { timestamps: true }));

// ── PAGE ROUTE ──
app.get("/property-requirement", requireAuth, (_req, res) =>
  res.sendFile(path.join(FRONTEND, "property-requirement.html")));

// ── GET ALL ──
app.get("/api/property-requirements", requireAuth, async (req, res) => {
  try {
    const records = await PropertyRequirement.find().sort({ createdAt: -1 });
    res.json({ records });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── CREATE ──
app.post("/api/property-requirements", requireAuth, async (req, res) => {
  try {
    const { customerId, propertyType, sft, budget, facing, location, amenities, description } = req.body;
    if (!propertyType || !sft || !budget)
      return res.status(400).json({ error: "Property type, SFT and budget are required" });

    // Auto-generate customerId if not provided
    let cid = (customerId || "").trim();
    if (!cid) cid = await generateCustomerId();

    const record = new PropertyRequirement({
      customerId:  cid,
      status:      "submitted",
      addedBy:     req.session.user?.email || "",
      addedByName: req.session.user?.name  || "",
      data: { propertyType, sft, budget, facing: facing||"", location: location||"", amenities: amenities||"", description: description||"" }
    });
    await record.save();
    res.json({ ok: true, record });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── UPDATE ──
app.put("/api/property-requirements/:id", requireAuth, async (req, res) => {
  try {
    const { propertyType, sft, budget, facing, location, amenities, description } = req.body;
    const record = await PropertyRequirement.findByIdAndUpdate(
      req.params.id,
      { $set: { "data.propertyType": propertyType, "data.sft": sft, "data.budget": budget,
                "data.facing": facing||"", "data.location": location||"",
                "data.amenities": amenities||"", "data.description": description||"",
                status: "submitted" } },
      { new: true }
    );
    if (!record) return res.status(404).json({ error: "Not found" });
    res.json({ ok: true, record });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── DELETE ──
app.delete("/api/property-requirements/:id", requireAuth, async (req, res) => {
  try {
    await PropertyRequirement.findByIdAndDelete(req.params.id);
    res.json({ ok: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── GENERATE PDF (no name/email/password) ──
app.get("/api/property-requirements/:id/pdf", requireAuth, async (req, res) => {
  try {
    const record = await PropertyRequirement.findById(req.params.id);
    if (!record) return res.status(404).json({ error: "Not found" });
    const d = record.data || {};

    // Build a clean HTML PDF using puppeteer-free approach with PDFKit
    const PDFDocument = require("pdfkit");
    const doc = new PDFDocument({ margin: 50, size: "A4" });

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition",
      `inline; filename="property-${record.customerId}.pdf"`);
    doc.pipe(res);

    // Header bar
    doc.rect(0, 0, 612, 80).fill("#329A9A");
    doc.fillColor("#ffffff").fontSize(22).font("Helvetica-Bold")
       .text("Property Requirement", 50, 25);
    doc.fontSize(11).font("Helvetica")
       .text(`Customer ID: ${record.customerId}`, 50, 52);

    doc.moveDown(3);
    doc.fillColor("#0a2a55").fontSize(14).font("Helvetica-Bold")
       .text("Property Details", 50, 100);
    doc.moveTo(50, 118).lineTo(545, 118).strokeColor("#329A9A").lineWidth(2).stroke();

    const fields = [
      ["Property Type",  d.propertyType || "—"],
      ["Square Feet",    d.sft ? d.sft + " sq.ft" : "—"],
      ["Budget",         d.budget || "—"],
      ["Facing",         d.facing || "—"],
      ["Location",       d.location || "—"],
      ["Amenities",      d.amenities || "—"],
    ];

    let y = 130;
    fields.forEach(([label, value]) => {
      doc.fillColor("#666666").fontSize(10).font("Helvetica-Bold").text(label, 50, y);
      doc.fillColor("#1a1a2e").fontSize(11).font("Helvetica").text(value, 200, y);
      y += 28;
      doc.moveTo(50, y - 6).lineTo(545, y - 6).strokeColor("#edf1f5").lineWidth(0.5).stroke();
    });

    if (d.description) {
      y += 10;
      doc.fillColor("#0a2a55").fontSize(13).font("Helvetica-Bold").text("Description", 50, y);
      y += 20;
      doc.fillColor("#556274").fontSize(11).font("Helvetica")
         .text(d.description, 50, y, { width: 495, lineGap: 4 });
    }

    // Footer
    doc.rect(0, 770, 612, 72).fill("#f0f4f8");
    doc.fillColor("#8d99ab").fontSize(9).font("Helvetica")
       .text("Generated by Customer Management System — Confidential",
             50, 782, { align: "center", width: 512 });
    doc.text(`Generated on: ${new Date().toLocaleDateString("en-IN")}`,
             50, 796, { align: "center", width: 512 });

    doc.end();
  } catch (err) {
    console.error("PDF error:", err.message);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

app.use((err, _req, res, _next) => { 
  console.error(err);
  res.status(500).json({ error: err.message });
});




