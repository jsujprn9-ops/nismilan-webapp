const express = require('express');
const multer = require('multer');
const nodemailer = require('nodemailer');
const { Resend } = require('resend');
const resend = new Resend(process.env.RESEND_API_KEY);
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(__dirname));

const CONFIG = {
  TOTAL_MEJA: 100,
  TOTAL_KERUSI: 800,
  PRICE_MEJA: 2000,
  PRICE_KERUSI: 250,
  EMAIL_USER: process.env.EMAIL_USER || 'jsujprn9@gmail.com',
  EMAIL_PASS: process.env.EMAIL_PASS || '',
  EMAIL_FROM: 'Malam Perdana #NismilanKitoTawan <jsujprn9@gmail.com>',
  DASHBOARD_SECRET: process.env.DASHBOARD_SECRET || 'jprn-urusetia-2026',
  DATA_FILE: path.join(__dirname, 'data', 'registrations.json'),
  UPLOADS_DIR: path.join(__dirname, 'uploads'),
};

fs.mkdirSync(path.join(__dirname, 'data'), { recursive: true });
fs.mkdirSync(CONFIG.UPLOADS_DIR, { recursive: true });

function loadData() {
  try {
    if (fs.existsSync(CONFIG.DATA_FILE)) return JSON.parse(fs.readFileSync(CONFIG.DATA_FILE, 'utf8'));
  } catch(e) {}
  return { registrations: [] };
}

function saveData(data) { fs.writeFileSync(CONFIG.DATA_FILE, JSON.stringify(data, null, 2)); }

function generateRefNo() {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  return `NSMN${yy}${mm}${dd}-${Math.floor(Math.random() * 9000) + 1000}`;
}

function computeStats(registrations) {
  const approved = registrations.filter(r => r.status === 'approved');
  let tablesBooked = 0, seatsBooked = 0, totalCollection = 0, attending = 0, vegetarian = 0;
  approved.forEach(r => {
    const qty = parseInt(r.quantity) || 0;
    if (r.ticketType === 'meja') { tablesBooked += qty; seatsBooked += qty * 10; }
    else seatsBooked += qty;
    totalCollection += parseInt(r.totalAmount) || 0;
    if (r.hadir === 'ya') attending += parseInt(r.hadirQty) || 0;
    if (r.menu === 'vege') vegetarian++;
  });
  return { tablesBooked, seatsTotal: seatsBooked, totalCollection, attending, vegetarian,
    pending: registrations.filter(r => r.status === 'pending').length, total: registrations.length };
}

const storage = multer.diskStorage({
  destination: CONFIG.UPLOADS_DIR,
  filename: (req, file, cb) => cb(null, `receipt_${Date.now()}${path.extname(file.originalname)}`)
});
const upload = multer({ storage, limits: { fileSize: 10 * 1024 * 1024 } });

async function sendEmail(to, subject, html) {
  if (!process.env.RESEND_API_KEY) { console.log('RESEND_API_KEY not set, skipping email'); return; }
  try {
    await resend.emails.send({
      from: 'onboarding@resend.dev',
      to,
      subject,
      html
    });
    console.log('Email sent to:', to);
  } catch (e) { console.error('Email failed:', e.message); }
}

async function sendConfirmationEmail(reg) {
  if (!reg.email) return;
  const typeStr = reg.ticketType === 'meja' ? `${reg.quantity} Meja` : `${reg.quantity} Kerusi`;
  const html = `<div style="font-family:Arial;max-width:500px;margin:0 auto;background:#111;color:#eee;padding:24px;border-radius:12px;">
    <h2 style="color:#C9A84C;">#NismilanKitoTawan</h2>
    <p>Assalamualaikum <strong>${reg.nama}</strong>,</p>
    <p>Tempahan anda telah diterima dan sedang dalam semakan.</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0;">
      <tr><td style="color:#888;padding:6px 0;">No. Rujukan</td><td style="color:#C9A84C;font-weight:700;">${reg.refNo}</td></tr>
      <tr><td style="color:#888;padding:6px 0;">Tempahan</td><td>${typeStr}</td></tr>
      <tr><td style="color:#888;padding:6px 0;">Jumlah</td><td style="color:#C9A84C;font-weight:700;">RM${parseInt(reg.totalAmount).toLocaleString()}</td></tr>
      <tr><td style="color:#888;padding:6px 0;">Tarikh</td><td>12 Julai 2026 (Sabtu)</td></tr>
      <tr><td style="color:#888;padding:6px 0;">Tempat</td><td>Casa Lagenda, Seremban</td></tr>
    </table>
    <p style="color:#888;font-size:13px;">Pertanyaan: <strong style="color:#C9A84C">013-6769 492</strong></p>
  </div>`;
  await sendEmail(reg.email, `Pengesahan Tempahan #${reg.refNo} — #NismilanKitoTawan`, html);
}

async function sendApprovedEmail(reg) {
  if (!reg.email) return;
  const typeStr = reg.ticketType === 'meja' ? `${reg.quantity} Meja` : `${reg.quantity} Kerusi`;
  const html = `<div style="font-family:Arial;max-width:500px;margin:0 auto;background:#111;color:#eee;padding:24px;border-radius:12px;">
    <h2 style="color:#27AE60;">✅ Pembayaran Disahkan!</h2>
    <p>Assalamualaikum <strong>${reg.nama}</strong>,</p>
    <p>Pembayaran anda telah <strong style="color:#27AE60">disahkan</strong>. Kami menantikan kehadiran anda!</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0;">
      <tr><td style="color:#888;padding:6px 0;">No. Rujukan</td><td style="color:#C9A84C;font-weight:700;">${reg.refNo}</td></tr>
      <tr><td style="color:#888;padding:6px 0;">Tempahan</td><td>${typeStr}</td></tr>
      <tr><td style="color:#888;padding:6px 0;">Jumlah</td><td style="color:#C9A84C;font-weight:700;">RM${parseInt(reg.totalAmount).toLocaleString()}</td></tr>
      <tr><td style="color:#888;padding:6px 0;">Tarikh</td><td>12 Julai 2026 — 7.00 malam</td></tr>
      <tr><td style="color:#888;padding:6px 0;">Tempat</td><td>Casa Lagenda Convention Center, Seremban</td></tr>
    </table>
    <p style="color:#888;font-size:13px;">Pertanyaan: <strong style="color:#C9A84C">013-6769 492</strong></p>
  </div>`;
  await sendEmail(reg.email, `✅ Pembayaran Disahkan! #${reg.refNo} — #NismilanKitoTawan`, html);
}

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get(`/urusetia/${CONFIG.DASHBOARD_SECRET}`, (req, res) => res.sendFile(path.join(__dirname, 'dashboard.html')));
app.use('/uploads', express.static(CONFIG.UPLOADS_DIR));

app.get('/api/availability', (req, res) => {
  const { registrations } = loadData();
  const stats = computeStats(registrations);
  res.json({ tablesBooked: stats.tablesBooked, seatsBooked: stats.seatsTotal });
});

app.post('/api/register', upload.single('receipt'), async (req, res) => {
  try {
    const { ticketType, quantity, totalAmount, nama, telefon, email, hadir, hadirQty, menu } = req.body;
    if (!ticketType || !quantity || !nama || !telefon || !hadir)
      return res.status(400).json({ success: false, message: 'Data tidak lengkap.' });
    const qty = parseInt(quantity);
    const expectedAmount = qty * (ticketType === 'meja' ? CONFIG.PRICE_MEJA : CONFIG.PRICE_KERUSI);
    const data = loadData();
    const registration = {
      id: uuidv4(), refNo: generateRefNo(), ticketType, quantity: qty,
      totalAmount: expectedAmount, nama, telefon, email: email || '', hadir,
      hadirQty: hadir === 'ya' ? parseInt(hadirQty) || 0 : 0,
      menu: menu || 'tiada',
      receiptFile: req.file?.filename || null,
      receiptUrl: req.file ? `/uploads/${req.file.filename}` : null,
      status: 'pending', createdAt: new Date().toISOString()
    };
    data.registrations.push(registration);
    saveData(data);
    if (email) sendConfirmationEmail(registration).catch(console.error);
    res.json({ success: true, refNo: registration.refNo, verified: false, message: 'Tempahan diterima!' });
  } catch(e) {
    console.error(e);
    res.status(500).json({ success: false, message: 'Ralat sistem.' });
  }
});

app.get('/api/admin/registrations', (req, res) => {
  const data = loadData();
  res.json({ registrations: data.registrations.sort((a,b) => new Date(b.createdAt)-new Date(a.createdAt)), stats: computeStats(data.registrations) });
});

app.patch('/api/admin/registrations/:id/status', async (req, res) => {
  const { id } = req.params;
  const { status } = req.body;
  const data = loadData();
  const idx = data.registrations.findIndex(r => r.id === id);
  if (idx === -1) return res.status(404).json({ success: false });
  const wasNotApproved = data.registrations[idx].status !== 'approved';
  data.registrations[idx].status = status;
  data.registrations[idx].updatedAt = new Date().toISOString();
  saveData(data);
  if (status === 'approved' && wasNotApproved && data.registrations[idx].email)
    sendApprovedEmail(data.registrations[idx]).catch(console.error);
  res.json({ success: true });
});

app.get('/api/admin/export', (req, res) => {
  const { registrations } = loadData();
  const rows = registrations.map((r,i) => ({
    'No.': i+1, 'No. Rujukan': r.refNo, 'Nama': r.nama, 'Telefon': r.telefon,
    'Emel': r.email||'', 'Jenis': r.ticketType==='meja'?'Meja':'Kerusi',
    'Qty': r.quantity, 'Jumlah': r.totalAmount,
    'Hadir': r.hadir==='ya'?'Ya':'Tidak', 'Bil. Hadir': r.hadirQty||0,
    'Menu': r.menu==='vege'?'Vegetarian':'Tiada',
    'Status': r.status==='approved'?'Disahkan':r.status==='rejected'?'Ditolak':'Semakan',
    'Tarikh': r.createdAt?new Date(r.createdAt).toLocaleString('ms-MY'):''
  }));
  const ws = xlsx.utils.json_to_sheet(rows);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Pendaftaran');
  const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="Pendaftaran_${new Date().toISOString().slice(0,10)}.xlsx"`);
  res.send(buf);
});

app.listen(PORT, () => {
  console.log(`Server running!`);
  console.log(`Borang: http://localhost:${PORT}`);
  console.log(`Dashboard: http://localhost:${PORT}/urusetia/${CONFIG.DASHBOARD_SECRET}`);
});
