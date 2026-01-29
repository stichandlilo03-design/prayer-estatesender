const express = require('express');
const nodemailer = require('nodemailer');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const XLSX = require('xlsx');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb' }));
app.use(express.static('public'));

// DATA STORES
let leads = [];
let templates = [
  { 
    id: 1, 
    name: 'Real Estate - Property Meeting',
    subject: 'Buyer Interest in {{propertyAddress}}',
    html: `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Property Meeting</title></head><body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif; background-color: #f8f8f8;"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #f8f8f8;"><tr><td align="center" style="padding: 40px 20px;"><table width="600" style="background-color: #ffffff;" border="0" cellpadding="0" cellspacing="0"><tr><td style="padding: 60px 40px 40px; background-color: #ffffff; text-align: center;"><p style="margin: 0 0 20px 0; font-size: 48px; padding: 0;">🏘️</p><h1 style="margin: 0 0 15px 0; color: #1a1a1a; font-size: 28px; font-weight: 300; line-height: 1.3; padding: 0;">Buyer Interest</h1><p style="margin: 0; color: #888888; font-size: 14px; padding: 0; text-transform: uppercase;">Virtual Meeting Invitation</p></td></tr><tr><td style="padding: 40px 40px; background-color: #ffffff;"><p style="margin: 0 0 30px 0; font-size: 16px; color: #333333; line-height: 1.8;">Hi {{firstName}},</p><p style="margin: 0 0 30px 0; font-size: 16px; color: #666666; line-height: 1.8;">A buyer is interested in viewing your property and would like to schedule a meeting.</p><div style="margin: 40px 0; padding: 30px 0; border-top: 1px solid #e0e0e0; border-bottom: 1px solid #e0e0e0;"><p style="margin: 0 0 8px 0; font-size: 12px; color: #999999; text-transform: uppercase;">Property Address</p><p style="margin: 0 0 20px 0; font-size: 20px; color: #1a1a1a; font-weight: 500;">{{propertyAddress}}</p><p style="margin: 0; font-size: 13px; color: #999999;">{{propertyType}} • {{propertyPrice}}</p></div><h3 style="margin: 30px 0 20px 0; font-size: 16px; color: #1a1a1a; font-weight: 500;">Meeting Overview</h3><ul style="margin: 0 0 30px 0; padding-left: 20px; color: #666666; font-size: 15px; line-height: 1.8;"><li style="margin-bottom: 8px;">Detailed property discussion</li><li style="margin-bottom: 8px;">Home features & history</li><li style="margin-bottom: 8px;">Neighborhood questions</li><li>Next steps overview</li></ul><table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin: 40px 0;"><tr><td align="center"><a href="{{zoomLink}}" style="display: inline-block; padding: 16px 48px; background-color: #1a1a1a; color: #ffffff; text-decoration: none; font-size: 15px; font-weight: 500;">Join Meeting</a></td></tr></table><div style="margin: 40px 0; padding: 25px; background-color: #f8f8f8;"><p style="margin: 0 0 15px 0; font-size: 12px; color: #999999; text-transform: uppercase;">Schedule</p><p style="margin: 0 0 8px 0; font-size: 16px; color: #1a1a1a;">{{meetingDate}}</p><p style="margin: 0; font-size: 16px; color: #1a1a1a;">{{meetingTime}} EST</p></div><p style="margin: 0 0 20px 0; font-size: 15px; color: #666666; line-height: 1.8;">Questions? Reach out anytime. We're here to help.</p></td></tr><tr><td style="padding: 40px 40px; background-color: #f8f8f8; border-top: 1px solid #e0e0e0;"><p style="margin: 0 0 8px 0; font-size: 15px; color: #1a1a1a; font-weight: 500;">{{senderName}}</p><p style="margin: 0 0 5px 0; font-size: 14px; color: #666666;">{{companyName}}</p><p style="margin: 0; font-size: 14px; color: #999999;">{{companyPhone}}</p></td></tr></table></td></tr></table></body></html>` 
  },
  { 
    id: 2, 
    name: 'Church - Prayer Meeting Zoom',
    subject: 'Join Us for {{prayerSection}} - {{meetingDate}}',
    html: `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Prayer Meeting Invitation</title></head><body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica Neue', sans-serif; background-color: #f5f5f5;"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5;"><tbody><tr><td align="center" style="padding: 40px 20px;"><table width="600" style="background-color: #ffffff;" border="0" cellpadding="0" cellspacing="0"><tr><td style="padding: 60px 40px 40px; background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%); text-align: center; color: white;"><p style="margin: 0 0 20px 0; font-size: 56px; padding: 0;">✝️</p><h1 style="margin: 0 0 10px 0; color: #ffffff; font-size: 32px; font-weight: 300; line-height: 1.3; padding: 0; letter-spacing: -0.5px;">Prayer Meeting</h1><p style="margin: 0; color: #ecf0f1; font-size: 14px; padding: 0; letter-spacing: 0.5px; text-transform: uppercase;">Join Us for Worship & Prayer</p></td></tr><tr><td style="padding: 40px 40px; background-color: #ffffff;"><p style="margin: 0 0 25px 0; font-size: 16px; color: #2c3e50; line-height: 1.8; padding: 0;">Dear {{firstName}},</p><p style="margin: 0 0 25px 0; font-size: 15px; color: #34495e; line-height: 1.8; padding: 0;">You are cordially invited to join our {{prayerSection}} prayer meeting. This is a sacred time for us to come together in faith, share our prayer requests, and intercede for our church, families, and community.</p><div style="margin: 35px 0; padding: 25px; background-color: #ecf0f1; border-left: 4px solid #3498db; border-radius: 4px;"><p style="margin: 0 0 15px 0; font-size: 12px; color: #2c3e50; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px; padding: 0;">Meeting Details</p><table width="100%" style="margin: 10px 0;"><tr><td style="padding: 8px 0; font-size: 14px; color: #2c3e50;"><strong>📅 Date:</strong> {{meetingDate}}</td></tr><tr><td style="padding: 8px 0; font-size: 14px; color: #2c3e50;"><strong>⏰ Time:</strong> {{meetingTime}}</td></tr><tr><td style="padding: 8px 0; font-size: 14px; color: #2c3e50;"><strong>🕯️ Focus:</strong> {{prayerSection}}</td></tr><tr><td style="padding: 8px 0; font-size: 14px; color: #2c3e50;"><strong>📍 Location:</strong> Online via Zoom</td></tr></table></div><table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin: 35px 0;"><tr><td align="center" style="padding: 0;"><a href="{{zoomLink}}" style="display: inline-block; padding: 16px 48px; background-color: #3498db; color: #ffffff; text-decoration: none; font-size: 15px; font-weight: 600; letter-spacing: 0.5px; border-radius: 4px;">Join Prayer Meeting</a></td></tr></table><div style="margin: 35px 0; padding: 25px; background-color: #f0f8ff; border-radius: 4px;"><p style="margin: 0 0 12px 0; font-size: 13px; color: #2c3e50; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px; padding: 0;">This Meeting's Focus Areas:</p><ul style="margin: 10px 0; padding-left: 20px; color: #34495e; font-size: 14px; line-height: 1.8;"><li style="margin-bottom: 8px;">Intercessory prayers for our church leadership</li><li style="margin-bottom: 8px;">Prayers for healing and restoration</li><li style="margin-bottom: 8px;">Community and missionary support prayers</li><li>Personal prayer requests from members</li></ul></div><p style="margin: 25px 0 0 0; font-size: 13px; color: #7f8c8d; line-height: 1.8; padding: 0;"><strong>How to Join:</strong><br>Click the "Join Prayer Meeting" button above to access the Zoom meeting. No special software needed - you can join from your computer, smartphone, or tablet. Please allow a few minutes to connect before the meeting starts.</p><p style="margin: 20px 0 0 0; font-size: 13px; color: #7f8c8d; line-height: 1.8; padding: 0;"><strong>Need Help?</strong><br>If you have any technical difficulties or questions, please contact us at {{companyPhone}} or reply to this email.</p></td></tr><tr><td style="padding: 30px 40px; background-color: #f0f8ff; border-top: 1px solid #e0e0e0; text-align: center;"><p style="margin: 0 0 15px 0; font-size: 14px; color: #2c3e50; font-weight: 500; padding: 0;">In Christ's Love,</p><p style="margin: 0 0 10px 0; font-size: 13px; color: #2c3e50; padding: 0;">{{senderName}}</p><p style="margin: 0; font-size: 12px; color: #7f8c8d; padding: 0;">{{companyName}}</p><p style="margin: 5px 0 0 0; font-size: 12px; color: #7f8c8d; padding: 0;">{{companyPhone}}</p></td></tr><tr><td style="padding: 25px 40px; background-color: #ecf0f1; border-top: 1px solid #bdc3c7; text-align: center;"><p style="margin: 0; font-size: 12px; color: #34495e; font-style: italic; line-height: 1.6; padding: 0;">"If my people, who are called by my name, will humble themselves and pray and seek my face and turn from their wicked ways, then I will hear from heaven, and I will forgive their sin and will heal their land." - 2 Chronicles 7:14</p></td></tr></table></td></tr></tbody></table></body></html>` 
  }
];
let campaigns = [];
let sendingLogs = [];
let smtpSettings = { 
  host: '', 
  port: 587, 
  secure: false, 
  user: '',
  fromEmail: '',
  senderName: '',
  companyName: '',
  companyPhone: '',
  pass: ''
};
let nextLeadId = 1;
let nextTemplateId = 3;
let nextCampaignId = 1;

const upload = multer({ storage: multer.memoryStorage() });

// ==================== LEADS ====================
app.get('/api/leads', (req, res) => { res.json(leads); });

app.post('/api/leads', (req, res) => {
  try {
    const { firstName, lastName, email, phone, propertyAddress, propertyPrice, propertyType } = req.body;
    if (!email) return res.json({ success: false, error: 'Email required' });
    if (leads.some(l => l.email === email)) return res.json({ success: false, error: 'Email exists' });
    leads.push({ 
      id: nextLeadId++, 
      firstName: firstName || '', 
      lastName: lastName || '', 
      email, 
      phone: phone || '', 
      propertyAddress: propertyAddress || '', 
      propertyPrice: propertyPrice || '', 
      propertyType: propertyType || '', 
      createdAt: new Date().toISOString() 
    });
    res.json({ success: true });
  } catch (err) { res.json({ success: false }); }
});

app.delete('/api/leads/:id', (req, res) => {
  leads = leads.filter(l => l.id !== parseInt(req.params.id));
  res.json({ success: true });
});

// ==================== PASTE ====================
app.post('/api/leads/paste', (req, res) => {
  try {
    const { text } = req.body;
    if (!text) return res.json({ success: false, inserted: 0, failed: 0 });
    const lines = text.trim().split('\n').filter(l => l.trim());
    if (lines.length === 0) return res.json({ success: false, inserted: 0, failed: 0 });
    const headers = lines[0].split(',').map(h => h.trim());
    const mapCol = (name) => {
      const n = name.toLowerCase();
      if (n.includes('first')) return 'firstName';
      if (n.includes('last')) return 'lastName';
      if (n.includes('email')) return 'email';
      if (n.includes('phone')) return 'phone';
      if (n.includes('address')) return 'propertyAddress';
      if (n.includes('price')) return 'propertyPrice';
      if (n.includes('type')) return 'propertyType';
      return null;
    };
    const colMap = {};
    headers.forEach((col, idx) => { const m = mapCol(col); if (m) colMap[idx] = m; });
    if (!Object.values(colMap).includes('email')) return res.json({ success: false, inserted: 0, failed: 0 });
    let inserted = 0, failed = 0;
    for (let i = 1; i < lines.length; i++) {
      try {
        const parts = lines[i].split(',').map(p => p.trim());
        const lead = {};
        Object.keys(colMap).forEach(idx => { lead[colMap[idx]] = parts[parseInt(idx)] || ''; });
        if (!lead.email || leads.some(l => l.email === lead.email)) { failed++; continue; }
        leads.push({ 
          id: nextLeadId++, 
          firstName: lead.firstName || '', 
          lastName: lead.lastName || '', 
          email: lead.email, 
          phone: lead.phone || '', 
          propertyAddress: lead.propertyAddress || '', 
          propertyPrice: lead.propertyPrice || '', 
          propertyType: lead.propertyType || '', 
          createdAt: new Date().toISOString() 
        });
        inserted++;
      } catch (e) { failed++; }
    }
    res.json({ success: true, inserted, failed });
  } catch (error) { res.json({ success: false, inserted: 0, failed: 0 }); }
});

// ==================== UPLOAD ====================
app.post('/api/leads/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.json({ success: false, inserted: 0, failed: 0 });
    const fileName = req.file.originalname.toLowerCase();
    let rows = [];
    if (fileName.endsWith('.csv')) {
      rows = parse(req.file.buffer.toString('utf-8'), { columns: true, skip_empty_lines: true });
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    } else {
      return res.json({ success: false, inserted: 0, failed: 0 });
    }
    if (!rows || rows.length === 0) return res.json({ success: false, inserted: 0, failed: 0 });
    const mapCol = (name) => {
      const n = (name || '').toLowerCase().trim();
      if (n.includes('first')) return 'firstName';
      if (n.includes('last')) return 'lastName';
      if (n.includes('email')) return 'email';
      if (n.includes('phone')) return 'phone';
      if (n.includes('address')) return 'propertyAddress';
      if (n.includes('price')) return 'propertyPrice';
      if (n.includes('type')) return 'propertyType';
      return null;
    };
    const colMap = {};
    const headers = Object.keys(rows[0] || {});
    headers.forEach(col => { const m = mapCol(col); if (m) colMap[col] = m; });
    if (!Object.values(colMap).includes('email')) return res.json({ success: false, inserted: 0, failed: 0 });
    let inserted = 0, failed = 0;
    for (const row of rows) {
      try {
        const lead = {};
        Object.keys(colMap).forEach(col => { lead[colMap[col]] = (row[col] || '').toString().trim(); });
        if (!lead.email || leads.some(l => l.email === lead.email)) { failed++; continue; }
        leads.push({ 
          id: nextLeadId++, 
          firstName: lead.firstName || '', 
          lastName: lead.lastName || '', 
          email: lead.email, 
          phone: lead.phone || '', 
          propertyAddress: lead.propertyAddress || '', 
          propertyPrice: lead.propertyPrice || '', 
          propertyType: lead.propertyType || '', 
          createdAt: new Date().toISOString() 
        });
        inserted++;
      } catch (e) { failed++; }
    }
    res.json({ success: true, inserted, failed });
  } catch (error) { res.json({ success: false, inserted: 0, failed: 0 }); }
});

// ==================== TEMPLATES ====================
app.get('/api/templates', (req, res) => { res.json(templates); });
app.get('/api/templates/:id', (req, res) => { const t = templates.find(x => x.id === parseInt(req.params.id)); res.json(t || {}); });
app.post('/api/templates', (req, res) => {
  try {
    const { id, name, subject, html } = req.body;
    if (!name || !subject || !html) return res.json({ success: false });
    if (id) {
      const idx = templates.findIndex(t => t.id === id);
      if (idx >= 0) templates[idx] = { id, name, subject, html };
    } else {
      templates.push({ id: nextTemplateId++, name, subject, html });
    }
    res.json({ success: true });
  } catch (err) { res.json({ success: false }); }
});

// ==================== SMTP ====================
app.get('/api/smtp-settings', (req, res) => { res.json(smtpSettings); });
app.post('/api/smtp-settings', (req, res) => {
  try {
    const { host, port, encryption, user, pass, senderName, fromEmail, companyName, companyPhone } = req.body;
    smtpSettings = {
      host: host || '',
      port: parseInt(port) || 587,
      encryption: encryption || 'true',
      user: user || '',
      pass: pass || '',
      senderName: senderName || '',
      fromEmail: fromEmail || user || '',
      companyName: companyName || '',
      companyPhone: companyPhone || ''
    };
    res.json({ success: true });
  } catch (err) { res.json({ success: false }); }
});

// ==================== SMTP TEST ====================
app.post('/api/smtp-test', async (req, res) => {
  try {
    if (!smtpSettings.host || !smtpSettings.user || !smtpSettings.pass) {
      return res.json({ success: false, error: 'Settings incomplete' });
    }

    const config = {
      host: smtpSettings.host,
      port: smtpSettings.port,
      auth: { user: smtpSettings.user, pass: smtpSettings.pass },
      tls: { rejectUnauthorized: false }
    };

    // Handle different encryption types
    if (smtpSettings.port === 465 || smtpSettings.encryption === 'true') {
      config.secure = true;  // SSL/TLS
    } else if (smtpSettings.encryption === 'starttls') {
      config.secure = false; // STARTTLS
    } else if (smtpSettings.encryption === 'none') {
      config.secure = false;
      config.tls = false;
    }

    const transporter = nodemailer.createTransport(config);
    const verified = await transporter.verify();
    
    if (verified) {
      res.json({ success: true, message: 'Connected successfully!' });
    } else {
      res.json({ success: false, error: 'Connection failed' });
    }
  } catch (error) {
    res.json({ success: false, error: error.message });
  }
});

// ==================== TEST EMAIL ====================
app.post('/api/smtp-test-email', async (req, res) => {
  try {
    const { email } = req.body;
    if (!email) return res.json({ success: false, error: 'Email required' });
    if (!smtpSettings.user || !smtpSettings.pass) return res.json({ success: false, error: 'SMTP not configured' });

    const config = {
      host: smtpSettings.host,
      port: smtpSettings.port,
      auth: { user: smtpSettings.user, pass: smtpSettings.pass },
      tls: { rejectUnauthorized: false }
    };

    // Handle different encryption types
    if (smtpSettings.port === 465 || smtpSettings.encryption === 'true') {
      config.secure = true;  // SSL/TLS
    } else if (smtpSettings.encryption === 'starttls') {
      config.secure = false; // STARTTLS
    } else if (smtpSettings.encryption === 'none') {
      config.secure = false;
      config.tls = false;
    }

    const transporter = nodemailer.createTransport(config);
    
    // Smart From Field
    let fromField = smtpSettings.user;
    if (smtpSettings.senderName && smtpSettings.fromEmail) {
      fromField = `${smtpSettings.senderName} <${smtpSettings.fromEmail}>`;
    } else if (smtpSettings.senderName) {
      fromField = `${smtpSettings.senderName} <${smtpSettings.user}>`;
    } else if (smtpSettings.fromEmail) {
      fromField = smtpSettings.fromEmail;
    }

    await transporter.sendMail({
      from: fromField,
      to: email,
      subject: 'Test Email - Campaign Manager',
      html: `<h1>✓ Success!</h1><p>Email sent from: ${smtpSettings.senderName || 'System'}</p><p>Your SMTP is configured correctly!</p>`
    });

    res.json({ success: true, message: 'Test email sent!' });
  } catch (error) {
    res.json({ success: false, error: error.message });
  }
});

// ==================== SEND CAMPAIGN ====================
app.post('/api/campaigns/send', async (req, res) => {
  try {
    const { templateId, zoomLink, meetingDate, meetingTime, prayerSection } = req.body;
    const template = templates.find(t => t.id === parseInt(templateId));
    if (!template) return res.json({ success: false, error: 'Template not found' });
    if (leads.length === 0) return res.json({ success: false, error: 'No leads' });
    if (!smtpSettings.user || !smtpSettings.pass) return res.json({ success: false, error: 'SMTP not configured' });
    const campaignId = nextCampaignId++;
    campaigns.push({ id: campaignId, name: 'Campaign ' + campaignId, templateId, leadsCount: leads.length, sentAt: new Date().toISOString(), status: 'sending' });
    const config = {
      host: smtpSettings.host,
      port: smtpSettings.port,
      auth: { user: smtpSettings.user, pass: smtpSettings.pass },
      tls: { rejectUnauthorized: false }
    };

    // Handle different encryption types
    if (smtpSettings.port === 465 || smtpSettings.encryption === 'true') {
      config.secure = true;  // SSL/TLS
    } else if (smtpSettings.encryption === 'starttls') {
      config.secure = false; // STARTTLS
    } else if (smtpSettings.encryption === 'none') {
      config.secure = false;
      config.tls = false;
    }
    const transporter = nodemailer.createTransport(config);
    let fromField = smtpSettings.user;
    if (smtpSettings.senderName && smtpSettings.fromEmail) {
      fromField = `${smtpSettings.senderName} <${smtpSettings.fromEmail}>`;
    } else if (smtpSettings.senderName) {
      fromField = `${smtpSettings.senderName} <${smtpSettings.user}>`;
    } else if (smtpSettings.fromEmail) {
      fromField = smtpSettings.fromEmail;
    }
    let sent = 0, failed = 0;
    for (const lead of leads) {
      try {
        let html = template.html;
        html = html
          .replace(/{{firstName}}/g, lead.firstName || '')
          .replace(/{{lastName}}/g, lead.lastName || '')
          .replace(/{{email}}/g, lead.email || '')
          .replace(/{{phone}}/g, lead.phone || '')
          .replace(/{{propertyAddress}}/g, lead.propertyAddress || '')
          .replace(/{{propertyPrice}}/g, lead.propertyPrice || '')
          .replace(/{{propertyType}}/g, lead.propertyType || '')
          .replace(/{{zoomLink}}/g, zoomLink || '')
          .replace(/{{meetingDate}}/g, meetingDate || '')
          .replace(/{{meetingTime}}/g, meetingTime || '')
          .replace(/{{prayerSection}}/g, prayerSection || '')
          .replace(/{{senderName}}/g, smtpSettings.senderName || 'Team')
          .replace(/{{companyName}}/g, smtpSettings.companyName || 'Company')
          .replace(/{{companyPhone}}/g, smtpSettings.companyPhone || 'N/A');
        let subject = template.subject;
        subject = subject
          .replace(/{{firstName}}/g, lead.firstName || '')
          .replace(/{{propertyAddress}}/g, lead.propertyAddress || '')
          .replace(/{{meetingDate}}/g, meetingDate || '')
          .replace(/{{prayerSection}}/g, prayerSection || '');
        await transporter.sendMail({
          from: fromField,
          to: lead.email,
          subject: subject,
          html: html
        });
        sent++;
        sendingLogs.push({
          campaignId,
          leadId: lead.id,
          leadEmail: lead.email,
          leadName: lead.firstName + ' ' + lead.lastName,
          status: 'success',
          message: 'Sent from: ' + (smtpSettings.senderName || smtpSettings.user),
          timestamp: new Date().toISOString()
        });
      } catch (e) {
        failed++;
        sendingLogs.push({
          campaignId,
          leadId: lead.id,
          leadEmail: lead.email,
          leadName: lead.firstName + ' ' + lead.lastName,
          status: 'failed',
          message: e.message,
          timestamp: new Date().toISOString()
        });
      }
    }
    const camp = campaigns.find(c => c.id === campaignId);
    if (camp) { camp.status = 'completed'; camp.sent = sent; camp.failed = failed; }
    res.json({ success: true, sent, failed, total: leads.length });
  } catch (error) {
    res.json({ success: false, error: error.message });
  }
});

// ==================== CAMPAIGNS & LOGS ====================
app.get('/api/campaigns/logs', (req, res) => { res.json(sendingLogs); });
app.get('/api/campaigns', (req, res) => { res.json(campaigns); });

// ==================== CLEAR ALL ====================
app.post('/api/clear-all', (req, res) => {
  leads = [];
  templates = [
    { id: 1, name: 'Real Estate - Property Meeting', subject: 'Buyer Interest in {{propertyAddress}}', html: '...' },
    { id: 2, name: 'Church - Prayer Meeting Zoom', subject: 'Join Us for {{prayerSection}} - {{meetingDate}}', html: '...' }
  ];
  campaigns = [];
  sendingLogs = [];
  nextLeadId = 1;
  nextTemplateId = 3;
  nextCampaignId = 1;
  res.json({ success: true });
});

module.exports = app;

if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
}
