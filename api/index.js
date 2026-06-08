const express = require('express');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const cors = require('cors');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');
const multer = require('multer');
const OpenAI = require('openai');

// Konfigurasi dari environment variables
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);
const upload = multer({ storage: multer.memoryStorage() });

// Inisialisasi OpenAI
let openai = null;
if (process.env.OPENAI_API_KEY) {
    openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

const app = express();

// Middleware
app.use(cors({ origin: true, credentials: true }));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
    secret: process.env.SESSION_SECRET || 'rahasia',
    resave: false,
    saveUninitialized: false,
    cookie: { secure: false, httpOnly: true, maxAge: 24 * 60 * 60 * 1000 }
}));

// Middleware auth
function requireAdmin(req, res, next) {
    if (req.session.user && req.session.user.role === 'admin') return next();
    res.status(401).json({ error: 'Unauthorized' });
}
function requireGuru(req, res, next) {
    if (req.session.user && (req.session.user.role === 'guru' || req.session.user.role === 'admin')) return next();
    res.status(401).json({ error: 'Unauthorized' });
}

// ========== ROUTES ==========

// Auth
app.post('/api/auth/login/admin', async (req, res) => {
    const { username, password } = req.body;
    const { data, error } = await supabase.from('admin').select('*').eq('username', username).single();
    if (error || !data) return res.status(401).json({ error: 'Invalid credentials' });
    const valid = await bcrypt.compare(password, data.password);
    if (!valid) return res.status(401).json({ error: 'Invalid credentials' });
    req.session.user = { id: data.id, role: 'admin', nama: data.nama_lengkap };
    res.json({ success: true, role: 'admin', redirect: 'index.html' });
});

app.post('/api/auth/login/guru', async (req, res) => {
    const { nip, password } = req.body;
    const { data, error } = await supabase.from('guru').select('*').eq('nip', nip).single();
    if (error || !data) return res.status(401).json({ error: 'Invalid NIP or password' });
    const valid = await bcrypt.compare(password, data.password);
    if (!valid) return res.status(401).json({ error: 'Invalid password' });
    req.session.user = { id: data.id, role: 'guru', nama: data.nama_lengkap };
    res.json({ success: true, role: 'guru', redirect: 'dashboard_guru.html' });
});

app.post('/api/auth/logout', (req, res) => {
    req.session.destroy();
    res.json({ success: true });
});

// Admin
app.get('/api/admin', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('admin').select('*').order('id');
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.post('/api/admin', requireAdmin, async (req, res) => {
    const { username, password, nama_lengkap, is_utama } = req.body;
    const hashed = await bcrypt.hash(password, 10);
    const { data, error } = await supabase.from('admin').insert([{ username, password: hashed, nama_lengkap, is_utama }]).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.delete('/api/admin/:id', requireAdmin, async (req, res) => {
    const { id } = req.params;
    const { data: admin } = await supabase.from('admin').select('is_utama').eq('id', id).single();
    if (admin?.is_utama) return res.status(403).json({ error: 'Cannot delete main admin' });
    const { error } = await supabase.from('admin').delete().eq('id', id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});

// Guru
app.get('/api/guru', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('guru').select('*').order('id');
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.post('/api/guru', requireAdmin, async (req, res) => {
    const { nip, username, password, nama_lengkap } = req.body;
    const { data: existing } = await supabase.from('guru').select('nip, username').or(`nip.eq.${nip},username.eq.${username}`);
    if (existing && existing.length) return res.status(400).json({ error: 'NIP atau Username sudah ada' });
    const hashed = await bcrypt.hash(password, 10);
    const { data, error } = await supabase.from('guru').insert([{ nip, username, password: hashed, nama_lengkap }]).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.delete('/api/guru/:id', requireAdmin, async (req, res) => {
    const { error } = await supabase.from('guru').delete().eq('id', req.params.id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});
app.get('/api/guru/dashboard', requireGuru, async (req, res) => {
    const guruId = req.session.user.id;
    const { data: pengawasList, error } = await supabase
        .from('pengawas')
        .select('ujian: id_ujian (*, mata_pelajaran: id_mapel (nama_mapel))')
        .eq('id_guru', guruId);
    if (error) return res.status(500).json({ error: error.message });
    const ujian = pengawasList.map(p => p.ujian);
    res.json({ ujian, jumlahPengawasan: ujian.length, ujianAktif: ujian.filter(u => u.status === 'aktif').length });
});

// Siswa
app.get('/api/siswa', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('siswa').select('*, kelas: id_kelas (nama_kelas)').order('id');
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.post('/api/siswa', requireAdmin, async (req, res) => {
    const { nis, username, password, nama_lengkap, id_kelas } = req.body;
    if (!password || password.length < 4) return res.status(400).json({ error: 'Password minimal 4 karakter' });
    const { data: existing } = await supabase.from('siswa').select('nis, username').or(`nis.eq.${nis},username.eq.${username}`);
    if (existing && existing.length) return res.status(400).json({ error: 'NIS atau Username sudah terdaftar' });
    const hashed = await bcrypt.hash(password, 10);
    const { data, error } = await supabase.from('siswa').insert([{ nis, username, password: hashed, nama_lengkap, id_kelas }]).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.delete('/api/siswa/:id', requireAdmin, async (req, res) => {
    const { error } = await supabase.from('siswa').delete().eq('id', req.params.id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});
app.get('/api/siswa/template', requireAdmin, async (req, res) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Template Siswa');
    worksheet.columns = [
        { header: 'NIS', key: 'nis', width: 15 },
        { header: 'Nama Lengkap', key: 'nama_lengkap', width: 30 },
        { header: 'Username', key: 'username', width: 20 },
        { header: 'Password', key: 'password', width: 20 },
        { header: 'Nama Kelas', key: 'kelas', width: 10 }
    ];
    worksheet.addRow({ nis: '001', nama_lengkap: 'Contoh Siswa', username: 'siswa1', password: 'siswa123', kelas: '9A' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="template_siswa.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
});
app.post('/api/siswa/import', requireAdmin, upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'File tidak ditemukan' });
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.getWorksheet(1);
        if (!worksheet) return res.status(400).json({ error: 'Sheet tidak ditemukan' });
        const rows = [];
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const nis = row.getCell(1).toString().trim();
            const nama_lengkap = row.getCell(2).toString().trim();
            const username = row.getCell(3).toString().trim();
            const password = row.getCell(4).toString().trim();
            const nama_kelas = row.getCell(5).toString().trim();
            if (nis && nama_lengkap && username && password && nama_kelas) rows.push({ nis, nama_lengkap, username, password, nama_kelas });
        });
        const { data: kelasList, error: kelasError } = await supabase.from('kelas').select('id, nama_kelas');
        if (kelasError) throw kelasError;
        const kelasMap = new Map(kelasList.map(k => [k.nama_kelas, k.id]));
        let successCount = 0, failedCount = 0, errors = [];
        for (const row of rows) {
            const id_kelas = kelasMap.get(row.nama_kelas);
            if (!id_kelas) { failedCount++; errors.push(`Kelas "${row.nama_kelas}" tidak ditemukan`); continue; }
            const { data: existing } = await supabase.from('siswa').select('nis, username').or(`nis.eq.${row.nis},username.eq.${row.username}`);
            if (existing && existing.length) { failedCount++; errors.push(`Duplikat NIS/Username: ${row.nis}`); continue; }
            const hashed = await bcrypt.hash(row.password, 10);
            const { error } = await supabase.from('siswa').insert({ nis: row.nis, nama_lengkap: row.nama_lengkap, username: row.username, password: hashed, id_kelas });
            if (error) { failedCount++; errors.push(error.message); } else successCount++;
        }
        res.json({ success: true, successCount, failedCount, errors: errors.slice(0,10) });
    } catch(err) { res.status(500).json({ error: err.message }); }
});

// Kelas
app.get('/api/kelas', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('kelas').select('*').order('id');
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.post('/api/kelas', requireAdmin, async (req, res) => {
    const { nama_kelas } = req.body;
    const { data, error } = await supabase.from('kelas').insert([{ nama_kelas }]).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.delete('/api/kelas/:id', requireAdmin, async (req, res) => {
    const { error } = await supabase.from('kelas').delete().eq('id', req.params.id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});

// Mata Pelajaran
app.get('/api/mapel', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('mata_pelajaran').select('*').order('id');
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.post('/api/mapel', requireAdmin, async (req, res) => {
    const { nama_mapel } = req.body;
    const { data, error } = await supabase.from('mata_pelajaran').insert([{ nama_mapel }]).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.delete('/api/mapel/:id', requireAdmin, async (req, res) => {
    const { error } = await supabase.from('mata_pelajaran').delete().eq('id', req.params.id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});

// Ujian
app.get('/api/ujian', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('ujian').select('*, mata_pelajaran: id_mapel (nama_mapel)').order('id', { ascending: false });
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.get('/api/ujian/:id', requireGuru, async (req, res) => {
    const { id } = req.params;
    const { data: ujian, error } = await supabase.from('ujian').select('*, mata_pelajaran: id_mapel (nama_mapel)').eq('id', id).single();
    if (error) return res.status(500).json({ error: error.message });
    if (req.session.user.role === 'guru') {
        const { data: pengawas } = await supabase.from('pengawas').select('id').eq('id_guru', req.session.user.id).eq('id_ujian', id).single();
        if (!pengawas) return res.status(403).json({ error: 'Tidak diizinkan' });
    }
    if (req.query.include_nilai === 'true') {
        const { data: nilai } = await supabase.from('nilai_ujian').select('*, siswa: id_siswa (*, kelas: id_kelas (nama_kelas))').eq('id_ujian', id);
        if (nilai) ujian.nilai_siswa = nilai;
    }
    res.json(ujian);
});
app.post('/api/ujian', requireAdmin, async (req, res) => {
    const { nama_ujian, id_mapel, tanggal_mulai, tanggal_akhir, durasi_menit, status } = req.body;
    const { data, error } = await supabase.from('ujian').insert([{ nama_ujian, id_mapel, tanggal_mulai, tanggal_akhir, durasi_menit, status }]).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.put('/api/ujian/:id', requireAdmin, async (req, res) => {
    const { id } = req.params;
    const updates = req.body;
    const { data, error } = await supabase.from('ujian').update(updates).eq('id', id).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.delete('/api/ujian/:id', requireAdmin, async (req, res) => {
    const { error } = await supabase.from('ujian').delete().eq('id', req.params.id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});

// Soal
app.get('/api/soal/ujian/:idUjian', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('soal').select('*').eq('id_ujian', req.params.idUjian).order('id');
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.post('/api/soal', requireAdmin, async (req, res) => {
    const { id_ujian, jenis_soal, pertanyaan, poin, pilihan_json, jawaban_json } = req.body;
    const { data, error } = await supabase.from('soal').insert([{ id_ujian, jenis_soal, pertanyaan, poin, pilihan_json, jawaban_json }]).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.put('/api/soal/:id', requireAdmin, async (req, res) => {
    const updates = req.body;
    const { data, error } = await supabase.from('soal').update(updates).eq('id', req.params.id).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.delete('/api/soal/:id', requireAdmin, async (req, res) => {
    const { error } = await supabase.from('soal').delete().eq('id', req.params.id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});

// Pengawas
app.get('/api/pengawas', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('pengawas').select('*, guru: id_guru (*), ujian: id_ujian (*)').order('id');
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.post('/api/pengawas', requireAdmin, async (req, res) => {
    const { id_guru, id_ujian } = req.body;
    const { data, error } = await supabase.from('pengawas').insert([{ id_guru, id_ujian }]).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json(data[0]);
});
app.delete('/api/pengawas/:id', requireAdmin, async (req, res) => {
    const { error } = await supabase.from('pengawas').delete().eq('id', req.params.id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});

// Token
function generateToken() {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    let token = '';
    for (let i = 0; i < 5; i++) token += chars[Math.floor(Math.random() * chars.length)];
    return token;
}
app.post('/api/token/refresh', requireAdmin, async (req, res) => {
    const { id_ujian } = req.body;
    if (!id_ujian) return res.status(400).json({ error: 'id_ujian required' });
    const newToken = generateToken();
    const { data, error } = await supabase.from('token_ujian').upsert({ id_ujian, token: newToken, is_active: true, expires_at: new Date(Date.now() + 24*60*60*1000) }).select();
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true, token: newToken });
});

// Monitoring
app.get('/api/monitoring/active-exams', requireAdmin, async (req, res) => {
    const now = new Date().toISOString();
    const { data, error } = await supabase
        .from('ujian')
        .select(`*, mata_pelajaran: id_mapel (nama_mapel), token_ujian (*), sesi_ujian_siswa (id, status, siswa: id_siswa (nama_lengkap))`)
        .eq('status', 'aktif')
        .lte('tanggal_mulai', now)
        .gte('tanggal_akhir', now);
    if (error) return res.status(500).json({ error: error.message });
    res.json(data);
});
app.post('/api/monitoring/update-sesi', requireAdmin, async (req, res) => {
    const { id_sesi, status_baru } = req.body;
    const { error } = await supabase.from('sesi_ujian_siswa').update({ status: status_baru, waktu_terakhir_aktivitas: new Date() }).eq('id', id_sesi);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true });
});

// Rekap
app.get('/api/rekap', requireAdmin, async (req, res) => {
    const { kelas, mapel } = req.query;
    let query = supabase.from('nilai_ujian').select(`*, siswa:id_siswa (nama_lengkap, nis, kelas:id_kelas (nama_kelas)), ujian:id_ujian (nama_ujian, mata_pelajaran:mapel_id (nama_mapel))`);
    if (kelas && kelas !== '') query = query.eq('siswa.kelas.nama_kelas', kelas);
    if (mapel && mapel !== '') query = query.eq('ujian.mata_pelajaran.nama_mapel', mapel);
    const { data, error } = await query;
    if (error) return res.status(500).json({ error: error.message });
    const result = data.map(item => ({
        id: item.id, siswa: item.siswa?.nama_lengkap || '-', nis: item.siswa?.nis || '-',
        kelas: item.siswa?.kelas?.nama_kelas || '-', mata_pelajaran: item.ujian?.mata_pelajaran?.nama_mapel || '-',
        ujian: item.ujian?.nama_ujian || '-', nilai: item.nilai_total, tanggal: item.selesai_pada ? new Date(item.selesai_pada).toLocaleDateString('id-ID') : '-'
    }));
    res.json(result);
});
app.get('/api/rekap/stats', requireAdmin, async (req, res) => {
    const { data, error } = await supabase.from('nilai_ujian').select('nilai_total, siswa: id_siswa (kelas: id_kelas (nama_kelas))');
    if (error) return res.status(500).json({ error: error.message });
    const totalNilai = data.reduce((s, n) => s + n.nilai_total, 0);
    const rataRata = data.length ? (totalNilai / data.length).toFixed(1) : 0;
    const kelasSet = new Set(data.map(d => d.siswa?.kelas?.nama_kelas).filter(Boolean));
    res.json({ rataRata, totalData: data.length, jumlahKelas: kelasSet.size });
});

// Export
app.get('/api/export/nilai', requireAdmin, async (req, res) => {
    const { kelas, mapel } = req.query;
    let query = supabase.from('nilai_ujian').select(`*, siswa:id_siswa (nama_lengkap, nis, kelas:id_kelas (nama_kelas)), ujian:id_ujian (nama_ujian, mata_pelajaran:mapel_id (nama_mapel))`);
    if (kelas && kelas !== '') query = query.eq('siswa.kelas.nama_kelas', kelas);
    if (mapel && mapel !== '') query = query.eq('ujian.mata_pelajaran.nama_mapel', mapel);
    const { data, error } = await query;
    if (error) return res.status(500).json({ error: error.message });
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Rekap Nilai');
    worksheet.columns = [
        { header: 'No', key: 'no', width: 6 }, { header: 'NIS', key: 'nis', width: 12 }, { header: 'Nama Siswa', key: 'siswa', width: 25 },
        { header: 'Kelas', key: 'kelas', width: 8 }, { header: 'Mata Pelajaran', key: 'mapel', width: 20 },
        { header: 'Ujian', key: 'ujian', width: 25 }, { header: 'Nilai', key: 'nilai', width: 8 }, { header: 'Tanggal', key: 'tanggal', width: 15 }
    ];
    worksheet.getRow(1).font = { bold: true };
    data.forEach((item, idx) => {
        worksheet.addRow({
            no: idx+1, nis: item.siswa?.nis, siswa: item.siswa?.nama_lengkap, kelas: item.siswa?.kelas?.nama_kelas,
            mapel: item.ujian?.mata_pelajaran?.nama_mapel, ujian: item.ujian?.nama_ujian,
            nilai: item.nilai_total, tanggal: item.selesai_pada ? new Date(item.selesai_pada).toLocaleDateString('id-ID') : ''
        });
    });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="rekap_nilai.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
});

// AI
async function callAI(prompt, systemPrompt = null) {
    if (process.env.USE_OLLAMA === 'true') {
        const response = await fetch(process.env.OLLAMA_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ model: process.env.OLLAMA_MODEL, prompt, stream: false, system: systemPrompt }) });
        const data = await response.json();
        return data.response;
    }
    if (!openai) throw new Error('AI tidak dikonfigurasi');
    const messages = [];
    if (systemPrompt) messages.push({ role: 'system', content: systemPrompt });
    messages.push({ role: 'user', content: prompt });
    const completion = await openai.chat.completions.create({ model: 'gpt-3.5-turbo', messages, temperature: 0.7, max_tokens: 1000 });
    return completion.choices[0].message.content;
}
function parseSoalResponse(response, jenis, id_ujian) {
    try {
        const jsonMatch = response.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            const parsed = JSON.parse(jsonMatch[0]);
            return { id_ujian, jenis_soal: parsed.jenis_soal || jenis, pertanyaan: parsed.pertanyaan, poin: parsed.poin || 1, pilihan_json: parsed.pilihan_json || null, jawaban_json: parsed.jawaban_json };
        }
        throw new Error('No JSON found');
    } catch(e) { return null; }
}
app.post('/api/ai/generate', requireAdmin, async (req, res) => {
    try {
        const { id_ujian, jenis_soal, topik, jumlah_soal = 1, level_kognitif = 'L2', kategori = 'C3' } = req.body;
        if (!id_ujian || !jenis_soal || !topik) return res.status(400).json({ error: 'id_ujian, jenis_soal, topik required' });
        const systemPrompt = `Anda adalah guru pembuat soal. Buat soal ${jenis_soal} level ${level_kognitif} kategori ${kategori}. Output dalam JSON valid.`;
        const prompt = `Buat ${jumlah_soal} soal ${jenis_soal} tentang "${topik}". Format: { "jenis_soal": "...", "pertanyaan": "...", "poin": 1, "pilihan_json": {...}, "jawaban_json": ... }`;
        const aiResponse = await callAI(prompt, systemPrompt);
        const soalData = parseSoalResponse(aiResponse, jenis_soal, id_ujian);
        if (!soalData) return res.status(500).json({ error: 'Gagal parsing AI response' });
        const { data, error } = await supabase.from('soal').insert([soalData]).select();
        if (error) return res.status(500).json({ error: error.message });
        res.json({ success: true, soal: data[0], raw: aiResponse });
    } catch(err) { res.status(500).json({ error: err.message }); }
});

// Export handler untuk Vercel
module.exports = app;