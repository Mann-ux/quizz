/* Kuis Lokal - Vanilla JS */
const STORAGE_KEY = 'quizSimState_v1';
const LETTERS = ['A', 'B', 'C', 'D', 'E'];
const NEXT_DELAY = 800; // ms (< 1 detik)

const $ = (id) => document.getElementById(id);
const el = {
  navHeader: $('navHeader'), navBoxes: $('navBoxes'), phaseLabel: $('phaseLabel'),
  screenStart: $('screenStart'), screenQuestion: $('screenQuestion'), screenPreview: $('screenPreview'),
  fileInput: $('fileInput'), fileInfo: $('fileInfo'), fileError: $('fileError'),
  btnStart: $('btnStart'), resumeBox: $('resumeBox'),
  qCounter: $('qCounter'), qText: $('qText'), qOptions: $('qOptions'),
  explainBox: $('explainBox'), explainText: $('explainText'),
  btnPrev: $('btnPrev'), btnNext: $('btnNext'), btnSkip: $('btnSkip'),
  quizControls: $('quizControls'), kbdHint: $('kbdHint'),
};

// ===== State =====
// questions: [{text, options:[{label,text}], correct:index, explanation}] (sudah diacak)
// answers: array panjang N; null = belum dijawab, -1 = dilewati, >=0 = indeks opsi dipilih
let state = null;
let pendingQuestions = null; // hasil parse CSV sebelum kuis dimulai
let locked = false;          // kunci input saat delay
let selectingMode = false;   // sedang di layar pilihan mode
let hintOpen = false;        // pembahasan terbuka (Mode Santai)
let lastIdx = -1;            // untuk reset pembahasan saat pindah soal
const isSantai = () => state && state.mode === 'santai';

function save() {
  if (state) localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}
function clearSave() { localStorage.removeItem(STORAGE_KEY); }
function loadSave() {
  try {
    const s = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (s && Array.isArray(s.questions) && s.questions.length && Array.isArray(s.answers)) return s;
  } catch (e) { /* abaikan */ }
  return null;
}

// ===== XLSX (SheetJS) =====
const HEADERS = ['Pertanyaan', 'Opsi A', 'Opsi B', 'Opsi C', 'Opsi D', 'Opsi E', 'Kunci', 'Pembahasan'];
const norm = (s) => String(s).trim().toLowerCase().replace(/\s+/g, ' ');
const EXPECTED = HEADERS.map(norm);

// Membaca ArrayBuffer .xlsx -> Array of Objects ([{Pertanyaan:..., 'Opsi A':..., ...}])
function parseXLSX(buffer) {
  const wb = XLSX.read(buffer, { type: 'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) throw new Error('Sheet tidak ditemukan di file.');
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
}

// Array of Objects -> soal terstruktur (header dicocokkan tanpa peduli huruf besar/kecil)
function rowsToQuestions(objs) {
  if (!objs.length) throw new Error('File kosong atau tidak memiliki baris soal.');
  const rows = objs.map(o => {
    const m = {};
    Object.keys(o).forEach(k => { m[norm(k)] = o[k]; });
    return m;
  });
  for (const h of EXPECTED) {
    if (!(h in rows[0])) throw new Error(`Header kolom "${h}" tidak ditemukan. Gunakan format: ${HEADERS.join(', ')}.`);
  }
  const idx = {}; EXPECTED.forEach(h => { idx[h] = h; });
  const out = [];
  rows.forEach((r, n) => {
    const get = (h) => String(r[idx[h]] ?? '').trim();
    const text = get('pertanyaan');
    const key = get('kunci').toUpperCase();
    const opts = ['opsi a', 'opsi b', 'opsi c', 'opsi d', 'opsi e']
      .map((h, i) => ({ orig: i, text: get(h) })).filter(o => o.text !== '');
    if (!text) throw new Error(`Baris ${n + 2}: pertanyaan kosong.`);
    if (opts.length < 2) throw new Error(`Baris ${n + 2}: minimal 2 opsi jawaban.`);
    const keyIdx = LETTERS.indexOf(key);
    if (keyIdx === -1 || !opts.some(o => o.orig === keyIdx)) throw new Error(`Baris ${n + 2}: kunci "${key}" tidak valid.`);
    out.push({ text, opts, keyIdx, explanation: get('pembahasan') });
  });
  return out;
}

// ===== Shuffle (Fisher-Yates) =====
function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function buildQuiz(raw) {
  const questions = shuffle(raw).map(q => {
    const opts = shuffle(q.opts);
    return {
      text: q.text,
      options: opts.map(o => o.text),
      correct: opts.findIndex(o => o.orig === q.keyIdx),
      explanation: q.explanation,
    };
  });
  return { questions, answers: questions.map(() => null), current: 0, phase: 'quiz' };
}

// ===== Upload =====
el.fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  pendingQuestions = null;
  el.btnStart.classList.add('hidden');
  el.fileError.classList.add('hidden');
  if (!file) return;
  if (typeof XLSX === 'undefined') return showFileError('Library SheetJS gagal dimuat (cek koneksi internet).');
  if (!/\.xlsx$/i.test(file.name)) return showFileError('File harus berformat .xlsx');
  const reader = new FileReader();
  reader.onload = () => {
    try {
      pendingQuestions = rowsToQuestions(parseXLSX(reader.result));
      el.fileInfo.textContent = `✓ ${file.name} — ${pendingQuestions.length} soal terdeteksi`;
      el.btnStart.classList.remove('hidden');
    } catch (err) { showFileError(err.message || 'File tidak dapat dibaca.'); }
  };
  reader.onerror = () => showFileError('Gagal membaca file.');
  reader.readAsArrayBuffer(file);
});

// ===== Unduh template .xlsx (header bold + lebar kolom otomatis) =====
function downloadTemplate() {
  if (typeof XLSX === 'undefined') return alert('Library SheetJS gagal dimuat (cek koneksi internet).');
  const data = [
    HEADERS,
    ['Apa ibu kota Indonesia?', 'Bandung', 'Jakarta', 'Surabaya', 'Medan', 'Makassar', 'B', 'Jakarta adalah ibu kota Indonesia saat ini, pusat pemerintahan dan ekonomi.'],
    ['Berapakah hasil dari 7 x 8?', '54', '56', '58', '64', '48', 'B', '7 x 8 = 56. Cara cepat: 7 x 10 - 7 x 2 = 70 - 14 = 56.'],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  // Header: bold + latar abu-abu muda (butuh build SheetJS dengan dukungan style)
  HEADERS.forEach((_, c) => {
    const ref = XLSX.utils.encode_cell({ r: 0, c });
    ws[ref].s = {
      font: { bold: true },
      fill: { patternType: 'solid', fgColor: { rgb: 'DBEAFE' } },
      alignment: { vertical: 'center' },
    };
  });
  // Lebar kolom otomatis berdasarkan teks terpanjang (maks 60 karakter)
  ws['!cols'] = HEADERS.map((_, c) => ({
    wch: Math.min(60, Math.max(...data.map(r => String(r[c] ?? '').length)) + 2),
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Soal');
  XLSX.writeFile(wb, 'template-soal.xlsx');
}
$('btnTemplate').addEventListener('click', downloadTemplate);

function showFileError(msg) {
  el.fileInfo.textContent = 'Klik untuk memilih file';
  el.fileError.textContent = msg;
  el.fileError.classList.remove('hidden');
}

el.btnStart.addEventListener('click', () => {
  if (!pendingQuestions) return;
  selectingMode = true;
  render();
});

function startWithMode(mode) {
  state = buildQuiz(pendingQuestions);
  state.mode = mode;
  selectingMode = false;
  hintOpen = false; lastIdx = -1;
  save();
  render();
}
$('modeSantai').addEventListener('click', () => startWithMode('santai'));
$('modeSerius').addEventListener('click', () => startWithMode('serius'));
$('modeBack').addEventListener('click', () => { selectingMode = false; render(); });

$('btnResume').addEventListener('click', () => { state = loadSave(); if (state) render(); });
$('btnDiscard').addEventListener('click', () => { clearSave(); el.resumeBox.classList.add('hidden'); });

// ===== Render =====
function show(screen) {
  el.screenStart.classList.toggle('hidden', screen !== 'start');
  $('screenMode').classList.toggle('hidden', screen !== 'mode');
  el.screenQuestion.classList.toggle('hidden', !(screen === 'quiz' || screen === 'review'));
  el.screenPreview.classList.toggle('hidden', screen !== 'preview');
  el.navHeader.classList.toggle('hidden', screen === 'start' || screen === 'mode');
}

function render() {
  if (!state) { show(selectingMode ? 'mode' : 'start'); if (!selectingMode) checkResume(); return; }
  show(state.phase === 'preview' ? 'preview' : state.phase);
  renderNav();
  $('btnFinish').classList.toggle('hidden', state.phase !== 'quiz');
  el.phaseLabel.textContent = { quiz: 'Kuis', preview: 'Hasil', review: 'Review' }[state.phase];
  if (state.phase === 'preview') renderPreview();
  else renderQuestion();
  window.scrollTo({ top: 0 });
}

function status(i) { // 'correct' | 'wrong' | 'skip' | 'answered' | 'empty'
  const a = state.answers[i];
  if (state.phase === 'quiz') return a !== null && a >= 0 ? 'answered' : 'empty';
  if (a === null || a === -1) return 'skip';
  return a === state.questions[i].correct ? 'correct' : 'wrong';
}

function renderNav() {
  el.navBoxes.innerHTML = '';
  state.questions.forEach((_, i) => {
    const b = document.createElement('button');
    b.textContent = i + 1;
    const st = status(i);
    const color = {
      empty: 'bg-slate-200 text-slate-700', answered: 'bg-blue-900 text-white', skip: 'bg-slate-300 text-slate-700',
      correct: 'bg-green-500 text-white', wrong: 'bg-red-500 text-white',
    }[st];
    const active = (state.phase !== 'preview' && i === state.current) ? 'ring-4 ring-offset-1 ring-amber-400 border-2 border-black' : 'border border-transparent';
    b.className = `w-9 h-9 rounded-md text-sm font-semibold transition hover:scale-110 ${color} ${active}`;
    b.title = `Soal ${i + 1}`;
    b.addEventListener('click', () => jumpTo(i));
    b.dataset.i = i;
    el.navBoxes.appendChild(b);
  });
  const act = el.navBoxes.querySelector(`[data-i="${state.current}"]`);
  if (act && state.phase !== 'preview') act.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

function jumpTo(i) {
  if (locked) return;
  if (state.phase === 'preview') state.phase = 'review'; // klik dari preview -> Mode Review
  state.current = i;
  save(); render();
}

function renderQuestion() {
  if (state.current !== lastIdx || state.phase !== 'quiz') { hintOpen = false; lastIdx = state.current; }
  const i = state.current, q = state.questions[i], ans = state.answers[i];
  const review = state.phase === 'review';
  el.qCounter.textContent = `Soal ${i + 1} dari ${state.questions.length}`;
  el.qText.textContent = q.text;
  el.qOptions.innerHTML = '';
  q.options.forEach((txt, k) => {
    const d = document.createElement(review ? 'div' : 'button');
    let cls = 'w-full text-left flex items-start gap-3 p-3 sm:p-4 rounded-xl border-2 transition ';
    let mark = '';
    if (review) {
      if (k === q.correct) { cls += 'border-green-500 bg-green-50'; mark = '<span class="ml-auto text-green-600 font-bold">✓</span>'; }
      else if (k === ans) { cls += 'border-red-500 bg-red-50'; mark = '<span class="ml-auto text-red-600 font-bold">✗</span>'; }
      else cls += 'border-slate-200 bg-white';
    } else {
      cls += (k === ans) ? 'border-blue-600 bg-blue-50 cursor-pointer' : 'border-slate-200 bg-white hover:border-blue-300 hover:bg-slate-50 cursor-pointer';
      d.addEventListener('click', () => selectOption(k));
    }
    d.className = cls;
    d.innerHTML = `<span class="shrink-0 w-7 h-7 rounded-full bg-slate-200 text-sm font-bold flex items-center justify-center">${LETTERS[k]}</span><span class="flex-1"></span>${mark}`;
    d.children[1].textContent = txt;
    el.qOptions.appendChild(d);
  });

  el.explainBox.classList.toggle('hidden', !(review || (isSantai() && hintOpen)));
  if (review || (isSantai() && hintOpen)) {
    el.explainText.textContent = q.explanation || 'Tidak ada pembahasan untuk soal ini.';
  }
  $('btnHint').classList.toggle('hidden', review || !isSantai());
  $('btnHint').textContent = hintOpen ? 'Sembunyikan' : 'Pembahasan';
  el.kbdHint.textContent = review ? 'Gunakan ← → atau Enter untuk berpindah soal.'
    : 'Tekan A–E untuk memilih, ← → untuk pindah, Enter = Next' + (isSantai() ? ', Spasi = pembahasan.' : '.');
  // kontrol
  el.btnSkip.classList.toggle('hidden', review);
  el.kbdHint.classList.toggle('hidden', review && false);
  el.btnPrev.disabled = i === 0;
  el.btnPrev.classList.toggle('opacity-40', i === 0);
  el.btnPrev.textContent = review ? '← Sebelumnya' : '← Sebelumnya';
  const last = i === state.questions.length - 1;
  if (review) {
    el.btnNext.textContent = last ? 'Ke Ringkasan' : 'Selanjutnya →';
  } else {
    el.btnNext.textContent = last ? 'Selesai' : 'Next →';
  }
  el.btnNext.className = 'px-6 py-2.5 rounded-lg bg-blue-700 text-white hover:bg-blue-800 font-medium';
}

function renderPreview() {
  let c = 0, w = 0, s = 0;
  state.questions.forEach((_, i) => {
    const st = status(i);
    if (st === 'correct') c++; else if (st === 'wrong') w++; else s++;
  });
  const total = state.questions.length;
  const pct = Math.round((c / total) * 100);
  $('sumCorrect').textContent = c; $('sumWrong').textContent = w; $('sumSkip').textContent = s;
  $('scorePct').textContent = pct + '%';
  const ring = $('ring'), circ = 2 * Math.PI * 52;
  ring.style.strokeDasharray = circ;
  ring.style.strokeDashoffset = circ;
  requestAnimationFrame(() => requestAnimationFrame(() => { ring.style.strokeDashoffset = circ * (1 - c / total); }));
}

// ===== Aksi kuis =====
function selectOption(k) {
  if (locked || state.phase !== 'quiz') return;
  state.answers[state.current] = k;
  save(); renderNav(); renderQuestion();
}

function goPreview() {
  state.phase = 'preview';
  state.answers = state.answers.map(a => a === null ? -1 : a); // belum dijawab = dilewati
  save(); render();
}

function next() {
  if (locked) return;
  const i = state.current, last = i === state.questions.length - 1;
  if (state.phase === 'review') {
    if (last) { state.phase = 'preview'; save(); render(); }
    else { state.current++; save(); render(); }
    return;
  }
  if (state.phase !== 'quiz') return;
  const a = state.answers[i];
  const advance = () => {
    locked = false;
    if (last) goPreview(); else { state.current++; save(); render(); }
  };
  // evaluasi tertutup
  if (a !== null && a >= 0 && a !== state.questions[i].correct) {
    locked = true;
    el.btnNext.classList.add('flash-red');
    setTimeout(() => { el.btnNext.classList.remove('flash-red'); advance(); }, NEXT_DELAY);
  } else if (a !== null && a >= 0) {
    locked = true;
    setTimeout(advance, 250);
  } else advance(); // belum dijawab: pindah saja (tetap kosong)
}

function prev() {
  if (locked || state.current === 0) return;
  state.current--; save(); render();
}

function skip() {
  if (locked || state.phase !== 'quiz') return;
  const i = state.current;
  state.answers[i] = -1;
  const last = i === state.questions.length - 1;
  if (last) goPreview(); else { state.current++; save(); render(); }
}

el.btnNext.addEventListener('click', next);
el.btnPrev.addEventListener('click', prev);
el.btnSkip.addEventListener('click', skip);
$('btnHint').addEventListener('click', toggleHint);
function toggleHint() {
  if (locked || !state || state.phase !== 'quiz' || !isSantai()) return; // Mode Serius: dinonaktifkan
  hintOpen = !hintOpen;
  renderQuestion();
}
$('btnFinish').addEventListener('click', () => {
  if (locked) return;
  const left = state.answers.filter(a => a === null || a === -1).length;
  if (left > 0 && !confirm(`Masih ada ${left} soal belum dijawab. Selesaikan kuis sekarang?`)) return;
  goPreview();
});
$('btnReview').addEventListener('click', () => { state.phase = 'review'; state.current = 0; save(); render(); });
$('btnRestart').addEventListener('click', exitToStart);
$('btnExit').addEventListener('click', () => {
  if (state.phase === 'quiz' && !confirm('Keluar ke layar awal? Progres tetap tersimpan dan bisa dilanjutkan.')) return;
  if (state.phase !== 'quiz') clearSave();
  exitToStart(true);
});

function exitToStart(keepSave) {
  if (keepSave !== true && state && state.phase === 'preview') clearSave();
  state = null; pendingQuestions = null; locked = false; selectingMode = false; hintOpen = false;
  el.fileInput.value = ''; el.fileInfo.textContent = 'Klik untuk memilih file';
  el.btnStart.classList.add('hidden'); el.fileError.classList.add('hidden');
  render();
}

// ===== Keyboard =====
document.addEventListener('keydown', (e) => {
  if (!state || state.phase === 'preview' || e.ctrlKey || e.metaKey || e.altKey) return;
  const k = e.key;
  if (k === 'Enter') { e.preventDefault(); if (!e.repeat) next(); }
  else if (k === ' ' || k === 'Spacebar') {
    e.preventDefault();
    if (!e.repeat) { if (document.activeElement) document.activeElement.blur(); toggleHint(); }
  }
  else if (k === 'ArrowRight') { e.preventDefault(); next(); }
  else if (k === 'ArrowLeft') { e.preventDefault(); prev(); }
  else if (state.phase === 'quiz' && k.length === 1) {
    const idx = LETTERS.indexOf(k.toUpperCase());
    if (idx !== -1 && idx < state.questions[state.current].options.length) selectOption(idx);
  }
});

// ===== Init =====
function checkResume() {
  el.resumeBox.classList.toggle('hidden', !loadSave());
}
render();
