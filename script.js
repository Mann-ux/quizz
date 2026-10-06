/* Kuis Lokal - Vanilla JS */
const STORAGE_KEY = 'quizSimState_v1';
const LETTERS = ['A', 'B', 'C', 'D'];
const NEXT_DELAY = 800; // ms (< 1 detik)

const $ = (id) => document.getElementById(id);
const el = {
  navHeader: $('navHeader'), navBoxes: $('navBoxes'), phaseLabel: $('phaseLabel'),
  screenStart: $('screenStart'), screenQuestion: $('screenQuestion'), screenPreview: $('screenPreview'),
  fileInput: $('fileInput'), fileInfo: $('fileInfo'), fileError: $('fileError'),
  resumeBox: $('resumeBox'),
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
const HEADERS = ['Kategori', 'Materi', 'Pertanyaan', 'Opsi A', 'Opsi B', 'Opsi C', 'Opsi D', 'Kunci', 'Pembahasan'];
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
    const category = get('kategori');
    const key = get('kunci').toUpperCase();
    const opts = ['opsi a', 'opsi b', 'opsi c', 'opsi d']
      .map((h, i) => ({ orig: i, text: get(h) })).filter(o => o.text !== '');
    if (!category) throw new Error(`Baris ${n + 2}: kategori kosong.`);
    if (!text) throw new Error(`Baris ${n + 2}: pertanyaan kosong.`);
    if (opts.length < 2) throw new Error(`Baris ${n + 2}: minimal 2 opsi jawaban.`);
    const keyIdx = LETTERS.indexOf(key);
    if (keyIdx === -1 || !opts.some(o => o.orig === keyIdx)) throw new Error(`Baris ${n + 2}: kunci "${key}" tidak valid.`);
    out.push({ category, material: get('materi'), text, opts, keyIdx, explanation: get('pembahasan') });
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
  el.fileError.classList.add('hidden');
  if (!file) return;
  if (typeof XLSX === 'undefined') return showFileError('Library SheetJS gagal dimuat (cek koneksi internet).');
  if (!/\.xlsx$/i.test(file.name)) return showFileError('File harus berformat .xlsx');
  const reader = new FileReader();
  reader.onload = () => {
    try {
      pendingQuestions = rowsToQuestions(parseXLSX(reader.result));
      el.fileInfo.textContent = `✓ ${file.name} — ${pendingQuestions.length} soal terdeteksi`;
      selectingMode = true; // tampilkan Dashboard Belajar
      render();
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
    ['Pre-test', '', 'Apa ibu kota Indonesia?', 'Bandung', 'Jakarta', 'Surabaya', 'Medan', 'B', 'Jakarta adalah ibu kota Indonesia saat ini, pusat pemerintahan dan ekonomi.'],
    ['Sesi 1', 'Perkalian adalah penjumlahan berulang. Contoh: 7 x 8 = 7 dijumlahkan 8 kali = 56.', 'Berapakah hasil dari 7 x 8?', '54', '56', '58', '64', 'B', '7 x 8 = 56. Cara cepat: 7 x 10 - 7 x 2 = 70 - 14 = 56.'],
    ['Final Test', '', 'Berapakah hasil dari 9 x 6?', '54', '56', '63', '45', 'A', '9 x 6 = 54.'],
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

// Mode otomatis: kategori berisi "sesi" -> Santai; "test" -> Serius (default Serius)
function modeFor(cat) {
  const c = cat.toLowerCase();
  if (c.includes('sesi')) return 'santai';
  return 'serius';
}

function renderDashboard() {
  const list = $('dashList');
  list.innerHTML = '';
  
  // Pastikan panel kategori muncul dan panel bacaan materi sembunyi
  $('dashCategories').classList.remove('hidden');
  $('dashMateriPanel').classList.add('hidden');
  
  // Tombol "Baca Materi Rangkuman" global
  const hasMat = (q) => String(q.material || '').trim() !== '';
  const firstGlobalMat = pendingQuestions.find(hasMat);
  if (firstGlobalMat) {
    const btnMat = document.createElement('button');
    btnMat.className = 'sm:col-span-2 text-center p-4 rounded-xl border border-amber-600 bg-amber-600 text-white transition hover:bg-amber-700 hover:-translate-y-0.5 shadow-lg font-bold text-lg';
    btnMat.textContent = '📚 Baca Materi Rangkuman';
    btnMat.addEventListener('click', () => {
      $('dashCategories').classList.add('hidden');
      $('dashMateriPanel').classList.remove('hidden');
      $('dashMateriText').textContent = firstGlobalMat.material.trim();
    });
    list.appendChild(btnMat);
  }

  const cats = [...new Set(pendingQuestions.map(q => q.category))]; // urutan kemunculan di Excel
  cats.forEach(cat => {
    const n = pendingQuestions.filter(q => q.category === cat).length;
    const santai = modeFor(cat) === 'santai';
    const b = document.createElement('button');
    b.className = 'text-left p-5 rounded-xl border border-slate-700 bg-indigo-500 text-white transition hover:bg-indigo-600 hover:-translate-y-0.5 shadow-lg';
    b.innerHTML = '<div class="text-lg font-bold"></div><p class="text-sm text-indigo-100 mt-1"></p>';
    b.children[0].textContent = (santai ? '☕ ' : '🎯 ') + cat;
    b.children[1].textContent = `${n} soal · ${santai ? 'Mode Santai' : 'Mode Serius'}`;
    b.addEventListener('click', () => startCategory(cat));
    list.appendChild(b);
  });
}

function startCategory(cat) {
  const subset = pendingQuestions.filter(q => q.category === cat);
  if (!subset.length) return;
  state = buildQuiz(subset);
  state.mode = modeFor(cat);
  state.category = cat;
  // Materi global: Materi pertama di kategori ini; bila kategori tidak punya (mis. Pre-test/Final Test),
  // pakai Materi pertama yang ditemukan di seluruh file.
  const hasMat = (q) => String(q.material || '').trim() !== '';
  const firstMat = subset.find(hasMat) || pendingQuestions.find(hasMat);
  state.globalMateri = firstMat ? String(firstMat.material).trim() : '';
  globalMateri = state.globalMateri;
  selectingMode = false;
  hintOpen = false; lastIdx = -1;
  save();
  render();
}
$('modeBack').addEventListener('click', () => exitToStart(true, true));
$('btnDashMateriBack').addEventListener('click', () => {
  $('dashMateriPanel').classList.add('hidden');
  $('dashCategories').classList.remove('hidden');
});

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
  if (!state) {
    if (selectingMode && pendingQuestions) renderDashboard(); else selectingMode = false;
    show(selectingMode ? 'mode' : 'start'); if (!selectingMode) checkResume(); return;
  }
  show(state.phase === 'preview' ? 'preview' : state.phase);
  globalMateri = state.globalMateri || '';
  closeMateri();
  updateMateriButton();
  renderNav();
  $('btnFinish').classList.toggle('hidden', state.phase !== 'quiz');
  el.phaseLabel.textContent = (state.category ? state.category + ' · ' : '') + { quiz: 'Kuis', preview: 'Hasil', review: 'Review' }[state.phase];
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
      empty: 'bg-slate-700 border border-slate-600 text-slate-200', answered: 'bg-indigo-500 border border-indigo-500 text-white', skip: 'bg-slate-700 border border-slate-600 text-slate-400',
      correct: 'bg-emerald-500 border border-emerald-500 text-white', wrong: 'bg-rose-500 border border-rose-500 text-white',
    }[st];
    const active = (state.phase !== 'preview' && i === state.current) ? 'ring-2 ring-inset ring-indigo-300' : '';
    b.className = `w-10 h-10 flex items-center justify-center rounded-md text-sm font-semibold transition-all ${color} ${active}`;
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
  $('qProgress').style.width = `${((i + 1) / state.questions.length) * 100}%`;
  el.qText.textContent = q.text;
  el.qOptions.innerHTML = '';
  q.options.forEach((txt, k) => {
    const d = document.createElement(review ? 'div' : 'button');
    let cls = 'w-full text-left flex items-start gap-3 p-3 sm:p-4 rounded-xl border-2 text-slate-200 transition ';
    let mark = '';
    if (review) {
      if (k === q.correct) { cls += 'border-emerald-500 bg-slate-700'; mark = '<span class="ml-auto text-emerald-500 font-bold">✓</span>'; }
      else if (k === ans) { cls += 'border-rose-500 bg-slate-700'; mark = '<span class="ml-auto text-rose-500 font-bold">✗</span>'; }
      else cls += 'border-slate-600 bg-slate-700';
    } else {
      cls += (k === ans) ? 'bg-indigo-500/20 border-indigo-500 ring-1 ring-indigo-500 cursor-pointer' : 'bg-slate-800 border-slate-600 hover:bg-slate-700 cursor-pointer';
      d.addEventListener('click', () => selectOption(k));
    }
    d.className = cls;
    d.innerHTML = `<span class="shrink-0 w-7 h-7 rounded-full bg-slate-800 border border-slate-600 text-slate-200 text-sm font-bold flex items-center justify-center">${LETTERS[k]}</span><span class="flex-1"></span>${mark}`;
    d.children[1].textContent = txt;
    el.qOptions.appendChild(d);
  });

  el.explainBox.classList.toggle('hidden', !(review || (isSantai() && hintOpen)));
  if (review || (isSantai() && hintOpen)) {
    el.explainText.textContent = q.explanation || 'Tidak ada pembahasan untuk soal ini.';
  }
  $('btnHint').classList.toggle('hidden', review || !isSantai());
  $('btnMateri').classList.toggle('hidden', !currentMaterial());
  $('btnHint').textContent = hintOpen ? 'Sembunyikan' : 'Pembahasan';
  el.kbdHint.textContent = review ? 'Gunakan ← → atau Enter untuk berpindah soal.'
    : 'Tekan A–D untuk memilih, ← → untuk pindah, Enter = Next' + (isSantai() ? ', Spasi = pembahasan.' : '.');
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
  el.btnNext.className = 'px-6 py-2.5 rounded-lg bg-indigo-500 text-white hover:bg-indigo-600 transition font-medium';
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
  
  if (pct >= 75 && typeof confetti !== 'undefined') {
    confetti({ particleCount: 150, spread: 70, origin: { y: 0.6 } });
  }
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
$('btnRestart').addEventListener('click', () => exitToStart(false));
$('btnExit').addEventListener('click', () => {
  if (state.phase === 'quiz' && !confirm('Keluar ke Dashboard? Progres tetap tersimpan dan bisa dilanjutkan.')) return;
  if (state.phase !== 'quiz') clearSave();
  exitToStart(true);
});

// Kembali ke Dashboard bila file masih dimuat; fullReset = kembali ke layar unggah
function exitToStart(keepSave, fullReset) {
  if (keepSave !== true && state && state.phase === 'preview') clearSave();
  state = null; locked = false; hintOpen = false;
  if (fullReset || !pendingQuestions) {
    pendingQuestions = null; selectingMode = false;
    el.fileInput.value = ''; el.fileInfo.textContent = 'Klik untuk memilih file';
    el.fileError.classList.add('hidden');
  } else selectingMode = true;
  render();
}

// ===== Referensi Materi (Bottom Sheet) =====
let globalMateri = '';
let materiOpen = false;
const currentMaterial = () => (globalMateri || '').trim();
// Visibilitas tombol Materi HANYA bergantung pada isi globalMateri (tidak peduli mode/kategori)
function updateMateriButton() {
  $('btnMateri').classList.toggle('hidden', !currentMaterial());
}
function openMateri() {
  const m = currentMaterial();
  if (!state || state.phase === 'preview' || !m || locked) return;
  $('materiText').textContent = m;
  materiOpen = true;
  if (document.activeElement) document.activeElement.blur();
  $('materiBackdrop').classList.remove('hidden');
  // paksa reflow agar transisi berjalan
  void $('materiSheet').offsetHeight;
  $('materiSheet').classList.remove('translate-y-full');
  $('materiSheet').classList.add('translate-y-0');
}
function closeMateri() {
  materiOpen = false;
  $('materiSheet').classList.remove('translate-y-0');
  $('materiSheet').classList.add('translate-y-full');
  $('materiBackdrop').classList.add('hidden');
}
function toggleMateri() { materiOpen ? closeMateri() : openMateri(); }
$('btnMateri').addEventListener('click', toggleMateri);
$('btnMateriClose').addEventListener('click', closeMateri);
$('materiBackdrop').addEventListener('click', closeMateri);

// ===== Keyboard =====
document.addEventListener('keydown', (e) => {
  if (!state || state.phase === 'preview' || e.ctrlKey || e.metaKey || e.altKey) return;
  const k = e.key;
  if (materiOpen) { // saat panel terbuka, hanya M / Esc yang aktif
    if (k === 'm' || k === 'M' || k === 'Escape') { e.preventDefault(); closeMateri(); }
    else if (k === 'Enter' || k === ' ') e.preventDefault();
    return;
  }
  if ((k === 'm' || k === 'M') && !e.repeat) { e.preventDefault(); toggleMateri(); return; }
  if (k === 'Enter') { e.preventDefault(); if (!e.repeat) next(); }
  else if (k === ' ' || k === 'Spacebar') {
    e.preventDefault();
    if (isSantai() && !e.repeat) { if (document.activeElement) document.activeElement.blur(); toggleHint(); }
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
