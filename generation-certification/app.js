/* Newton Academy — генератор сертификатов.
   Всё работает в браузере: данные учеников никуда не отправляются. */
(function () {
  'use strict';

  /* ---------- Константы шаблона (в пунктах PDF, отсчёт сверху) ---------- */
  var PAGE_W = 842, PAGE_H = 595, PX = 2; // канвас: 2 пикселя на пункт (1684×1190)
  var VIOLET = [69, 65, 250], BLACK = [0, 0, 0];
  var X0 = 328.8;
  var LAY = {
    intro: { f: 'regular', s: 15.3, y: 203, c: BLACK, maxW: 446, min: 11 },
    name:  { f: 'medium',  s: 24,   y: 257, c: VIOLET, maxW: 434, min: 12 },
    body:  { f: 'regular', b: 'bold', s: 15.3, y: 334, lh: 19, maxW: 372, c: BLACK },
    dir:   { f: 'bold',    s: 15.3, y: 509.5, c: VIOLET, maxW: 340, min: 10 },
    dirT:  { f: 'italic',  s: 13,   y: 528, c: BLACK, maxW: 340, min: 9 }
  };
  var FONT_FILES = {
    regular: 'assets/fonts/Montserrat-Regular.ttf',
    medium:  'assets/fonts/Montserrat-Medium.ttf',
    bold:    'assets/fonts/Montserrat-Bold.ttf',
    italic:  'assets/fonts/Montserrat-Italic.ttf'
  };
  var CANVAS_FAMILY = { regular: 'NA-Regular', medium: 'NA-Medium', bold: 'NA-Bold', italic: 'NA-Italic' };

  var DEFAULT_SIGN = {
    ruName: 'Солихова Комила Мухиддиновна', ruTitle: 'Директор',
    uzName: 'Solixova Komila Muxiddinovna', uzTitle: 'Direktor'
  };

  /* ---------- Интерфейс: RU / UZ ---------- */
  var I18N = {
    ru: {
      title: 'Генератор сертификатов',
      heroTitle: 'Сертификаты Newton за минуту',
      heroSub: 'Введите данные ученика или загрузите Excel — получите готовые PDF в фирменном шаблоне. Без регистрации. Данные не покидают ваш браузер.',
      tabSingle: 'Один сертификат', tabBulk: 'Массово из Excel',
      surname: 'Фамилия', name: 'Имя', cls: 'Класс', course: 'Курс', certLang: 'Язык сертификата',
      phSurname: 'Анваров', phName: 'Акромбек', phCls: '4 или 4А', phCourse: 'Впишите название курса', demoCourse: 'Почемучка',
      gender: 'Пол', boy: 'Мальчик', girl: 'Девочка', thGender: 'Пол',
      genderHint: 'Определяется по фамилии — можно поменять',
      myCourses: 'Мои курсы:', saveCourse: 'Сохранить в мои курсы', removeCourse: 'Удалить из списка',
      courseSaved: 'Курс добавлен в список', gM: 'М', gF: 'Ж', gToggle: 'Нажмите, чтобы поменять пол', gAuto: 'определено по фамилии',
      generate: 'Сгенерировать сертификат', jpg: 'Картинка JPG', print: 'Печать',
      preview: 'Предпросмотр', previewHint: 'Обновляется по мере ввода',
      signature: 'Подпись', edit: 'изменить',
      settingsTitle: 'Подпись на сертификатах', sRuName: 'ФИО (русский)', sRuTitle: 'Должность (русский)',
      sUzName: 'ФИО (узбекский)', sUzTitle: 'Должность (узбекский)', save: 'Сохранить', reset: 'Вернуть по умолчанию', close: 'Закрыть',
      settingsNote: 'Сохраняется только в этом браузере.',
      step1: 'Скачайте шаблон Excel', step1Sub: 'Колонки: Фамилия, Имя, Класс, Пол (м/ж), Курс, Язык. Можно загрузить и свою таблицу с такими заголовками.',
      downloadTpl: 'Скачать шаблон .xlsx',
      step2: 'Загрузите заполненную таблицу', drop: 'Перетащите файл сюда или нажмите, чтобы выбрать', dropSub: '.xlsx или .csv',
      step3: 'Проверьте и скачайте', defCourse: 'Курс, если в таблице пусто', defLang: 'Язык, если в таблице пусто',
      thSurname: 'Фамилия', thName: 'Имя', thCls: 'Класс', thCourse: 'Курс', thLang: 'Язык',
      zipPdf: 'ZIP: отдельный PDF на каждого', onePdf: 'Один PDF для печати', zipJpg: 'ZIP: картинки JPG',
      cancel: 'Отменить', rowsReady: 'Готово к генерации', rowsErr: 'с ошибками (будут пропущены)',
      errName: 'нет ФИО', errCourse: 'нет курса', clear: 'Загрузить другой файл',
      progress: 'Создаю', of: 'из', done: 'Готово! Файл скачан.', canceled: 'Отменено.',
      needFields: 'Заполните фамилию или имя и курс', loading: 'Загружаю шаблон…', ready: 'Шаблон готов',
      xlsOld: 'Старый формат .xls не поддерживается. Откройте файл в Excel и сохраните как .xlsx.',
      readErr: 'Не удалось прочитать файл. Проверьте, что это .xlsx или .csv.',
      noRows: 'В таблице не нашлось строк с учениками.',
      clickRow: 'Нажмите на строку, чтобы посмотреть сертификат',
      footer: 'Newton Academy · Данные обрабатываются только в вашем браузере',
      langRu: 'Русский', langUz: 'O‘zbekcha', certs: 'Сертификаты'
    },
    uz: {
      title: 'Sertifikat generatori',
      heroTitle: 'Newton sertifikatlari bir daqiqada',
      heroSub: 'O‘quvchi ma’lumotlarini kiriting yoki Excel yuklang — brend shablonida tayyor PDF oling. Ro‘yxatdan o‘tish shart emas. Ma’lumotlar brauzeringizdan chiqmaydi.',
      tabSingle: 'Bitta sertifikat', tabBulk: 'Excel orqali ommaviy',
      surname: 'Familiya', name: 'Ism', cls: 'Sinf', course: 'Kurs', certLang: 'Sertifikat tili',
      phSurname: 'Anvarov', phName: 'Akrombek', phCls: '4 yoki 4A', phCourse: 'Kurs nomini yozing', demoCourse: 'Pochemuchka',
      gender: 'Jinsi', boy: 'O‘g‘il bola', girl: 'Qiz bola', thGender: 'Jinsi',
      genderHint: 'Familiyadan aniqlanadi — o‘zgartirish mumkin',
      myCourses: 'Mening kurslarim:', saveCourse: 'Kurslarimga saqlash', removeCourse: 'Ro‘yxatdan o‘chirish',
      courseSaved: 'Kurs ro‘yxatga qo‘shildi', gM: 'O‘', gF: 'Q', gToggle: 'Jinsini o‘zgartirish uchun bosing', gAuto: 'familiyadan aniqlandi',
      generate: 'Sertifikat yaratish', jpg: 'JPG rasm', print: 'Chop etish',
      preview: 'Oldindan ko‘rish', previewHint: 'Yozganingiz sari yangilanadi',
      signature: 'Imzo', edit: 'o‘zgartirish',
      settingsTitle: 'Sertifikatdagi imzo', sRuName: 'F.I.Sh. (rus tilida)', sRuTitle: 'Lavozim (rus tilida)',
      sUzName: 'F.I.Sh. (o‘zbek tilida)', sUzTitle: 'Lavozim (o‘zbek tilida)', save: 'Saqlash', reset: 'Standartga qaytarish', close: 'Yopish',
      settingsNote: 'Faqat shu brauzerda saqlanadi.',
      step1: 'Excel shablonini yuklab oling', step1Sub: 'Ustunlar: Familiya, Ism, Sinf, Jinsi (o‘g‘il/qiz), Kurs, Til. Shunday sarlavhali o‘z jadvalingizni ham yuklash mumkin.',
      downloadTpl: '.xlsx shablonni yuklab olish',
      step2: 'To‘ldirilgan jadvalni yuklang', drop: 'Faylni shu yerga tashlang yoki tanlash uchun bosing', dropSub: '.xlsx yoki .csv',
      step3: 'Tekshiring va yuklab oling', defCourse: 'Jadvalda bo‘sh bo‘lsa, kurs', defLang: 'Jadvalda bo‘sh bo‘lsa, til',
      thSurname: 'Familiya', thName: 'Ism', thCls: 'Sinf', thCourse: 'Kurs', thLang: 'Til',
      zipPdf: 'ZIP: har biriga alohida PDF', onePdf: 'Chop etish uchun bitta PDF', zipJpg: 'ZIP: JPG rasmlar',
      cancel: 'Bekor qilish', rowsReady: 'Yaratishga tayyor', rowsErr: 'xato bilan (o‘tkazib yuboriladi)',
      errName: 'F.I.Sh. yo‘q', errCourse: 'kurs yo‘q', clear: 'Boshqa fayl yuklash',
      progress: 'Yaratilmoqda', of: '/', done: 'Tayyor! Fayl yuklab olindi.', canceled: 'Bekor qilindi.',
      needFields: 'Familiya yoki ism va kursni kiriting', loading: 'Shablon yuklanmoqda…', ready: 'Shablon tayyor',
      xlsOld: 'Eski .xls formati qo‘llab-quvvatlanmaydi. Faylni Excelda ochib, .xlsx sifatida saqlang.',
      readErr: 'Faylni o‘qib bo‘lmadi. Bu .xlsx yoki .csv ekanini tekshiring.',
      noRows: 'Jadvalda o‘quvchilar qatori topilmadi.',
      clickRow: 'Sertifikatni ko‘rish uchun qatorni bosing',
      footer: 'Newton Academy · Ma’lumotlar faqat brauzeringizda qayta ishlanadi',
      langRu: 'Русский', langUz: 'O‘zbekcha', certs: 'Sertifikatlar'
    }
  };

  /* ---------- Утилиты ---------- */
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  function store(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; } }
  function clean(s) { return String(s == null ? '' : s).replace(/[ \s]+/g, ' ').trim(); }
  function capWord(w) {
    // Приводим к виду «Анваров», если слово набрано целиком заглавными или строчными
    if (w.length < 2) return w.toUpperCase();
    var up = w.toUpperCase(), lo = w.toLowerCase();
    if (w !== up && w !== lo) return w;
    return lo.replace(/(^|-)(\S)/g, function (m, p, ch) { return p + ch.toUpperCase(); });
  }
  function normName(s) { return clean(s).split(' ').filter(Boolean).map(capWord).join(' '); }
  function stripQuotes(s) { return clean(s).replace(/^[«"„“'‘]+|[»"“”'’]+$/g, '').trim(); }
  function fixUz(s) { return s.replace(/([OoGg])['`ʻʼ’]/g, '$1‘'); }
  function safeFile(s) { return clean(s).replace(/[\\/:*?"<>|]+/g, '').slice(0, 120) || 'sertifikat'; }
  function download(blob, filename) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  function tick() { return new Promise(function (r) { setTimeout(r, 0); }); }

  /* ---------- Класс: «4А» → «4 «А» класса» / «4-«A» sinf» ---------- */
  function classPhrase(raw, lang) {
    var s = clean(raw).replace(/(класс[а-я]*|sinf[a-z‘']*|синф[а-я]*|kl\.?|class|grade)/gi, '').replace(/[«»"']/g, '').trim();
    if (!s) return '';
    var m = s.match(/^(\d{1,2})\s*[-–—.]?\s*([A-Za-zА-Яа-яЁёЎўҚқҒғҲҳ])?$/);
    if (m) {
      var n = m[1], L = m[2] ? m[2].toUpperCase() : '';
      if (lang === 'uz') return L ? n + '-«' + L + '» sinf' : n + '-sinf';
      return L ? n + ' «' + L + '» класса' : n + ' класса';
    }
    return lang === 'uz' ? s + ' sinf' : s + ' класса';
  }

  /* ---------- Пол: 'm' / 'f' / '' (не определён) ---------- */
  function guessGender(surname, name) {
    var full = (clean(surname) + ' ' + clean(name)).toLowerCase().replace(/[‘’'ʻʼ`]/g, '');
    if (/(^|\s)(qizi|kizi|қизи|кизи)(\s|$)/.test(full)) return 'f';
    if (/(^|\s)(ogli|ugli|ўғли|угли|огли)(\s|$)/.test(full)) return 'm';
    var s = clean(surname).toLowerCase().replace(/[‘’'ʻʼ`]/g, '').split(/[\s-]/).pop() || '';
    if (s.length < 3) return '';
    if (/(ова|ева|ёва|ина|ына|ая|яя)$/.test(s) || /(ova|eva|yova|ina|aya|yaya)$/.test(s)) return 'f';
    if (/(ов|ев|ёв|ин|ын|ий|ый|ой)$/.test(s) || /(ov|ev|yov|in|iy|yy|oy)$/.test(s)) return 'm';
    return '';
  }
  function genderOf(v) {
    var s = String(v == null ? '' : v).toLowerCase().replace(/[^a-zа-яёўқғҳ]+/g, '');
    if (!s) return '';
    if (/^(ж|f|дев|girl|qiz|қиз|киз|ayol|аёл|д$|q$|к$)/.test(s)) return 'f';
    if (/^(м|m|boy|ogil|ўғил|угил|огил|erkak|эркак|o$|у$|ў$)/.test(s)) return 'm';
    return '';
  }

  /* ---------- Модель текста сертификата ---------- */
  function sign(lang) {
    var st = getSign();
    return lang === 'uz' ? { name: st.uzName, title: st.uzTitle } : { name: st.ruName, title: st.ruTitle };
  }
  function buildModel(d) {
    var lang = d.lang === 'uz' ? 'uz' : 'ru';
    var course = stripQuotes(d.course);
    if (lang === 'uz') course = fixUz(course);
    var cls = classPhrase(d.cls, lang);
    var fullName = [normName(d.surname), normName(d.name)].filter(Boolean).join(' ');
    if (lang === 'uz') fullName = fixUz(fullName);
    var intro, segs;
    var B = function (t) { return { t: t, b: true }; }, R = function (t) { return { t: t, b: false }; };
    if (lang === 'uz') {
      intro = 'Ushbu sertifikat bilan tasdiqlanadiki, ' + (cls ? cls + ' o‘quvchisi' : 'o‘quvchi');
      segs = [
        [B('«' + course + '»'), R(' kursini muvaffaqiyatli tamomladi hamda tirishqoqlik, qiziquvchanlik va bilimga intilishni namoyon etdi.')]
      ];
    } else {
      var g = d.gender === 'm' || d.gender === 'f' ? d.gender : guessGender(d.surname, d.name);
      var who = g === 'f' ? 'ученица' : g === 'm' ? 'ученик' : 'ученик(ца)';
      intro = 'Настоящим подтверждается, что ' + who + (cls ? ' ' + cls : '');
      segs = [
        [R('за успешное окончание курса '), B('«' + course + '»'), R(',')],
        [R('проявленные старание, любознательность')],
        [R('и стремление к знаниям.')]
      ];
    }
    var sg = sign(lang);
    return { lang: lang, intro: intro, name: fullName, segs: segs, dirName: clean(sg.name), dirTitle: clean(sg.title) };
  }

  /* ---------- Загрузка ресурсов ---------- */
  var A = { fontBytes: {}, fk: {}, template: null, bg: null };
  var assetsReady = (async function () {
    var keys = Object.keys(FONT_FILES);
    var fetchBuf = function (u) { return fetch(u).then(function (r) { if (!r.ok) throw new Error(u); return r.arrayBuffer(); }); };
    var res = await Promise.all(keys.map(function (k) { return fetchBuf(FONT_FILES[k]); }).concat([
      fetchBuf('assets/template.pdf'),
      new Promise(function (ok, bad) { var im = new Image(); im.onload = function () { ok(im); }; im.onerror = bad; im.src = 'assets/bg.jpg'; })
    ]));
    keys.forEach(function (k, i) {
      A.fontBytes[k] = new Uint8Array(res[i]);
      A.fk[k] = window.fontkit.create(A.fontBytes[k]);
    });
    A.template = new Uint8Array(res[keys.length]);
    A.bg = res[keys.length + 1];
    await Promise.all(keys.map(function (k) {
      var ff = new FontFace(CANVAS_FAMILY[k], A.fontBytes[k].slice().buffer);
      document.fonts.add(ff); return ff.load();
    }));
  })();

  /* ---------- Вёрстка ---------- */
  var mCache = new Map();
  function measure(f, text, size) {
    var key = f + '|' + text;
    var w = mCache.get(key);
    if (w === undefined) {
      var font = A.fk[f];
      w = text ? font.layout(text).advanceWidth / font.unitsPerEm : 0;
      if (mCache.size > 5000) mCache.clear();
      mCache.set(key, w);
    }
    return w * size;
  }
  function fit(spec, text) {
    var s = spec.s;
    var w = measure(spec.f, text, s);
    if (w > spec.maxW) s = Math.max(spec.min || 8, Math.floor(s * spec.maxW / w * 10) / 10);
    return s;
  }
  function wrapSegment(runs, spec) {
    // слова = непрерывные куски без пробелов, могут состоять из жирных и обычных частей
    var words = [], cur = [];
    runs.forEach(function (r) {
      var parts = r.t.split(/( )/);
      parts.forEach(function (p) {
        if (p === ' ') { if (cur.length) { words.push(cur); cur = []; } }
        else if (p) cur.push({ t: p, b: r.b });
      });
    });
    if (cur.length) words.push(cur);
    var wWidth = function (w) { return w.reduce(function (a, p) { return a + measure(p.b ? spec.b : spec.f, p.t, spec.s); }, 0); };
    var space = measure(spec.f, ' ', spec.s);
    var lines = [], line = [], lw = 0;
    words.forEach(function (w) {
      var ww = wWidth(w);
      if (line.length && lw + space + ww > spec.maxW) { lines.push(line); line = [w]; lw = ww; }
      else { lw += (line.length ? space : 0) + ww; line.push(w); }
    });
    if (line.length) lines.push(line);
    // не оставляем одно слово на последней строке
    var n = lines.length;
    if (n > 1 && lines[n - 1].length === 1 && lines[n - 2].length > 2) {
      var moved = lines[n - 2].pop();
      var lastW = wWidth(moved) + space + wWidth(lines[n - 1][0]);
      if (lastW <= spec.maxW) lines[n - 1].unshift(moved); else lines[n - 2].push(moved);
    }
    // слова → куски строки, соседние куски одного начертания склеиваем
    return lines.map(function (ws) {
      var pieces = [];
      ws.forEach(function (w, i) {
        if (i) {
          var prev = ws[i - 1], bold = prev[prev.length - 1].b && w[0].b;
          pieces.push({ t: ' ', b: bold });
        }
        w.forEach(function (p) { pieces.push({ t: p.t, b: p.b }); });
      });
      var merged = [];
      pieces.forEach(function (p) {
        var last = merged[merged.length - 1];
        if (last && last.b === p.b) last.t += p.t; else merged.push({ t: p.t, b: p.b });
      });
      return merged;
    });
  }
  function layout(model) {
    var body = [];
    model.segs.forEach(function (seg) { body = body.concat(wrapSegment(seg, LAY.body)); });
    return {
      lang: model.lang,
      items: [
        { f: LAY.intro.f, s: fit(LAY.intro, model.intro), y: LAY.intro.y, c: LAY.intro.c, t: model.intro },
        { f: LAY.name.f, s: fit(LAY.name, model.name), y: LAY.name.y, c: LAY.name.c, t: model.name },
        { f: LAY.dir.f, s: fit(LAY.dir, model.dirName), y: LAY.dir.y, c: LAY.dir.c, t: model.dirName },
        { f: LAY.dirT.f, s: fit(LAY.dirT, model.dirTitle), y: LAY.dirT.y, c: LAY.dirT.c, t: model.dirTitle }
      ],
      body: body.map(function (pieces, i) { return { y: LAY.body.y + i * LAY.body.lh, pieces: pieces }; })
    };
  }

  /* ---------- Отрисовка: канвас (превью и JPG) ---------- */
  function cssColor(c) { return 'rgb(' + c[0] + ',' + c[1] + ',' + c[2] + ')'; }
  function drawCanvas(canvas, lay) {
    canvas.width = PAGE_W * PX; canvas.height = PAGE_H * PX;
    var ctx = canvas.getContext('2d');
    ctx.drawImage(A.bg, 0, 0, canvas.width, canvas.height);
    ctx.textBaseline = 'alphabetic';
    var setFont = function (f, s) { ctx.font = (s * PX) + 'px "' + CANVAS_FAMILY[f] + '"'; };
    lay.items.forEach(function (it) {
      if (!it.t) return;
      setFont(it.f, it.s); ctx.fillStyle = cssColor(it.c);
      ctx.fillText(it.t, X0 * PX, it.y * PX);
    });
    ctx.fillStyle = cssColor(LAY.body.c);
    lay.body.forEach(function (ln) {
      var x = X0 * PX;
      ln.pieces.forEach(function (p) {
        setFont(p.b ? LAY.body.b : LAY.body.f, LAY.body.s);
        ctx.fillText(p.t, x, ln.y * PX);
        x += ctx.measureText(p.t).width;
      });
    });
  }

  /* ---------- Отрисовка: PDF ---------- */
  async function newPdf(lang, title) {
    var PL = window.PDFLib;
    var doc = await PL.PDFDocument.create();
    doc.registerFontkit(window.fontkit);
    var emb = await doc.embedPdf(A.template);
    var fonts = {};
    for (var k in A.fontBytes) fonts[k] = await doc.embedFont(A.fontBytes[k], { subset: true });
    doc.setTitle(title || 'Certificate'); doc.setAuthor('Newton Academy'); doc.setCreator('Newton Academy — certificate generator');
    doc.setProducer('Newton Academy'); doc.setLanguage(lang === 'uz' ? 'uz-Latn' : 'ru-RU');
    return { doc: doc, tpl: emb[0], fonts: fonts };
  }
  function addPdfPage(P, lay) {
    var PL = window.PDFLib;
    var page = P.doc.addPage([PAGE_W, PAGE_H]);
    page.drawPage(P.tpl, { x: 0, y: 0, width: PAGE_W, height: PAGE_H });
    var col = function (c) { return PL.rgb(c[0] / 255, c[1] / 255, c[2] / 255); };
    lay.items.forEach(function (it) {
      if (!it.t) return;
      page.drawText(it.t, { x: X0, y: PAGE_H - it.y, size: it.s, font: P.fonts[it.f], color: col(it.c) });
    });
    lay.body.forEach(function (ln) {
      var x = X0;
      ln.pieces.forEach(function (p) {
        var f = P.fonts[p.b ? LAY.body.b : LAY.body.f];
        page.drawText(p.t, { x: x, y: PAGE_H - ln.y, size: LAY.body.s, font: f, color: col(LAY.body.c) });
        x += f.widthOfTextAtSize(p.t, LAY.body.s);
      });
    });
  }
  async function pdfBlob(P) {
    var bytes = await P.doc.save({ useObjectStreams: true });
    return new Blob([bytes], { type: 'application/pdf' });
  }
  function canvasBlob(canvas) {
    return new Promise(function (ok) { canvas.toBlob(ok, 'image/jpeg', 0.92); });
  }

  /* ---------- Подпись (настройки) ---------- */
  function getSign() {
    var raw = store('na-cert-sign'), o = {};
    try { o = raw ? JSON.parse(raw) : {}; } catch (e) { o = {}; }
    var r = {};
    Object.keys(DEFAULT_SIGN).forEach(function (k) { r[k] = clean(o[k]) || DEFAULT_SIGN[k]; });
    return r;
  }

  /* ---------- Интерфейс ---------- */
  var uiLang = store('na-cert-ui') === 'uz' ? 'uz' : 'ru';
  function t(k) { return (I18N[uiLang] && I18N[uiLang][k]) || I18N.ru[k] || k; }
  function applyI18n() {
    document.documentElement.lang = uiLang === 'uz' ? 'uz' : 'ru';
    $$('[data-i18n]').forEach(function (el) { el.textContent = t(el.getAttribute('data-i18n')); });
    $$('[data-i18n-ph]').forEach(function (el) { el.placeholder = t(el.getAttribute('data-i18n-ph')); });
    $$('[data-ui-lang]').forEach(function (b) { b.setAttribute('aria-pressed', String(b.getAttribute('data-ui-lang') === uiLang)); });
    updateSignLine(); renderTable(); renderCourses();
  }
  function seg(groupSel) {
    var g = $(groupSel);
    return {
      get: function () { var b = $('button[aria-pressed="true"]', g); return b ? b.getAttribute('data-v') : 'ru'; },
      set: function (v) { $$('button', g).forEach(function (b) { b.setAttribute('aria-pressed', String(b.getAttribute('data-v') === v)); }); },
      on: function (fn) { g.addEventListener('click', function (e) { var b = e.target.closest('button[data-v]'); if (!b) return; this.set(b.getAttribute('data-v')); fn(b.getAttribute('data-v')); }.bind(this)); }
    };
  }
  function toast(msg, kind) {
    var el = $('#toast'); el.textContent = msg; el.className = 'toast show ' + (kind || '');
    clearTimeout(toast._t); toast._t = setTimeout(function () { el.className = 'toast'; }, 3800);
  }

  /* ---- Мои курсы: свой список, хранится в этом браузере ---- */
  var COURSES_KEY = 'na-cert-courses';
  function getCourses() {
    var raw = store(COURSES_KEY), list = null;
    try { list = raw ? JSON.parse(raw) : null; } catch (e) { list = null; }
    if (!Array.isArray(list)) list = ['Почемучка'];
    return list.map(clean).filter(Boolean);
  }
  function setCourses(list) {
    var seen = {}, out = [];
    list.forEach(function (c) { c = stripQuotes(c); var k = c.toLowerCase(); if (c && !seen[k]) { seen[k] = 1; out.push(c); } });
    store(COURSES_KEY, JSON.stringify(out.slice(0, 40)));
    renderCourses();
  }
  function rememberCourse(name) {
    var c = stripQuotes(name); if (!c) return;
    var list = getCourses();
    if (!list.some(function (v) { return v.toLowerCase() === c.toLowerCase(); })) setCourses(list.concat([c]));
  }
  function renderCourses() {
    var list = getCourses();
    $$('.courses[data-input]').forEach(function (box) {
      var input = $(box.getAttribute('data-input'));
      var cur = stripQuotes(input.value).toLowerCase();
      var has = list.some(function (c) { return c.toLowerCase() === cur; });
      var html = list.length ? '<span class="lbl">' + esc(t('myCourses')) + '</span>' : '';
      html += list.map(function (c) {
        return '<span class="chip' + (c.toLowerCase() === cur ? ' on' : '') + '" role="button" tabindex="0" data-c="' + esc(c) + '">' +
          '<span class="t">' + esc(c) + '</span>' +
          '<span class="x" role="button" tabindex="0" title="' + esc(t('removeCourse')) + '" aria-label="' + esc(t('removeCourse') + ': ' + c) + '">×</span></span>';
      }).join('');
      if (cur && !has) html += '<button type="button" class="chip add" data-add="1">＋ ' + esc(t('saveCourse')) + '</button>';
      box.innerHTML = html;
    });
  }
  function onCourseChip(e) {
    var box = e.target.closest('.courses[data-input]'); if (!box) return;
    var input = $(box.getAttribute('data-input'));
    var chip = e.target.closest('.chip'); if (!chip) return;
    e.preventDefault();
    if (e.target.closest('.x')) {
      var c = chip.getAttribute('data-c');
      setCourses(getCourses().filter(function (v) { return v !== c; }));
      return;
    }
    if (chip.hasAttribute('data-add')) { rememberCourse(input.value); toast(t('courseSaved'), 'ok'); return; }
    input.value = chip.getAttribute('data-c');
    input.dispatchEvent(new Event('input', { bubbles: true }));
  }

  /* ---- Один сертификат ---- */
  var singleLang = seg('#single-lang');
  singleLang.set(store('na-cert-lang') === 'uz' ? 'uz' : 'ru');
  var genderManual = false;
  function getGender() { var b = $('#single-gender button[aria-pressed="true"]'); return b ? b.getAttribute('data-v') : ''; }
  function setGender(v) { $$('#single-gender button').forEach(function (b) { b.setAttribute('aria-pressed', String(b.getAttribute('data-v') === v)); }); }
  function autoGender() { if (!genderManual) setGender(guessGender($('#f-surname').value, $('#f-name').value)); }
  function singleData() {
    return { surname: $('#f-surname').value, name: $('#f-name').value, cls: $('#f-cls').value, gender: getGender(), course: $('#f-course').value, lang: singleLang.get() };
  }
  var previewTimer;
  function schedulePreview() { clearTimeout(previewTimer); previewTimer = setTimeout(renderSinglePreview, 60); }
  async function renderSinglePreview() {
    await assetsReady;
    var d = singleData();
    var m = buildModel({ surname: d.surname || t('phSurname'), name: d.name || (d.surname ? '' : t('phName')), cls: d.cls, gender: d.gender, course: d.course || t('demoCourse'), lang: d.lang });
    drawCanvas($('#single-canvas'), layout(m));
    $('#single-canvas').classList.toggle('is-demo', !(clean(d.surname) || clean(d.name)));
  }
  function singleValid() {
    var d = singleData();
    var ok = (clean(d.surname) || clean(d.name)) && clean(d.course);
    if (!ok) {
      toast(t('needFields'), 'err');
      [['#f-surname', clean(d.surname) || clean(d.name)], ['#f-name', clean(d.surname) || clean(d.name)], ['#f-course', clean(d.course)]].forEach(function (p) {
        $(p[0]).classList.toggle('invalid', !p[1]);
      });
      return null;
    }
    $$('.invalid').forEach(function (el) { el.classList.remove('invalid'); });
    rememberCourse(d.course);
    return d;
  }
  function fileBase(d) {
    return safeFile([normName(d.surname), normName(d.name)].filter(Boolean).join(' '));
  }
  async function busy(btn, fn) {
    if (btn.disabled) return;
    btn.disabled = true; btn.classList.add('loading');
    try { await fn(); } catch (e) { console.error(e); toast(String(e && e.message || e), 'err'); }
    finally { btn.disabled = false; btn.classList.remove('loading'); }
  }

  /* ---- Массовая генерация ---- */
  var bulkLang = seg('#bulk-lang');
  bulkLang.set(store('na-cert-lang') === 'uz' ? 'uz' : 'ru');
  var rows = [], selectedRow = -1, cancelFlag = false;

  var HEAD = {
    surname: ['фамилия', 'familiya', 'familya', 'surname', 'lastname', 'фамилияученика', 'фамилиясы'],
    name: ['имя', 'ism', 'ismi', 'name', 'firstname', 'имяученика', 'исм'],
    fio: ['фио', 'фиоученика', 'фиоребенка', 'fio', 'fish', 'ismfamiliya', 'familiyaism', 'fullname', 'ученик', 'oquvchi', 'фамилияимя', 'имяфамилия', 'фиш'],
    cls: ['класс', 'sinf', 'class', 'grade', 'синф', 'sinfi'],
    course: ['курс', 'kurs', 'course', 'программа', 'dastur', 'направление', 'kursnomi', 'названиекурса'],
    lang: ['язык', 'til', 'lang', 'language', 'tili', 'языксертификата', 'sertifikattili'],
    gender: ['пол', 'полмж', 'полученика', 'jins', 'jinsi', 'gender', 'sex', 'жинси', 'жинс']
  };
  function normHead(v) { return String(v == null ? '' : v).toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-яўқғҳ0-9]+/g, ''); }
  function headKey(v) {
    var h = normHead(v); if (!h) return null;
    if (/^(полное|полнфио)/.test(h)) return 'fio';
    for (var k in HEAD) { if (HEAD[k].some(function (x) { return h === x || h.indexOf(x) === 0; })) return k; }
    return null;
  }
  function cellStr(v) {
    if (v == null) return '';
    if (v instanceof Date) return v.toLocaleDateString('ru-RU');
    return clean(v);
  }
  function langOf(v) {
    var s = normHead(v);
    if (!s) return '';
    if (/^(uz|узб|ўзб|ozb|uzb|узбек|ozbek)/.test(s)) return 'uz';
    if (/^(ru|рус|rus)/.test(s)) return 'ru';
    return '';
  }
  function mapRows(table) {
    table = table.filter(function (r) { return r && r.some(function (c) { return cellStr(c) !== ''; }); });
    if (!table.length) return [];
    var map = null, start = 0;
    for (var i = 0; i < Math.min(5, table.length); i++) {
      var m = {}, n = 0;
      table[i].forEach(function (c, j) { var k = headKey(c); if (k && m[k] === undefined) { m[k] = j; n++; } });
      if (n >= 2 || m.fio !== undefined) { map = m; start = i + 1; break; }
    }
    if (!map) map = { surname: 0, name: 1, cls: 2, gender: 3, course: 4, lang: 5 };
    var g = function (r, k) { return map[k] === undefined ? '' : cellStr(r[map[k]]); };
    return table.slice(start).map(function (r) {
      var surname = g(r, 'surname'), name = g(r, 'name');
      if (!surname && !name && map.fio !== undefined) {
        var parts = g(r, 'fio').split(' ').filter(Boolean);
        surname = parts.shift() || ''; name = parts.join(' ');
      }
      return { surname: surname, name: name, cls: g(r, 'cls'), course: g(r, 'course'), lang: langOf(g(r, 'lang')), gender: genderOf(g(r, 'gender')) };
    });
  }
  function parseCSV(text) {
    text = text.replace(/^﻿/, '');
    var first = text.split(/\r?\n/)[0] || '';
    var delim = (first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ';' : ((first.match(/\t/g) || []).length ? '\t' : ',');
    var out = [], row = [], cell = '', q = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === delim) { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); out.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell || row.length) { row.push(cell); out.push(row); }
    return out;
  }
  async function readTable(file) {
    var nm = file.name.toLowerCase();
    if (/\.xls$/.test(nm)) throw new Error(t('xlsOld'));
    if (/\.(csv|txt)$/.test(nm) || file.type === 'text/csv') {
      var buf = await file.arrayBuffer();
      var text = new TextDecoder('utf-8').decode(buf);
      if (text.indexOf('�') !== -1) { try { text = new TextDecoder('windows-1251').decode(buf); } catch (e) { /* оставляем utf-8 */ } }
      return parseCSV(text);
    }
    var res = await window.readXlsxFile(file);
    if (Array.isArray(res) && res.length && res[0] && Array.isArray(res[0].data)) {
      var sheet = res.find(function (s) { return s.data && s.data.some(function (r) { return r.some(function (c) { return c != null && c !== ''; }); }); });
      return sheet ? sheet.data : [];
    }
    return res;
  }
  function rowProblems(r) {
    var p = [];
    if (!clean(r.surname) && !clean(r.name)) p.push(t('errName'));
    if (!clean(r.course) && !clean($('#def-course').value)) p.push(t('errCourse'));
    return p;
  }
  function rowData(r) {
    return { surname: r.surname, name: r.name, cls: r.cls, gender: r.gender || '', course: clean(r.course) || clean($('#def-course').value), lang: r.lang || bulkLang.get() };
  }
  function genderPill(r) {
    var g = r.gender || guessGender(r.surname, r.name);
    if (!g) return '<button type="button" class="gpill none" title="' + esc(t('gToggle')) + '">—</button>';
    return '<button type="button" class="gpill ' + g + (r.gender ? '' : ' auto') + '" title="' + esc(t('gToggle') + (r.gender ? '' : ' · ' + t('gAuto'))) + '">' +
      esc(g === 'f' ? t('gF') : t('gM')) + '</button>';
  }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function renderTable() {
    var box = $('#bulk-result');
    if (!rows.length) { box.hidden = true; return; }
    box.hidden = false;
    var ok = 0, bad = 0;
    var html = rows.map(function (r, i) {
      var p = rowProblems(r); if (p.length) bad++; else ok++;
      var lang = r.lang || bulkLang.get();
      return '<tr data-i="' + i + '" class="' + (p.length ? 'bad' : '') + (i === selectedRow ? ' sel' : '') + '">' +
        '<td class="num">' + (i + 1) + '</td>' +
        '<td>' + esc(r.surname) + '</td><td>' + esc(r.name) + '</td><td>' + esc(r.cls) + '</td>' +
        '<td>' + genderPill(r) + '</td>' +
        '<td>' + (r.course ? esc(r.course) : '<span class="muted">' + esc(clean($('#def-course').value) || '—') + '</span>') + '</td>' +
        '<td><span class="pill">' + lang.toUpperCase() + '</span></td>' +
        '<td class="st">' + (p.length ? '<span class="errtxt">' + esc(p.join(', ')) + '</span>' : '<span class="oktxt">✓</span>') + '</td>' +
        '<td><button class="rm" type="button" title="×" aria-label="remove">×</button></td></tr>';
    }).join('');
    $('#bulk-tbody').innerHTML = html;
    $('#bulk-count').innerHTML = '<b>' + ok + '</b> ' + esc(t('rowsReady')) + (bad ? ' · <span class="errtxt"><b>' + bad + '</b> ' + esc(t('rowsErr')) + '</span>' : '');
    $$('#bulk-actions button').forEach(function (b) { b.disabled = ok === 0; });
  }
  async function previewRow(i) {
    await assetsReady;
    var r = rows[i]; if (!r) return;
    selectedRow = i;
    $$('#bulk-tbody tr').forEach(function (tr) { tr.classList.toggle('sel', Number(tr.getAttribute('data-i')) === i); });
    var d = rowData(r);
    drawCanvas($('#bulk-canvas'), layout(buildModel({ surname: d.surname, name: d.name, cls: d.cls, gender: d.gender, course: d.course || '…', lang: d.lang })));
    $('#bulk-preview').hidden = false;
  }
  async function loadFile(file) {
    try {
      var table = await readTable(file);
      var mapped = mapRows(table);
      if (!mapped.length) { toast(t('noRows'), 'err'); return; }
      rows = mapped; selectedRow = -1;
      $('#file-name').textContent = file.name;
      $('#drop').classList.add('has-file');
      renderTable();
      var first = rows.findIndex(function (r) { return !rowProblems(r).length; });
      await previewRow(first >= 0 ? first : 0);
      $('#bulk-result').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
      console.error(e); toast(e && e.message && e.message.indexOf('.xls') !== -1 ? e.message : t('readErr'), 'err');
    }
  }
  function setProgress(n, total, text) {
    var box = $('#progress'); box.hidden = false;
    $('#progress-bar').style.width = (total ? Math.round(n / total * 100) : 0) + '%';
    $('#progress-text').textContent = text || (t('progress') + ' ' + n + ' ' + t('of') + ' ' + total);
  }
  function uniqueName(used, base, ext) {
    var name = base + ext, k = 2;
    while (used[name.toLowerCase()]) name = base + ' (' + k++ + ')' + ext;
    used[name.toLowerCase()] = 1; return name;
  }
  async function runBulk(kind, btn) {
    await assetsReady;
    var list = rows.filter(function (r) { return !rowProblems(r).length; }).map(rowData);
    if (!list.length) return;
    cancelFlag = false;
    $$('#bulk-actions button').forEach(function (b) { b.disabled = true; });
    btn.classList.add('loading'); $('#cancel').hidden = false;
    var course = clean($('#def-course').value) || clean(list[0].course);
    var stamp = new Date().toISOString().slice(0, 10);
    try {
      if (kind === 'one') {
        var P = await newPdf(list[0].lang, t('certs') + ' — ' + course);
        for (var i = 0; i < list.length; i++) {
          if (cancelFlag) throw 'cancel';
          addPdfPage(P, layout(buildModel(list[i])));
          setProgress(i + 1, list.length); if (i % 5 === 4) await tick();
        }
        setProgress(list.length, list.length, '…');
        await tick();
        download(await pdfBlob(P), safeFile(t('certs') + ' ' + course + ' ' + stamp) + ' (' + list.length + ').pdf');
      } else {
        var zip = new window.JSZip(), used = {};
        var cv = document.createElement('canvas');
        for (var j = 0; j < list.length; j++) {
          if (cancelFlag) throw 'cancel';
          var d = list[j], lay = layout(buildModel(d)), base = fileBase(d);
          if (kind === 'pdf') {
            var Pj = await newPdf(d.lang, base);
            addPdfPage(Pj, lay);
            zip.file(uniqueName(used, base, '.pdf'), await Pj.doc.save({ useObjectStreams: true }));
          } else {
            drawCanvas(cv, lay);
            zip.file(uniqueName(used, base, '.jpg'), await canvasBlob(cv));
          }
          setProgress(j + 1, list.length); await tick();
        }
        setProgress(list.length, list.length, '…');
        var blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
        download(blob, safeFile(t('certs') + ' ' + course + ' ' + stamp) + (kind === 'jpg' ? ' JPG' : '') + '.zip');
      }
      setProgress(list.length, list.length, t('done')); toast(t('done'), 'ok');
      var used = {};
      list.forEach(function (d) { var c = stripQuotes(d.course); if (c) used[c.toLowerCase()] = c; });
      var uniq = Object.keys(used).map(function (k) { return used[k]; });
      if (uniq.length <= 8) uniq.forEach(rememberCourse);
    } catch (e) {
      if (e === 'cancel') { setProgress(0, 1, t('canceled')); }
      else { console.error(e); toast(String(e && e.message || e), 'err'); }
    } finally {
      btn.classList.remove('loading'); $('#cancel').hidden = true;
      renderTable();
    }
  }

  /* ---- Подпись: строка и диалог ---- */
  function updateSignLine() {
    var el = $('#sign-line'); if (!el) return;
    var s = sign(singleLang.get());
    el.textContent = s.name + ', ' + s.title;
  }
  function openSettings() {
    var s = getSign();
    $('#s-ru-name').value = s.ruName; $('#s-ru-title').value = s.ruTitle;
    $('#s-uz-name').value = s.uzName; $('#s-uz-title').value = s.uzTitle;
    $('#settings').showModal();
  }

  /* ---------- Привязка событий ---------- */
  document.addEventListener('DOMContentLoaded', function () {
    applyI18n();
    $('#status').textContent = t('loading');
    assetsReady.then(function () {
      $('#status').textContent = ''; document.body.classList.add('ready');
      renderSinglePreview();
    }).catch(function (e) { console.error(e); $('#status').textContent = 'Error: ' + e.message; });

    $$('[data-ui-lang]').forEach(function (b) {
      b.addEventListener('click', function () { uiLang = b.getAttribute('data-ui-lang'); store('na-cert-ui', uiLang); applyI18n(); renderSinglePreview(); });
    });
    $$('.tab').forEach(function (tb) {
      tb.addEventListener('click', function () {
        $$('.tab').forEach(function (x) { x.setAttribute('aria-selected', String(x === tb)); });
        $$('.panel').forEach(function (p) { p.hidden = p.id !== tb.getAttribute('aria-controls'); });
        store('na-cert-tab', tb.id);
      });
    });
    var savedTab = store('na-cert-tab'); if (savedTab && $('#' + savedTab)) $('#' + savedTab).click();

    ['#f-surname', '#f-name', '#f-cls', '#f-course'].forEach(function (s) {
      $(s).addEventListener('input', function () { this.classList.remove('invalid'); schedulePreview(); });
    });
    ['#f-surname', '#f-name'].forEach(function (s) { $(s).addEventListener('input', autoGender); });
    $('#f-course').addEventListener('input', renderCourses);
    $('#single-gender').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-v]'); if (!b) return;
      genderManual = true; setGender(b.getAttribute('data-v')); schedulePreview();
    });
    document.addEventListener('click', onCourseChip);
    document.addEventListener('keydown', function (e) {
      if ((e.key === 'Enter' || e.key === ' ') && e.target.closest && e.target.closest('.courses .chip, .courses .x') && e.target.tagName !== 'BUTTON') onCourseChip(e);
    });
    renderCourses();
    singleLang.on(function (v) { store('na-cert-lang', v); updateSignLine(); schedulePreview(); });
    $('#single-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var btn = $('#btn-pdf');
      busy(btn, async function () {
        var d = singleValid(); if (!d) return;
        await assetsReady;
        var lay = layout(buildModel(d));
        var P = await newPdf(d.lang, fileBase(d));
        addPdfPage(P, lay);
        download(await pdfBlob(P), fileBase(d) + '.pdf');
        toast(t('done'), 'ok');
      });
    });
    $('#btn-jpg').addEventListener('click', function () {
      var btn = this;
      busy(btn, async function () {
        var d = singleValid(); if (!d) return;
        await assetsReady;
        var cv = document.createElement('canvas');
        drawCanvas(cv, layout(buildModel(d)));
        download(await canvasBlob(cv), fileBase(d) + '.jpg');
      });
    });
    $('#btn-print').addEventListener('click', function () {
      var btn = this;
      busy(btn, async function () {
        var d = singleValid(); if (!d) return;
        await assetsReady;
        var P = await newPdf(d.lang, fileBase(d));
        addPdfPage(P, layout(buildModel(d)));
        var url = URL.createObjectURL(await pdfBlob(P));
        var w = window.open(url, '_blank');
        if (!w) download(await pdfBlob(P), fileBase(d) + '.pdf');
      });
    });

    $('#sign-edit').addEventListener('click', openSettings);
    $('#settings-form').addEventListener('submit', function (e) {
      e.preventDefault();
      store('na-cert-sign', JSON.stringify({
        ruName: $('#s-ru-name').value, ruTitle: $('#s-ru-title').value,
        uzName: $('#s-uz-name').value, uzTitle: $('#s-uz-title').value
      }));
      $('#settings').close(); updateSignLine(); renderSinglePreview(); if (selectedRow >= 0) previewRow(selectedRow);
    });
    $('#s-reset').addEventListener('click', function () {
      $('#s-ru-name').value = DEFAULT_SIGN.ruName; $('#s-ru-title').value = DEFAULT_SIGN.ruTitle;
      $('#s-uz-name').value = DEFAULT_SIGN.uzName; $('#s-uz-title').value = DEFAULT_SIGN.uzTitle;
    });
    $('#s-close').addEventListener('click', function () { $('#settings').close(); });

    var drop = $('#drop'), input = $('#file');
    drop.addEventListener('click', function () { input.click(); });
    drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', function () { if (input.files[0]) loadFile(input.files[0]); input.value = ''; });
    ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
    drop.addEventListener('drop', function (e) { var f = e.dataTransfer.files[0]; if (f) loadFile(f); });

    $('#def-course').value = store('na-cert-defcourse') || '';
    renderCourses();
    $('#def-course').addEventListener('input', function () { store('na-cert-defcourse', this.value); renderCourses(); renderTable(); if (selectedRow >= 0) previewRow(selectedRow); });
    bulkLang.on(function (v) { store('na-cert-lang', v); renderTable(); if (selectedRow >= 0) previewRow(selectedRow); });
    $('#bulk-tbody').addEventListener('click', function (e) {
      var tr = e.target.closest('tr'); if (!tr) return;
      var i = Number(tr.getAttribute('data-i'));
      if (e.target.closest('.rm')) {
        rows.splice(i, 1);
        if (selectedRow === i) selectedRow = -1; else if (selectedRow > i) selectedRow--;
        renderTable();
        if (!rows.length) { $('#bulk-preview').hidden = true; $('#drop').classList.remove('has-file'); }
        else if (selectedRow >= 0) previewRow(selectedRow);
        return;
      }
      if (e.target.closest('.gpill')) {
        var r = rows[i], g = r.gender || guessGender(r.surname, r.name);
        r.gender = g === 'm' ? 'f' : 'm';
        renderTable();
      }
      previewRow(i);
    });
    $('#btn-zip-pdf').addEventListener('click', function () { runBulk('pdf', this); });
    $('#btn-one-pdf').addEventListener('click', function () { runBulk('one', this); });
    $('#btn-zip-jpg').addEventListener('click', function () { runBulk('jpg', this); });
    $('#cancel').addEventListener('click', function () { cancelFlag = true; });
  });

  // для автотестов
  window.NewtonCert = { buildModel: buildModel, layout: layout, ready: assetsReady, classPhrase: classPhrase, mapRows: mapRows, normName: normName, guessGender: guessGender, genderOf: genderOf, getCourses: getCourses };
})();
