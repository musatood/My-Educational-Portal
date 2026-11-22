/* =====================================================================
   خادم الحقيبة التدريبية — ملف واحد كامل (Code.gs)
   يخزّن: النتائج، الإشعارات، قائمة المسجلين، ورموز الدخول
   ===================================================================== */

/* ======= عدّل هذه الأسطر فقط ======= */
const PIN   = 'musa2026';   // رمز المدرب: حروف إنجليزية وأرقام بلا مسافات، ولا تستعمل رمزاً ظهر في لقطة
const SS_ID = '1Am_4ik80BIdBsGW8psqVBZm9Dk06ndTJnfB2zNV3vnc';   // رقم جدول Google Sheets: الجزء الطويل في رابطه بين /d/ و /edit
const AI_ATTEMPTS     = 3;            // محاولات الذكاء الاصطناعي لكل متدرب في اليوم
const AI_DAILY_RESET  = true;         // true: تتجدد محاولات كل متدرب تلقائياً كل يوم. false: لا تتجدد إلا بتجديد المدرب
const AI_GLOBAL_DAILY = 75;           // السقف اليومي لكل المتدربين معاً (راجع حد حساب OpenRouter)
const AI_TZ           = 'Africa/Khartoum'; // المنطقة الزمنية التي يبدأ فيها اليوم الجديد
const AI_MAX_SET      = 50;           // أقصى رقم يستطيع المدرب تعيينه لمتدرب من لوحة المدرب
/* ===================================== */

const RES_HEADERS = ['name','wa','email','score','total','start','end','mins','right','wrong','rows','phase','cert','oral','flags','sig','calib'];
const CERT_NEEDS_VIDEO   = true;   // الشهادة تتطلب إرسال فيديو المقابلة المسجلة (دقيقة)
const CERT_NEEDS_ORAL    = false;  // الشهادة تتطلب أن يؤكد المدرب إجراء المقابلة الشفوية
const CERT_NEEDS_PRE      = true;   // الشهادة تتطلب إجراء الاختبار القبلي
const CERT_NEEDS_PROJECT  = true;   // الشهادة تتطلب اعتماد المدرب للمشروع التطبيقي
const CERT_NEEDS_VERIFY   = true;   // الشهادة تتطلب تقرير المدرب على الفيديو (التحقق أن المتدرب هو من أدّى الاختبار)
const ADMIN_PIN = 'ous2027'; 
const VIDEO_MAX_PER_USER = 3;      // أقصى عدد فيديوهات لكل متدرب
const VIDEO_FOLDER       = 'ODL_Videos'; // مجلد Google Drive الذي تُحفظ فيه الفيديوهات
const VID_HEADERS = ['ph','name','start','fileId','url','secs','mime','prompt','ts','fb','rating','fbts','ver','vchk','vnote','vts'];   // fb/rating/fbts: ملاحظات المدرب وتقديره وزمنها (تصل المتدرب)
const PROJ_HEADERS = ['ph','name','title','link','checks','score','note','ts','feedback','fileUrl','fileId','fileName','status','rubric','fbts'];   // status/rubric/fbts: حالة المشروع وتقييم المدرب لكل معيار وزمن آخر تحديث
const PROJ_FOLDER = 'ODL_Projects';          // مجلد Google Drive الذي تُحفظ فيه ملفات مشاريع المتدربين
const PROJ_FILE_MAX = 5*1024*1024;           // أقصى حجم للملف: 5 ميغابايت
const PROJ_EXT = {pdf:'application/pdf', doc:'application/msword', docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document', ppt:'application/vnd.ms-powerpoint', pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation', xls:'application/vnd.ms-excel', xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', txt:'text/plain', rtf:'application/rtf', odt:'application/vnd.oasis.opendocument.text'};
const PASS_PCT = 60;                   // نسبة النجاح في الاختبار البعدي لإصدار الشهادة

function J_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_(name, headers){
  const ss = (SS_ID && SS_ID.indexOf('ضع') !== 0) ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if(s && s.getLastColumn() < headers.length){                 // ورقة قديمة: أضف الأعمدة الجديدة
    const n0 = s.getLastColumn();
    s.getRange(1,1,1,headers.length).setValues([headers]);
    s.getRange(1,n0+1,s.getMaxRows(),headers.length-n0).setNumberFormat('@');
  }
  if(!s){
    s = ss.insertSheet(name);
    s.getRange(1,1,1,headers.length).setValues([headers]);
    s.getRange(1,1,s.getMaxRows(),headers.length).setNumberFormat('@');
    s.setFrozenRows(1);
  }
  return s;
}
function readRows_(){
  const s = sheet_('Results', RES_HEADERS), v = s.getDataRange().getValues();
  const h = v[0], out = [];
  for(let i=1;i<v.length;i++){
    if(!v[i][0] && !v[i][1]) continue;
    const o = {}; h.forEach((k,j)=>{ o[k] = String(v[i][j]); });
    {   // إصلاح الدرجة إن حوّلها Sheets إلى تاريخ/نص غير رقمي: تُحسب من نقاط كل سؤال المحفوظة في rows
      const t = String(o.score||'').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[٫,]/g, '.').trim();
      if(t !== '' && isFinite(+t)) o.score = String(+t);
      else { try{ const rr = JSON.parse(o.rows); if(Array.isArray(rr) && rr.length) o.score = String(Math.max(0, Math.round(rr.reduce((s,x) => s + (+x.pt||0), 0)*100)/100)); }catch(e){} }
    }
    out.push(o);
  }
  return out;
}
function projRows_(){
  const s = sheet_('Projects', PROJ_HEADERS), v = s.getDataRange().getValues(), out = [];
  for(let i=1;i<v.length;i++){
    const o = {}; PROJ_HEADERS.forEach((k,j)=>{ o[k] = String(v[i][j]===undefined?'':v[i][j]); });
    // علّق المدرب أو قيّم بعد آخر تسليم ولم يختر قراراً: فالمشروع «قيد المراجعة» لا «بانتظار المراجعة»
    if((!o.status || o.status === 'submitted' || o.status === 'resubmitted') && +o.fbts > 0 && +o.fbts >= (+o.ts||0) && (o.feedback.trim() || /[012]/.test(o.rubric))) o.status = 'reviewing';
    out.push(o);
  }
  return out;
}
function vidRows_(){
  const s = sheet_('Videos', VID_HEADERS), v = s.getDataRange().getValues(), out = [];
  for(let i=1;i<v.length;i++){ const o = {}; VID_HEADERS.forEach((k,j)=>{ o[k] = String(v[i][j]===undefined?'':v[i][j]); }); out.push(o); }
  return out;
}
/* ===================== التقييم بنسبة الثقة + التصحيح في الخادم ===================== */
/* الدرجة تُحسب هنا فقط؛ ما يرسله المتصفح من درجة لا يُعتمد. عدّل الجدولين لتغيير شدة النظام (90 = «80% فأكثر») */
const CONF_CORRECT = {25:0.4, 50:0.5, 75:0.75, 90:1};     // نقاط الإجابة الصحيحة حسب الثقة
const CONF_WRONG   = {25:0,   50:0.1, 75:0.5,  90:1.5};   // ما يُخصم عند الإجابة الخاطئة حسب الثقة
const CONF_TF = [50,75,90], CONF_MC = [25,50,75,90];      // مستويات الثقة المسموحة (صح/خطأ يبدأ من 50)
const EXAM_COUNTS = {tf:2, mc:15, fl:3};                   // عدد أسئلة كل نوع في الاختبار الواحد
/* مفتاح الإجابات بترتيب رقم السؤال qid (صح/خطأ ثم اختيار ثم فراغ). إن عدّلتَ أسئلة الصفحة (TF وMC وFL) فعدّل هذا المفتاح بالترتيب نفسه */
const QKEY_OLD = [
 {"t": "tf", "v": true, "e": "صح: الانفصال المكاني من أبرز سمات التعليم عن بُعد."},
 {"t": "tf", "v": true, "e": "صح: كانت المراسلة من أولى صوره."},
 {"t": "tf", "v": true, "e": "صح: كالجلسات المباشرة."},
 {"t": "tf", "v": true, "e": "صح."},
 {"t": "tf", "v": true, "e": "صح: BY تعني نسب المصنف لصاحبه."},
 {"t": "tf", "v": false, "e": "خطأ: قد يُستخدم التعليم الإلكتروني داخل القاعة التقليدية."},
 {"t": "mc", "o": ["مايكل مور", "تشارلز ويدماير", "بورج هولمبرغ", "بنجامين بلوم"], "c": 0},
 {"t": "mc", "o": ["نسب المصنف", "غير تجاري", "المشاركة بالمثل", "منع الاشتقاق"], "c": 1},
 {"t": "mc", "o": ["حضور الجميع في موعد واحد", "بث مباشر إلزامي", "الدراسة في الوقت المناسب للمتعلم", "غياب أي وسيلة للتواصل"], "c": 2},
 {"t": "mc", "o": ["الإذاعة والتلفاز", "الهواتف الذكية", "الذكاء الاصطناعي", "المواد المطبوعة والبريد"], "c": 3},
 {"t": "mc", "o": ["مزج مخطط بين التعلم الحضوري والتعلم عبر الإنترنت", "بث مباشر فقط عبر الإنترنت", "تسجيل المحاضرات وحدها", "التعلم الذاتي بلا لقاءات"], "c": 0},
 {"t": "mc", "o": ["عنوان عام للموضوعات", "عبارة قابلة للقياس تبدأ بفعل سلوكي", "قائمة بالكتب المقررة", "وصف لما سيفعله المحاضر"], "c": 1},
 {"t": "mc", "o": ["عقاب الطالب", "ملء وقت المحاضرة", "إعلام الطالب بمستواه وكيف يتحسن", "زيادة عدد الواجبات"], "c": 2},
 {"t": "mc", "o": ["لا فرق بينهما", "المجاني أكثر حرية من المفتوح", "المفتوح هو ما ليس له مؤلف", "المفتوح يُتاح استخدامه وتعديله وإعادة نشره وفق ترخيص مفتوح، والمجاني قد لا يتيح ذلك"], "c": 3},
 {"t": "mc", "o": ["ND يمنع تعديل العمل، وNC يمنع استخدامه تجارياً", "ND يمنع الاستخدام التجاري، وNC يمنع التعديل", "كلاهما يمنع ذكر المؤلف", "كلاهما يسمح بكل أنواع الاستخدام"], "c": 0},
 {"t": "mc", "o": ["CC BY-ND", "CC BY-SA", "CC BY-NC-ND", "عمل بلا ترخيص"], "c": 1},
 {"t": "mc", "o": ["هما تسميتان لشيء واحد", "المدمج يعني بثاً مباشراً دائماً", "المدمج مزج مخطط بين الحضوري والرقمي، والمتزامن وصف لتوقيت التفاعل", "المتزامن يعني الحضور في القاعة فقط"], "c": 2},
 {"t": "mc", "o": ["ربط الدراسة بحضور المحاضرات", "إلغاء دور المعلم كلياً", "اعتماد امتحان مركزي واحد", "تحرير المتعلم من قيود الزمان والمكان وتمكينه من تنظيم دراسته"], "c": 3},
 {"t": "mc", "o": ["العلاقة الشخصية الودية بين المعلم والمتعلم تدعم الدافعية والتعلم", "الفصل التام بين المعلم والمتعلم", "التركيز على الوسيلة التقنية وحدها", "قياس التعلم بالمدة التي يقضيها الطالب"], "c": 0},
 {"t": "mc", "o": ["يُلقى المحتوى كله أثناء اللقاء", "يدرس الطالب المحتوى قبل اللقاء ويُخصص اللقاء للتطبيق والنقاش", "يُلغى اللقاء نهائياً", "يدرس الطالب المحتوى بعد الاختبار"], "c": 1},
 {"t": "mc", "o": ["لكل نشاط درجة مستقلة", "المحتوى أكثر من التقويم", "لكل مخرج نشاط يدرّب عليه وتقويم يقيسه", "التقويم متنوع بغض النظر عن المخرجات"], "c": 2},
 {"t": "mc", "o": ["لأن الطلاب لا يحبون الفيديو", "لأن ذلك يلغي الحاجة إلى التقويم", "لأن مدة المحاضرة تصبح أطول", "لأنها أسهل في المتابعة وأخف حجماً ويمكن العودة إلى نقطة محددة"], "c": 3},
 {"t": "mc", "o": ["يركز على إزالة الحواجز أمام الالتحاق، بينما يركز التعليم عن بُعد على التقديم عبر المسافة", "التعليم المفتوح لا يكون عن بُعد أبداً", "هما مترادفان تماماً", "التعليم المفتوح يشترط امتحاناً واحداً للجميع"], "c": 0},
 {"t": "mc", "o": ["انتظار أن يطلب هو المساعدة", "التواصل معه مبكراً عند ظهور علامات التعثر قبل أن تتفاقم", "الاكتفاء بالإعلانات العامة", "تحويله إلى محاضر آخر"], "c": 1},
 {"t": "mc", "o": ["نشر الأسئلة قبل الاختبار", "إلغاء الوقت المحدد", "بنك أسئلة كبير وترتيب عشوائي ومهام تطبيقية يصعب نسخها", "السماح بمشاركة الإجابات"], "c": 2},
 {"t": "mc", "o": ["جودة المحتوى فقط", "عدد الطلاب فقط", "سرعة الإنترنت فقط", "جودة التصميم والدعم والتقويم والعمليات، لا المحتوى وحده"], "c": 3},
 {"t": "mc", "o": ["تقسيمها إلى مقاطع قصيرة مع سؤال بعد كل مقطع ونسخة صوتية خفيفة", "زيادة مدة المحاضرة لتغطي المحتوى كاملاً", "إلغاء التسجيل والاكتفاء بقراءة المقرر", "طلب حضور الطلاب إلى الجامعة"], "c": 0},
 {"t": "mc", "o": ["كثرة الأنشطة التفاعلية", "غياب الاتساق البنائي: التقويم لا يقيس المخرج المعلن", "وضوح مخرجات التعلم", "تنوع مصادر المحتوى"], "c": 1},
 {"t": "mc", "o": ["إلغاء الاختبار لجميع الطلاب فوراً", "تصحيح من أنهى الاختبار وتجاهل المنقطعين", "التحقق من سجل المنصة وإتاحة فرصة بديلة بأسئلة بديلة من البنك", "إعادة الاختبار نفسه بالأسئلة نفسها دون ضوابط"], "c": 2},
 {"t": "mc", "o": ["لا يلزم شيء لأن الاستخدام غير تجاري", "يكفي ذكر المصدر لتعديلها", "يجوز التعديل بشرط نشر المشتقات بالترخيص نفسه", "استئذان صاحب العمل لأن الترخيص يمنع التعديل"], "c": 3},
 {"t": "mc", "o": ["تقديم مواد خفيفة قابلة للتنزيل ومنحه مرونة في موعد التسليم", "إنذاره رسمياً لأن الغياب غير مبرر", "دعوته إلى مقر الجامعة", "تركه ليتدبر أمره"], "c": 0},
 {"t": "mc", "o": ["جلسة مباشرة واحدة إلزامية في وقت واحد", "نقاش غير متزامن في منتدى مع مهلة تكفي للمشاركة", "إلغاء النقاش", "اختبار شفهي لكل طالب في الوقت نفسه"], "c": 1},
 {"t": "mc", "o": ["الاكتفاء بالدرجة المرتفعة دليلاً", "منع الإنترنت عن مدينته", "مقابلة قصيرة يشرح فيها أسباب اختيار إجاباته", "تكرار الاختبار نفسه بالأسئلة نفسها"], "c": 2},
 {"t": "mc", "o": ["أن الطلاب كسالى جميعاً", "حذف الاختبار كله", "رفع صعوبة بقية الأسئلة", "مراجعة السؤال ووضوحه ومدى تغطية المحتوى له قبل لوم الطلاب"], "c": 3},
 {"t": "mc", "o": ["تعليمات واضحة: الهدف والمدة وطريقة التسليم ومعيار التقويم", "مزيد من المحاضرات المسجلة", "تقليل عدد الطلاب", "تغيير منصة التعلم"], "c": 0},
 {"t": "mc", "o": ["إلقاء المحاضرة كاملة في الجلسة ثم واجب", "مقطع قصير ونشاط تمهيدي قبل الجلسة، ثم تخصيص الجلسة لحل مشكلة تطبيقية", "تسجيل الجلسة ونشرها فقط", "اختبار في بداية الجلسة دون تمهيد"], "c": 1},
 {"t": "mc", "o": ["يفهم الدارس التصميم التعليمي", "يقدّر الدارس أهمية الأنشطة", "يصمّم الدارس نشاطاً تفاعلياً مرتبطاً بمخرج تعلم محدد", "يطّلع الدارس على المقرر"], "c": 2},
 {"t": "mc", "o": ["مورد تعليمي مفتوح كامل", "غير مجاني ومفتوح", "عمل مباح لا يحتاج إلى ترخيص", "مجاني لكنه ليس مورداً مفتوحاً بالمعنى الكامل"], "c": 3},
 {"t": "mc", "o": ["أساس غير متزامن مع مواد خفيفة قابلة للتنزيل وقناة تواصل بديلة", "الاعتماد على بث مباشر عالي الجودة", "مقاطع فيديو عالية الدقة فقط", "إلزام الجميع بجلسة مصورة أسبوعية"], "c": 0},
 {"t": "fl", "a": ["التبادلية", "تبادلية"]},
 {"t": "fl", "a": ["المشاركة", "مشاركة"]},
 {"t": "fl", "a": ["البنائي", "بنائي"]},
 {"t": "fl", "a": ["التعليمية"]},
 {"t": "fl", "a": ["الصف", "الفصل"]},
 {"t": "fl", "a": ["تجزئة", "تقسيم", "تقطيع"]},
 {"t": "fl", "a": ["منخفضة", "قليلة", "خفيفة"]},
 {"t": "fl", "a": ["غير المتزامن", "غير متزامن", "اللامتزامن"]}
];

function normAr_(s){ return String(s||'').replace(/[\u064B-\u0652\u0640]/g,'').replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/\s+/g,' ').trim().toLowerCase(); }

/* يصحّح إجابات المتدرب. يرجع null إن كانت الصيغة غير سليمة (أسئلة مكررة أو عدد خاطئ أو ثقة غير مسموحة) */
/* الاختبار القبلي: لا تُرسل الإجابات الصحيحة ولا الشروح إلى المتدرب حتى لا يحفظها فتفقد الاختبار البعدي قيمته (البنك 32 سؤالاً ويُسحب منه 16) */
function stripAns_(rows){ return (rows||[]).map(r => { const o = Object.assign({}, r); delete o.c; delete o.e; return o; }); }

function grade_(ans){
  const total = EXAM_COUNTS.tf + EXAM_COUNTS.mc + EXAM_COUNTS.fl;
  if(!Array.isArray(ans) || ans.length !== total) return null;
  const seen = {}, got = {tf:0, mc:0, fl:0}, rows = [], R = [], W = [];
  let sum = 0, k = 0, cs = 0, okN = 0, over = 0, under = 0;
  for(let n=0;n<ans.length;n++){
    const x = ans[n] || {}, qid = +x.qid, q = QKEY[qid];
    if(!q || qid % 1 !== 0 || seen[qid]) return null;
    seen[qid] = 1; got[q.t]++;
    const a = String(x.a==null ? '' : x.a).trim().slice(0,200);
    let ok = false, av = '', c = '', e, cf = 0, pt = 0;
    if(q.t === 'tf'){ if(a !== '' && a !== '0' && a !== '1') return null; ok = a !== '' && (a === '1') === q.v; av = a === '' ? '' : (a === '1' ? 'صح' : 'خطأ'); c = q.v ? 'صح' : 'خطأ'; e = q.e; }
    if(q.t === 'mc'){ if(a !== '' && !(/^[0-3]$/.test(a))) return null; ok = a !== '' && +a === q.c; av = a === '' ? '' : String(q.o[+a]); c = q.o[q.c]; }
    if(q.t === 'fl'){ ok = a !== '' && q.a.some(function(z){ return normAr_(z) === normAr_(a); }); av = a; c = q.a[0]; pt = ok ? 1 : 0; }
    if(q.t !== 'fl' && a !== ''){
      cf = +x.c; if((q.t === 'tf' ? CONF_TF : CONF_MC).indexOf(cf) < 0) return null;
      pt = ok ? CONF_CORRECT[cf] : -CONF_WRONG[cf];
      k++; cs += cf; if(ok) okN++; if(!ok && cf >= 75) over++; if(ok && cf <= 50) under++;
    }
    sum += pt;
    const r = {n:n+1, qid:qid, q:String(x.q||'').slice(0,300), a:av, c:c, ok:ok, pt:pt}; if(cf) r.cf = cf; if(e !== undefined) r.e = e;
    rows.push(r); (ok ? R : W).push(n+1);
  }
  if(got.tf !== EXAM_COUNTS.tf || got.mc !== EXAM_COUNTS.mc || got.fl !== EXAM_COUNTS.fl) return null;
  return {rows:rows, score:Math.max(0, Math.round(sum*100)/100), total:total, right:R.join('، '), wrong:W.join('، '),
          calib:{k:k, conf:k?Math.round(cs/k):0, acc:k?Math.round(okN/k*100):0, over:over, under:under}};
}

/* ===================== إشارات المتصفح: زمن كل سؤال والخروج واللصق ===================== */
/* المتصفح هو من يرسل هذه البيانات، فهي إشارات للمراجعة لا دليل قاطع. الخادم يقيّد قيمها ويبني نص الإشارات المختصر بنفسه
   (الزمن الكلي من ساعة الخادم) كي لا يُمحى بحذف الإشارات من الطلب. تفسير الخطورة بالألوان يتم في لوحة المدرب. */
function sigClean_(raw, n, mins){
  let o = null; try{ o = JSON.parse(String(raw||'')); }catch(e){}
  const num = (v, mx) => { v = +v; return isFinite(v) && v > 0 ? Math.min(Math.round(v), mx) : 0; };
  const f = [];
  if(mins < 2) f.push('إنهاء سريع جداً ('+mins+' د)');
  if(!o || typeof o !== 'object'){ f.push('لا توجد بيانات تتبع من المتصفح'); return {sig:'', flags:f.join('، ')}; }
  const out = {v:2, blur:num(o.blur,999), blurMs:num(o.blurMs,36e6), paste:num(o.paste,999), tm:o.tm?1:0, pl:o.pl?1:0, t:{}, lim:[], q:[]};
  if(o.t && typeof o.t === 'object') Object.keys(o.t).slice(0,n).forEach(k => { if(/^\d{1,2}$/.test(k)) out.t[k] = num(o.t[k], 36e6); });
  (Array.isArray(o.lim) ? o.lim : []).slice(0,n).forEach(v => out.lim.push(num(v,600)));
  (Array.isArray(o.q) ? o.q : []).slice(0,n).forEach(x => { x = x || {};
    out.q.push({m:num(x.m,36e6), a:num(x.a,36e6), n:num(x.n,99), c:num(x.c,99), f:(+x.f >= 0 ? num(x.f,36e6) : -1), t:x.t?1:0, p:num(x.p,99), u:num(x.u,36e6), o:num(x.o,36e6), d:x.d?1:0}); });
  if(out.blur >= 2 || out.blurMs > 45000) f.push('خروج من الصفحة ×'+out.blur+' ('+Math.round(out.blurMs/1000)+' ث)');
  if(out.paste > 0) f.push('لصق نص ×'+out.paste);
  if(out.q.length < n) f.push('بيانات زمن الأسئلة ناقصة');
  return {sig:JSON.stringify(out), flags:f.join('، ')};
}

function certState_(ph, start, ctx){   // ctx اختياري: بيانات مقروءة مسبقاً لتسريع الحلقات                    // حالة الشهادة لنتيجة بعدية ناجحة: لا تصدر إلا باكتمال كل الشروط
  const rs = ctx ? ctx.rs : readRows_(), r = rs.filter(x => x.wa === ph && x.start === String(start) && x.phase !== 'pre')[0];
  if(!r || !r.cert) return {pass:false, allowed:false, needs:[]};
  const needs = [];
  if(CERT_NEEDS_PRE && !rs.some(x => x.wa === ph && x.phase === 'pre')) needs.push('pre');
  if(CERT_NEEDS_PROJECT){ const pr = (ctx ? ctx.prs : projRows_()).filter(x => x.ph === ph)[0]; if(!pr || pr.status !== 'approved') needs.push('project'); }
  const vs = (ctx ? ctx.vrs : vidRows_()).filter(v => v.ph === ph && v.start === String(start));
  if(CERT_NEEDS_VIDEO && !vs.length) needs.push('video');
  else if(CERT_NEEDS_VERIFY && !vs.some(v => v.ver === 'ok')) needs.push('verify');
  if(CERT_NEEDS_ORAL && r.oral !== '1') needs.push('oral');
  return {pass:true, allowed:needs.length===0, needs:needs, cert:r.cert};
}
function certNameFor_(ph, fallback){ const v = PropertiesService.getScriptProperties().getProperty('cname_'+ph); return v ? v : fallback; }   // الاسم المعدَّل من مدير النظام
function readNotes_(){
  const s = sheet_('Notes', ['ts','text']), v = s.getDataRange().getValues(), out = [];
  for(let i=1;i<v.length;i++) if(v[i][1]) out.push({ts:String(v[i][0]), text:String(v[i][1])});
  return out;
}


/* ===================== التغذية الراجعة للمتدرب ===================== */
function pctOf_(r){ const sc = +r.score, tt = (+r.total) || ((+r.right||0)+(+r.wrong||0)); return (isFinite(sc) && tt > 0) ? Math.round(sc/tt*100) : null; }
function jparse_(s, d){ try{ const v = JSON.parse(s); return (v === null || v === undefined) ? d : v; }catch(e){ return d; } }
/* يجمع للمتدرب نفسه: نتائج اختباراته وتحليلها + مشروعه وتقييم المدرب + مقابلاته وملاحظاتها + حالة الشهادة.
   مقارنة المجموعة مجهولة الهوية ولا تظهر إلا إذا بلغ عدد المتدربين 3 فأكثر */
function myFeedback_(p){
  const ph = odlPhone_(p.tok); if(!ph) return J_({ok:false, err:'unauth', code:'SESSION_INVALID'});
  const res = readRows_(), mine = res.filter(r => r.wa === ph).sort((a,b) => a.start < b.start ? 1 : -1);
  const pr = projRows_().filter(r => r.ph === ph)[0] || null;
  const vids = vidRows_().filter(v => v.ph === ph).sort((a,b) => (+b.ts) - (+a.ts));
  let fbts = 0;
  if(pr && pr.fbts) fbts = Math.max(fbts, +pr.fbts || 0);
  vids.forEach(v => { if(v.fbts) fbts = Math.max(fbts, +v.fbts || 0); if(v.vts) fbts = Math.max(fbts, +v.vts || 0); });
  if(String(p.light||'') === '1') return J_({ok:true, fbts:fbts, attempts:mine.length, proj:pr ? {status:pr.status||'submitted'} : null});
  const u = odlGet_(ph) || {};
  const attempts = mine.slice(0,8).map(r => ({phase:r.phase||'post', start:r.start, end:r.end, mins:r.mins, score:r.score, total:r.total, pct:pctOf_(r), cert:!!r.cert, rows:(r.phase === 'pre' ? stripAns_(jparse_(r.rows,[])) : jparse_(r.rows,[])), calib:jparse_(r.calib,null)}));
  const lat = {}; res.forEach(r => { const k = r.wa+'|'+(r.phase||'post'); if(!lat[k] || r.start > lat[k].start) lat[k] = r; });
  const L = Object.keys(lat).map(k => lat[k]), who = {}; L.forEach(r => { who[r.wa] = 1; });
  let group = null;
  if(Object.keys(who).length >= 3){
    const perQ = {}, A = {pre:[], post:[]};
    L.forEach(r => {
      const pc = pctOf_(r); if(pc !== null) A[r.phase === 'pre' ? 'pre' : 'post'].push(pc);
      jparse_(r.rows,[]).forEach(x => { const q = x && x.qid; if(q === undefined || q === null) return; perQ[q] = perQ[q] || {ok:0, n:0}; perQ[q].n++; if(x.ok) perQ[q].ok++; });
    });
    const avg = a => a.length >= 3 ? Math.round(a.reduce((s,v) => s+v, 0)/a.length) : null;
    Object.keys(perQ).forEach(q => { if(perQ[q].n < 3) delete perQ[q]; });
    group = {n:Object.keys(who).length, avgPre:avg(A.pre), avgPost:avg(A.post), perQ:perQ};
  }
  let cert = {none:true};
  const m = mine.filter(r => r.phase !== 'pre' && r.cert)[0];
  if(m){ const st = certState_(ph, m.start); cert = {none:false, allowed:st.allowed, needs:st.needs}; }
  return J_({ok:true, name:u.name||'', passPct:PASS_PCT, fbts:fbts, attempts:attempts, group:group, cert:cert,
    proj: pr ? {title:pr.title, link:pr.link, checks:pr.checks, score:pr.score, note:pr.note, ts:pr.ts, feedback:pr.feedback, fileName:pr.fileName, status:pr.status||'submitted', rubric:pr.rubric, fbts:pr.fbts} : null,
    vids: vids.map(v => ({ts:v.ts, start:v.start, secs:v.secs, prompt:v.prompt, fb:v.fb, rating:v.rating, fbts:v.fbts, ver:v.ver, vchk:v.vchk, vnote:v.vnote, vts:v.vts}))});
}

function doGet(e){
  const p = (e && e.parameter) || {};
  if(p.pin !== undefined && String(p.pin) !== '') return J_({ok:false, err:'post_only'});   // رمز المدرب لا يُقبل في الرابط: أرسله بـ POST
  if(p.action === 'sys_status') return J_(swPublic_());                  // حالة المنصة + هوية الجهة (علني، بلا بيانات حساسة)
  const pg = swGate_(p.action, false); if(pg) return pg;                // المنصة موقوفة: تُحجب إجراءات المتدرب
  return handleGet_(p);
}
/* أسماء التوقيع في الشهادة: تُحفظ في Script Properties ويعدّلها المدرب من لوحته */
function certNames_(){
  let o = {}; try{ o = JSON.parse(PropertiesService.getScriptProperties().getProperty('cert_names')||'{}') || {}; }catch(e){}
  return {exec: o.exec === undefined ? 'أ.د. عبد الباسط محمد شريف' : String(o.exec), mgr: o.mgr === undefined ? '' : String(o.mgr)};
}

function handleGet_(p){
  try{
    const a = odlHandle_(p); if(a) return a;
    const tk0 = trkHandle_(p, false); if(tk0) return tk0;            // trk_me للمتدرب (رمزه)؛ المدير والمدرب POST فقط
    if(p.action === 'list'){
      if(p.pin !== PIN) return J_({ok:false, err:'pin'});
      return J_({ok:true, rows:readRows_()});
    }
    if(p.action === 'cert_names') return J_(Object.assign({ok:true}, certNames_()));   // يقرؤها المتدرب لرسم الشهادة
    if(p.action === 'notes') return J_({ok:true, rows:readNotes_()});
    if(p.action === 'proj_mine'){                  // مشروع المتدرب نفسه وتعليق المدرب
      const ph = odlPhone_(p.tok); if(!ph) return J_({ok:false, err:'unauth', code:'SESSION_INVALID'});
      return J_({ok:true, row:projRows_().filter(r => r.ph === ph)[0] || null});
    }
    if(p.action === 'proj_list'){                  // كل المشاريع (للمدرب)
      if(p.pin !== PIN) return J_({ok:false, err:'pin'});
      return J_({ok:true, rows:projRows_()});
    }
    if(p.action === 'proj_fb'){                    // تقييم المدرب لمشروع: تعليق + حالة + تقدير لكل معيار (يصل المتدرب)
      if(p.pin !== PIN) return J_({ok:false, err:'pin'});
      const s = sheet_('Projects', PROJ_HEADERS), v = s.getDataRange().getValues(), ph = odlNorm_(p.ph);
      const ST = ['submitted','reviewing','needs_work','approved','resubmitted'];
      for(let i=1;i<v.length;i++) if(String(v[i][0]) === ph){
        let st = ST.indexOf(String(p.status)) >= 0 ? String(p.status) : String(v[i][12]||'reviewing');
        const rub = (p.rubric === undefined) ? String(v[i][13]||'') : String(p.rubric||'').replace(/[^012\-]/g,'').slice(0,20);
        if((st === 'submitted' || st === 'resubmitted') && (String(p.text||'').trim() || /[012]/.test(rub))) st = 'reviewing';   // التعليق أو التقييم يعني أن المراجعة بدأت
        const oldFb = String(v[i][8]||''), oldSt = String(v[i][12]||''), oldRub = String(v[i][13]||''), newFb = String(p.text||'').slice(0,600);
        s.getRange(i+1, 9).setValue(newFb);
        s.getRange(i+1, 13, 1, 3).setValues([[st, rub, String(Date.now())]]);
        if(newFb !== oldFb || st !== oldSt || rub !== oldRub){ try{ trkNotifyProjFb_(ph, st, newFb, rub); }catch(e){} }   // تصل التغذية الراجعة المتدرب تلقائياً على واتساب
        return J_({ok:true});
      }
      return J_({ok:false, err:'notfound'});
    }
    if(p.action === 'my_feedback') return myFeedback_(p);   // المتدرب: لوحة تغذيته الراجعة (اختبارات + مشروع + مقابلة + شهادة)
    if(p.action === 'video_fb'){                   // تقدير المدرب وملاحظاته على مقابلة مسجلة (يصل المتدرب)
      if(p.pin !== PIN) return J_({ok:false, err:'pin'});
      const s = sheet_('Videos', VID_HEADERS), v = s.getDataRange().getValues(), ph = odlNorm_(p.ph);
      for(let i=1;i<v.length;i++) if(String(v[i][0]) === ph && String(v[i][8]) === String(p.ts)){
        const oldFb = String(v[i][9]||''), oldRt = String(v[i][10]||''), newFb = String(p.text||'').slice(0,600), newRt = /^[123]$/.test(String(p.rating)) ? String(p.rating) : '';
        s.getRange(i+1, 10, 1, 3).setValues([[newFb, newRt, String(Date.now())]]);
        if((newFb !== oldFb || newRt !== oldRt) && (newFb || newRt)){ try{ trkNotifyVidFb_(ph, newRt, newFb); }catch(e){} }
        return J_({ok:true});
      }
      return J_({ok:false, err:'notfound'});
    }
    if(p.action === 'video_verify'){               // تقرير المدرب على الفيديو: تحقق أن المتدرب هو من أدّى الاختبار (شرط إصدار الشهادة)
      if(p.pin !== PIN) return J_({ok:false, err:'pin'});
      const dec = String(p.ver) === 'ok' ? 'ok' : String(p.ver) === 'redo' ? 'redo' : '';
      if(!dec) return J_({ok:false, err:'decision'});
      const chk = String(p.chk||'').replace(/[^01]/g,'').slice(0,3), note = String(p.text||'').trim().slice(0,600);
      if(dec === 'ok' && chk !== '111') return J_({ok:false, err:'incomplete'});     // لا اعتماد دون استيفاء الشروط الثلاثة
      if(dec === 'redo' && !note) return J_({ok:false, err:'note'});                  // سبب الإعادة إلزامي
      const s = sheet_('Videos', VID_HEADERS), v = s.getDataRange().getValues(), ph = odlNorm_(p.ph);
      for(let i=1;i<v.length;i++) if(String(v[i][0]) === ph && String(v[i][8]) === String(p.ts)){
        const old = String(v[i][12]||'');
        s.getRange(i+1, 13, 1, 4).setValues([[dec, (chk+'000').slice(0,3), note, String(Date.now())]]);
        try{ trkNotifyVerify_(ph, dec, note); }catch(e){}
        return J_({ok:true});
      }
      return J_({ok:false, err:'notfound'});
    }
    if(p.action === 'my_cert'){                    // المتدرب: حالة شهادته
      const ph = odlPhone_(p.tok); if(!ph) return J_({ok:false, err:'unauth', code:'SESSION_INVALID'});
      const m = readRows_().filter(r => r.wa === ph && r.phase !== 'pre' && r.cert).sort((a,b) => a.start < b.start ? 1 : -1)[0];
      if(!m) return J_({ok:true, none:true});
      const st = certState_(ph, m.start);
      const sc_ = isFinite(+m.score) && String(m.score||'') !== '' ? +m.score : +m.right, tt_ = (+m.total) || ((+m.right||0)+(+m.wrong||0)), pc_ = tt_ > 0 && isFinite(sc_) ? Math.round(sc_/tt_*100) : null;
      return J_({ok:true, start:m.start, needs:st.needs, cert:st.allowed?st.cert:'', pct:pc_, score:sc_, total:tt_, name:certNameFor_(ph, m.name)});
    }
    if(p.action === 'video_list'){                 // فيديوهات المقابلة (للمدرب)
      if(p.pin !== PIN) return J_({ok:false, err:'pin'});
      return J_({ok:true, rows:vidRows_().sort((a,b) => (+b.ts)-(+a.ts))});
    }
    if(p.action === 'oral_set'){                   // تأكيد المقابلة الشفوية (للمدرب)
      if(p.pin !== PIN) return J_({ok:false, err:'pin'});
      const s = sheet_('Results', RES_HEADERS), v = s.getDataRange().getValues();
      for(let i=1;i<v.length;i++) if(String(v[i][1])===String(p.wa) && String(v[i][5])===String(p.start)){ s.getRange(i+1,14).setValue(p.val==='1'?'1':''); return J_({ok:true}); }
      return J_({ok:false, err:'notfound'});
    }
    if(p.action === 'cert'){                       // تحقق عام من شهادة
      const code = String(p.code||'').trim().toUpperCase(); if(!code) return J_({ok:false});
      const r = readRows_().filter(x => x.cert === code)[0];
      if(!r || !certState_(r.wa, r.start).allowed){ const ar = certArchiveFind_(code); return ar ? J_({ok:true, name:ar.name, score:ar.score, total:ar.total, date:ar.date, archived:true}) : J_({ok:false}); }   // شهادات الدورات السابقة بعد التصفير
      return J_({ok:true, name:certNameFor_(r.wa, r.name), score:r.score, total:r.total, date:r.end});
    }
    return J_({ok:true, msg:'خادم الحقيبة التدريبية يعمل'});
  }catch(err){ return J_({ok:false, err:String(err)}); }
}

function doPost(e){
  try{
    const b = JSON.parse(e.postData.contents);
    const gate = pinGate_(b); if(gate) return gate;                 // قفل المحاولات الخاطئة لرمز المدرب
    if(b.action === 'sys_status') return J_(swPublic_());
    const swr = swHandle_(b); if(swr) return swr;                   // مفاتيح التشغيل + دورة جديدة (Switches.gs)
    const pg = swGate_(b.action, b.pin !== undefined && String(b.pin) !== ''); if(pg) return pg;   // المنصة موقوفة؟
    if(b.pin !== undefined && String(b.pin) !== '' && ['list','proj_list','proj_fb','video_list','oral_set','video_fb','video_verify'].indexOf(b.action) >= 0) return handleGet_(b);   // طلبات المدرب كانت GET وصارت POST
    const a = odlHandle_(b); if(a) return a;
    const tk = trkHandle_(b, true); if(tk) return tk;                      // التتبع وتقارير واتساب (Tracking.gs)
    const sm = simHandle_(b); if(sm) return sm;                            // أصوات الشخصيات (tts) ونتائج المحاكاة (sim_save/sim_list)

    if(b.action === 'cert_name_list' || b.action === 'cert_name_set'){   // مدير النظام فقط: تعديل اسم المتدرب المطبوع في الشهادة
      if(String(b.pin) !== PIN) return J_({ok:false, err:'pin'});
      const ag = adminGate_(b); if(ag) return ag;
      const P = PropertiesService.getScriptProperties();
      if(b.action === 'cert_name_list'){
        const all = P.getProperties(), map = {};
        Object.keys(all).forEach(k => { if(k.indexOf('u_') === 0){ try{ const u = JSON.parse(all[k]); map[k.slice(2)] = {ph:k.slice(2), reg:u.name||'', res:'', rs:'', passed:false}; }catch(e){} } });
        readRows_().forEach(r => { if(r.phase === 'pre') return; const m = map[r.wa] = map[r.wa] || {ph:r.wa, reg:'', res:'', rs:'', passed:false}; if(!m.rs || r.start > m.rs){ m.res = r.name; m.rs = r.start; } if(r.cert) m.passed = true; });
        const rows = Object.keys(map).map(ph => ({ph:ph, reg:map[ph].reg, res:map[ph].res, passed:map[ph].passed, over:all['cname_'+ph]||''}));
        rows.sort((x,y) => (y.passed?1:0)-(x.passed?1:0));
        return J_({ok:true, rows:rows});
      }
      const ph = odlNorm_(b.ph); if(!ph) return J_({ok:false, err:'ph'});
      const nm = String(b.name||'').replace(/[\u0000-\u001f<>]/g,'').replace(/\s+/g,' ').trim().slice(0,100);
      if(nm && nm.length < 3) return J_({ok:false, err:'short'});
      const oldNm = P.getProperty('cname_'+ph) || '';
      if(nm) P.setProperty('cname_'+ph, nm); else P.deleteProperty('cname_'+ph);
      try{ sheet_('NameEdits', ['ts','ph','old','new']).appendRow([new Date().toISOString(), ph, oldNm, nm]); }catch(e){}
      return J_({ok:true, name:nm});
    }
    if(b.action === 'cert_names_set'){             // المدرب: تعديل أسماء التوقيع في الشهادة
      if(String(b.pin) !== PIN) return J_({ok:false, err:'pin'});
      const o = {exec:String(b.exec||'').replace(/\s+/g,' ').trim().slice(0,80), mgr:String(b.mgr||'').replace(/\s+/g,' ').trim().slice(0,80)};
      PropertiesService.getScriptProperties().setProperty('cert_names', JSON.stringify(o));
      return J_({ok:true, exec:o.exec, mgr:o.mgr});
    }
    if(b.action === 'save'){                       // حفظ نتيجة اختبار (للمتدرب الموثّق فقط): الخادم يصحّح ويحسب الدرجة
      const ph = odlPhone_(b.tok); if(!ph) return J_({ok:false, err:'unauth', code:'SESSION_INVALID'});
      const g = grade_(b.ans); if(!g) return J_({ok:false, err:(b.ans ? 'bad' : 'old_client')});
      const d = b.data || {}, now = Date.now(), st = new Date(String(d.start||'')).getTime();
      if(!isFinite(st) || st > now+60000 || st < now-12*3600000) return J_({ok:false, err:'bad_start'});
      d.wa = ph;
      d.phase = (d.phase === 'pre') ? 'pre' : 'post';
      d.start = new Date(st).toISOString(); d.end = new Date(now).toISOString(); d.mins = Math.round((now-st)/600)/100;
      d.name = String(d.name||'').trim().slice(0,100); d.email = String(d.email||'').trim().slice(0,100);
      d.score = g.score; d.total = g.total; d.right = g.right; d.wrong = g.wrong; d.rows = JSON.stringify(g.rows); d.calib = JSON.stringify(g.calib);
      const sgc = sigClean_(d.sig, g.rows.length, d.mins); d.sig = sgc.sig; d.flags = sgc.flags;   // إشارات التتبع: تُنظَّف وتُحسب في الخادم
      const pass = d.phase === 'post' && g.score/(g.total||1)*100 >= PASS_PCT;
      const cert = pass ? odlHash_('cert|'+ph+'|'+d.start).slice(0,10).toUpperCase() : '';
      d.cert = cert; d.oral = '';
      const out = {ok:true, score:g.score, total:g.total, right:g.right, wrong:g.wrong, rows:(d.phase === 'pre' ? stripAns_(g.rows) : g.rows), calib:g.calib, start:d.start, end:d.end, mins:d.mins};
      const lock = LockService.getScriptLock(); lock.waitLock(10000);
      try{
        const s = sheet_('Results', RES_HEADERS), v = s.getDataRange().getValues();
        for(let i=1;i<v.length;i++) if(String(v[i][1])===ph){
          if(String(v[i][5])===String(d.start)){ const st2 = certState_(ph, d.start); out.dup = true; out.rep = 'dup'; out.cert = st2.allowed?st2.cert:''; out.needs = st2.needs; return J_(out); }   // إعادة إرسال لنفس المحاولة
          const pe = Date.parse(String(v[i][6])); if(isFinite(pe) && now-pe < 60000) return J_({ok:false, err:'wait'});   // تباعد دقيقة بين نتيجتين
        }
        const rowV = RES_HEADERS.map(k => String(d[k]===undefined ? '' : d[k]).slice(0,45000));
        s.getRange(s.getLastRow()+1, 1, 1, rowV.length).setNumberFormat('@').setValues([rowV]);   // نص عادي دائماً: يمنع تحويل 1.5 إلى تاريخ
      }finally{ lock.releaseLock(); }
      const stN = pass ? certState_(ph, d.start) : {allowed:false, needs:[]};
      out.cert = stN.allowed ? cert : ''; out.needs = stN.needs;
      try{ out.rep = trkReport_(ph, d, g, out.cert); }catch(e){ out.rep = 'off'; }   // التقرير يُرسل من الخادم إلى رقم المتدرب المسجَّل
      return J_(out);
    }
    if(b.action === 'proj_save'){                  // حفظ/تحديث مشروع المتدرب (واحد لكل متدرب) مع ملف اختياري يُحفظ في Drive المدرب
      const ph = odlPhone_(b.tok); if(!ph) return J_({ok:false, err:'unauth', code:'SESSION_INVALID'});
      const u = odlGet_(ph) || {}, s = sheet_('Projects', PROJ_HEADERS), v = s.getDataRange().getValues();
      let at = -1; for(let i=1;i<v.length;i++) if(String(v[i][0]) === ph){ at = i; break; }
      const old = at > 0 ? v[at] : [];
      let fUrl = String(old[9]||''), fId = String(old[10]||''), fName = String(old[11]||'');
      const data = String(b.fileData||'');
      if(data){
        const nm = String(b.fileName||'ملف').replace(/[\\\/:*?"<>|]/g,'_').slice(0,80), ext = (nm.split('.').pop()||'').toLowerCase();
        if(!PROJ_EXT[ext]) return J_({ok:false, err:'bad_type'});
        if(data.length > PROJ_FILE_MAX*4/3 + 8) return J_({ok:false, err:'too_big'});
        let bytes; try{ bytes = Utilities.base64Decode(data); }catch(e){ return J_({ok:false, err:'bad_file'}); }
        if(bytes.length > PROJ_FILE_MAX) return J_({ok:false, err:'too_big'});
        const it = DriveApp.getFoldersByName(PROJ_FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(PROJ_FOLDER);
        const f = folder.createFile(Utilities.newBlob(bytes, PROJ_EXT[ext], ph+'_'+(u.name||'')+'_'+Date.now()+'.'+ext));   // خاص بالمدرب فقط: لا يُشارك مع أحد
        if(fId){ try{ DriveApp.getFileById(fId).setTrashed(true); }catch(e){} }                                           // ملف المتدرب السابق يُنقل إلى المحذوفات
        fUrl = f.getUrl(); fId = f.getId(); fName = nm;
      }
      const row = [ph, u.name||'', String(b.title||'').slice(0,200), String(b.link||'').slice(0,500), String(b.checks||'').slice(0,20), String(+b.score||0), String(b.note||'').slice(0,3000), String(Date.now()), at>0?String(old[8]||''):'', fUrl, fId, fName];
      const oldSt = at > 0 ? String(old[12]||'') : '';
      const changed = at > 0 && (data || String(old[2]) !== row[2] || String(old[3]) !== row[3] || String(old[4]) !== row[4] || String(old[6]) !== row[6]);
      const st = (changed && (oldSt === 'needs_work' || oldSt === 'approved')) ? 'resubmitted' : (oldSt || 'submitted');   // التعديل بعد التقييم يعيد المشروع للمراجعة
      row.push(st, at > 0 ? String(old[13]||'') : '', at > 0 ? String(old[14]||'') : '');
      if(at > 0){ s.getRange(at+1,1,1,row.length).setValues([row]); return J_({ok:true, file:fName, status:st}); }
      s.appendRow(row); return J_({ok:true, file:fName, status:st});
    }
    if(b.action === 'video_save'){                 // فيديو المقابلة المسجلة → Google Drive
      const ph = odlPhone_(b.tok); if(!ph) return J_({ok:false, err:'unauth', code:'SESSION_INVALID', error:'انتهت جلسة الدخول. سجّل الدخول من جديد ثم أعد الإرسال.'});
      const u = odlGet_(ph) || {};
      if(vidRows_().filter(v => v.ph === ph && v.ver !== 'redo').length >= VIDEO_MAX_PER_USER) return J_({ok:false, error:'بلغتَ الحد الأقصى لمحاولات التسجيل.'});
      const data = String(b.data||'').replace(/^[\s\S]*?base64,/,'').replace(/\s+/g,''); if(!data || data.length > 8000000) return J_({ok:false, error:'حجم الفيديو غير مناسب. سجّل مقطعاً أقصر.'});
      const mime = /^video\/mp4/.test(String(b.mime||'')) ? 'video/mp4' : 'video/webm', ext = mime === 'video/mp4' ? 'mp4' : 'webm';
      try{
        const blob = Utilities.newBlob(Utilities.base64Decode(data), mime, ph+'_'+Date.now()+'.'+ext);
        const it = DriveApp.getFoldersByName(VIDEO_FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(VIDEO_FOLDER);
        const f = folder.createFile(blob);
        sheet_('Videos', VID_HEADERS).appendRow([ph, u.name||'', String(b.start||''), f.getId(), f.getUrl(), String(Math.min(60,+b.secs||0)), mime, String(b.prompt||'').slice(0,800), String(Date.now())]);
      }catch(er){
        return J_({ok:false, err:'drive', error:'تعذّر حفظ الفيديو في Google Drive. على المدرب تشغيل الدالة authorize() من محرر Apps Script ثم إعادة النشر بإصدار جديد. (' + String(er).slice(0,140) + ')'});
      }
      return J_({ok:true});
    }
    if(b.action === 'note'){                       // نشر إشعار (للمدرب)
      if(b.pin !== PIN) return J_({ok:false, err:'pin'});
      const text = String(b.text||'').slice(0,500); if(!text) return J_({ok:false});
      const s = sheet_('Notes', ['ts','text']);
      s.appendRow([String(Date.now()), text]);
      const n = s.getLastRow(); if(n > 51) s.deleteRows(2, n-51);
      return J_({ok:true});
    }
    if(b.action === 'delete'){                     // حذف نتيجة أو كل النتائج (للمدرب)
      if(b.pin !== PIN) return J_({ok:false, err:'pin'});
      const s = sheet_('Results', RES_HEADERS), n = s.getLastRow();
      if(b.all){ if(n>1) s.deleteRows(2, n-1); return J_({ok:true}); }
      const v = s.getDataRange().getValues();
      for(let i=v.length-1;i>=1;i--) if(String(v[i][5])===String(b.start) && String(v[i][1])===String(b.wa)) s.deleteRow(i+1);
      return J_({ok:true});
    }
    if(b.action === 'ai'){                         // توليد بالذكاء الاصطناعي: للمتدرب الموثّق بعدد محدود من المحاولات
      const key = PropertiesService.getScriptProperties().getProperty('OPENROUTER_KEY');
      if(!key) return J_({ok:false, error:'لم يُضبط مفتاح الذكاء الاصطناعي في الخادم'});
      const q = aiTake_(b); if(!q.ok) return J_(q);
      const mdl = /(:free|^openrouter\/free)$/.test(String(b.model||'')) ? String(b.model) : 'openrouter/free';   // نماذج مجانية فقط
      let t = '', busy = false;
      const tryM = function(m){                                    // طلب واحد إلى OpenRouter بنموذج محدد
        const r = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {method:'post', contentType:'application/json',
          headers:{Authorization:'Bearer '+key}, muteHttpExceptions:true,
          payload:JSON.stringify({model:m, messages:[{role:'user', content:String(b.prompt||'').slice(0,30000)}]})});
        busy = (r.getResponseCode ? r.getResponseCode() : 200) === 429;
        const j = JSON.parse(r.getContentText()); return (j.choices && j.choices[0] && j.choices[0].message.content) || '';
      };
      try{
        t = tryM(mdl);
        if(!t && !busy && mdl !== 'openrouter/free'){              // النموذج المختار غير متاح: جرّب الموجّه التلقائي مرة واحدة
          aiCount_(); t = tryM('openrouter/free');
        }
      }catch(err){ t = ''; }
      if(!t){ const left = aiRefund_(b); return J_({ok:false, error: busy ? 'الخدمة المجانية مزدحمة الآن. انتظر دقيقة وأعد المحاولة (لم تُخصم محاولتك).' : 'تعذّر التوليد (لم تُخصم محاولتك).', left:left}); }
      return J_({ok:true, text:t, left:q.left});
    }
    return J_({ok:false, err:'unknown'});
  }catch(err){ return J_({ok:false, err:String(err)}); }
}

/* ===================== محاولات الذكاء الاصطناعي ===================== */
function aiToday_(){ return Utilities.formatDate(new Date(), AI_TZ, 'yyyy-MM-dd'); }
function aiLeftOf_(u){                                                // المتبقي للعرض دون حفظ
  if(AI_DAILY_RESET && u.aiDay !== aiToday_()) return AI_ATTEMPTS;
  return u.aiLeft === undefined ? AI_ATTEMPTS : u.aiLeft;
}
function aiFresh_(u){                                                 // يبدأ يوماً جديداً للمتدرب عند الحاجة
  const d = aiToday_();
  if((AI_DAILY_RESET && u.aiDay !== d) || u.aiLeft === undefined){ u.aiDay = d; u.aiLeft = AI_ATTEMPTS; u.aiReq = 0; u.run = null; }
}
function aiTake_(b){
  if(b.pin === PIN){                                                  // المدرب بلا حد، لكن طلباته تُحتسب في العدّاد اليومي
    const P0 = odlProps_(), gk0 = 'aig_'+aiToday_();
    P0.setProperty(gk0, String(+(P0.getProperty(gk0)||0)+1)); return {ok:true, left:null};
  }
  const ph = odlPhone_(b.tok);
  if(!ph) return {ok:false, code:'unauth', err:'unauth', error:'سجّل الدخول أولاً لاستخدام الذكاء الاصطناعي.'};
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const u = odlGet_(ph); if(!u) return {ok:false, code:'unauth', err:'unauth', error:'رقمك غير مسجل.'};
    aiFresh_(u);
    const run = String(b.run||'').slice(0,40);
    if(run && u.run && u.run.id === run){                             // خطوة ثانية داخل التشغيل نفسه (فريق الوكلاء) لا تُحتسب
      if(u.run.calls >= 4) return {ok:false, code:'run', error:'تجاوز الطلب الحد المسموح.', left:u.aiLeft};
      const P2 = odlProps_(), gk2 = 'aig_'+aiToday_(), g2 = +(P2.getProperty(gk2)||0);
      if(g2 >= AI_GLOBAL_DAILY) return {ok:false, code:'global', left:u.aiLeft, error:'اكتمل الحد اليومي لخدمة الذكاء الاصطناعي. حاول غداً.'};
      P2.setProperty(gk2, String(g2+1));                              // كل طلب فعلي إلى OpenRouter يُحتسب في السقف اليومي
      u.run.calls++; odlPut_(ph,u); return {ok:true, left:u.aiLeft};
    }
    if(u.aiLeft <= 0) return {ok:false, code:'quota', left:0, error:(AI_DAILY_RESET?'نفدت محاولاتك لهذا اليوم. تتجدد تلقائياً غداً، أو تواصل مع المدرب لتجديدها.':'نفدت محاولاتك المجانية. تواصل مع المدرب لتجديدها.')};
    const P = odlProps_(), gk = 'aig_'+aiToday_(), g = +(P.getProperty(gk)||0);
    if(g >= AI_GLOBAL_DAILY) return {ok:false, code:'global', left:u.aiLeft, error:'اكتمل الحد اليومي لخدمة الذكاء الاصطناعي. حاول غداً.'};
    P.setProperty(gk, String(g+1));
    u.aiLeft--; u.run = {id:run||('r'+Date.now()), calls:1}; odlPut_(ph,u);
    return {ok:true, left:u.aiLeft};
  } finally { lock.releaseLock(); }
}
function aiCount_(){ const P = odlProps_(), gk = 'aig_'+aiToday_(); P.setProperty(gk, String(+(P.getProperty(gk)||0)+1)); }
function aiRefund_(b){                                                // فشل الاستدعاء: نُعيد المحاولة
  if(b.pin === PIN) return null;
  const ph = odlPhone_(b.tok); if(!ph) return null;
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const u = odlGet_(ph); if(!u) return null;
    if(u.run && u.run.calls === 1){
      u.aiLeft = (u.aiLeft||0)+1; u.run = null; odlPut_(ph,u);
      const P = odlProps_(), gk = 'aig_'+aiToday_(); P.setProperty(gk, String(Math.max(0,+(P.getProperty(gk)||0)-1)));
    }
    return u.aiLeft;
  } finally { lock.releaseLock(); }
}

/* ===================== الدخول برمز التحقق ===================== */
const ODL_CODE_TTL_MIN=15, ODL_TOKEN_TTL_H=168, ODL_MAX_TRIES=5;   // جلسة المتدرب أسبوع كامل (168 ساعة)
const PIN_MAX_TRIES=5, PIN_LOCK_MIN=15;                           // قفل دخول المدرب: 5 محاولات خاطئة ثم قفل 15 دقيقة

/* بوابة رمز المدرب: تحسب المحاولات الخاطئة وتقفل مؤقتاً. ترجع null إن لم يكن الطلب يحمل رمزاً أو كان الرمز صحيحاً */
function adminGate_(o){                               // رمز مدير النظام: قفل مؤقت بعد المحاولات الخاطئة
  if(!ADMIN_PIN || String(ADMIN_PIN).indexOf('ضع') === 0 || ADMIN_PIN === PIN) return J_({ok:false, err:'admin_not_set'});
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const P = PropertiesService.getScriptProperties(), now = Date.now();
    let st = {n:0, until:0}; try{ st = JSON.parse(P.getProperty('apin_fail')||'') || st; }catch(e){}
    if(st.until > now) return J_({ok:false, err:'locked', wait:Math.ceil((st.until-now)/60000)});
    if(String(o.apin||'') === ADMIN_PIN){ if(st.n) P.setProperty('apin_fail', JSON.stringify({n:0, until:0})); return null; }
    st.n = (st.n||0) + 1;
    if(st.n >= PIN_MAX_TRIES){ P.setProperty('apin_fail', JSON.stringify({n:0, until:now+PIN_LOCK_MIN*60000})); return J_({ok:false, err:'locked', wait:PIN_LOCK_MIN}); }
    P.setProperty('apin_fail', JSON.stringify({n:st.n, until:0}));
    return J_({ok:false, err:'apin', left:PIN_MAX_TRIES-st.n});
  } finally { lock.releaseLock(); }
}
function pinGate_(o){
  if(!o || o.pin===undefined || o.pin===null || String(o.pin)==='') return null;
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const P = PropertiesService.getScriptProperties(), now = Date.now();
    let st = {n:0, until:0}; try{ st = JSON.parse(P.getProperty('pin_fail')||'') || st; }catch(e){}
    if(st.until > now) return J_({ok:false, err:'locked', wait:Math.ceil((st.until-now)/60000)});   // مقفل حتى لو كان الرمز صحيحاً
    if(o.pin === PIN){ if(st.n){ P.setProperty('pin_fail', JSON.stringify({n:0, until:0})); } return null; }
    st.n = (st.n||0) + 1;
    if(st.n >= PIN_MAX_TRIES){ P.setProperty('pin_fail', JSON.stringify({n:0, until:now+PIN_LOCK_MIN*60000})); return J_({ok:false, err:'locked', wait:PIN_LOCK_MIN}); }
    P.setProperty('pin_fail', JSON.stringify({n:st.n, until:0}));
    return J_({ok:false, err:'pin', left:PIN_MAX_TRIES-st.n});
  } finally { lock.releaseLock(); }
}

function odlJson_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function odlNorm_(p){return String(p||'').replace(/\D/g,'').replace(/^00/,'')}
function odlProps_(){return PropertiesService.getScriptProperties()}
function odlSalt_(){const P=odlProps_();let s=P.getProperty('AUTH_SALT');if(!s){s=Utilities.getUuid();P.setProperty('AUTH_SALT',s)}return s}
function odlHash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,odlSalt_()+s).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('')}
function odlGet_(ph){const v=odlProps_().getProperty('u_'+ph);return v?JSON.parse(v):null}
function odlPut_(ph,o){odlProps_().setProperty('u_'+ph,JSON.stringify(o))}

/* يرجع رقم الهاتف الموثّق من الرمز، أو null (لا جلسة لمن حُذف من التسجيل) */
function odlPhone_(tok){
  const i=String(tok||'').indexOf('.');if(i<8)return null;
  const ph=tok.slice(0,i),u=odlGet_(ph);
  if(!u||!u.tokHash||u.tokExp<Date.now()||u.tokHash!==odlHash_(tok))return null;
  return ph;
}

/* حذف متدرب نهائياً من الخادم: التسجيل والجلسة والرمز واسم الشهادة المعدَّل */
function odlRemoveUser_(ph){
  const P=odlProps_();
  P.deleteProperty('u_'+ph);
  P.deleteProperty('cname_'+ph);
}

/* المعالج الرئيسي: يرجع استجابة JSON أو null إن لم يكن الطلب خاصاً بالدخول */
function odlHandle_(s){
  const a=s&&s.action;if(!a||String(a).indexOf('auth_')!==0)return null;
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    const isAdmin=s.pin===PIN, ph=odlNorm_(s.ph), now=Date.now();
    if(a==='auth_roster'){                       // POST: إضافة/تحديث متدربين في قائمة المسجلين (لا يحذف أحداً؛ الحذف بزر 🗑 فقط)
      if(!isAdmin)return odlJson_({ok:false,err:'pin'});
      let n=0;(JSON.parse(s.rows||'[]')).forEach(r=>{const p=odlNorm_(r.ph);if(!/^\d{8,15}$/.test(p))return;
        const u=odlGet_(p)||{};u.name=String(r.name||u.name||'').slice(0,80);odlPut_(p,u);n++});
      return odlJson_({ok:true,n:n,total:(JSON.parse(s.rows||'[]')).length});
    }
    if(a==='auth_pending'){                      // طلبات بانتظار الرمز
      if(!isAdmin)return odlJson_({ok:false,err:'pin'});
      const all=odlProps_().getProperties(),rows=[];
      Object.keys(all).forEach(k=>{if(k.indexOf('u_')!==0)return;const u=JSON.parse(all[k]);
        if(u.reqAt&&(u.status==='requested'||u.status==='issued'))rows.push({ph:k.slice(2),name:u.name||'',reqAt:u.reqAt,status:u.status})});
      rows.sort((x,y)=>y.reqAt-x.reqAt);return odlJson_({ok:true,rows:rows});
    }
    if(a==='auth_issue'){                        // المدرب يولّد رمزاً
      if(!isAdmin)return odlJson_({ok:false,err:'pin'});
      const u=odlGet_(ph);if(!u)return odlJson_({ok:false,err:'الرقم غير مسجل'});
      const code=String(parseInt(Utilities.getUuid().replace(/-/g,'').slice(0,10),16)%900000+100000);
      u.codeHash=odlHash_(ph+code);u.codeExp=now+ODL_CODE_TTL_MIN*60000;u.tries=0;u.status='issued';odlPut_(ph,u);
      return odlJson_({ok:true,code:code,name:u.name||'',ttl:ODL_CODE_TTL_MIN});
    }
    if(a==='auth_request'){                      // المتدرب يطلب رمزاً
      const u=odlGet_(ph);if(!u)return odlJson_({ok:false,err:'رقمك غير مسجل. تواصل مع المدرب لإضافته.'});
      if(u.reqAt&&now-u.reqAt<60000&&u.status==='requested')return odlJson_({ok:false,err:'تم إرسال طلبك للتو. انتظر دقيقة.'});
      u.reqAt=now;u.status='requested';odlPut_(ph,u);return odlJson_({ok:true});
    }
    if(a==='auth_verify'){                       // المتدرب يتحقق من الرمز
      const u=odlGet_(ph),code=String(s.code||'').replace(/\D/g,'');
      if(!u||!u.codeHash||u.codeExp<now)return odlJson_({ok:false,err:'الرمز غير صالح أو منتهي. اطلب رمزاً جديداً.'});
      if((u.tries||0)>=ODL_MAX_TRIES)return odlJson_({ok:false,err:'تجاوزت عدد المحاولات. اطلب رمزاً جديداً.'});
      if(u.codeHash!==odlHash_(ph+code)){u.tries=(u.tries||0)+1;odlPut_(ph,u);return odlJson_({ok:false,err:'الرمز غير صحيح.'})}
      const tok=ph+'.'+Utilities.getUuid()+Utilities.getUuid();
      u.tokHash=odlHash_(tok);u.tokExp=now+ODL_TOKEN_TTL_H*3600000;u.codeHash='';u.status='active';odlPut_(ph,u);
      return odlJson_({ok:true,tok:tok,ph:ph,name:u.name||''});
    }
    if(a==='auth_check'){                        // فحص الجلسة: من حُذف يتلقى SESSION_INVALID فتطرده الواجهة
      const p=odlPhone_(s.tok);
      if(!p)return odlJson_({ok:false,err:'unauth',code:'SESSION_INVALID'});
      const u=odlGet_(p)||{};
      return odlJson_({ok:true,ph:p,name:u.name||'',left:aiLeftOf_(u)});
    }
    if(a==='auth_users'){                        // المدرب: قائمة المسجلين ومحاولاتهم
      if(!isAdmin)return odlJson_({ok:false,err:'pin'});
      const all=odlProps_().getProperties(),rows=[];
      Object.keys(all).forEach(k=>{if(k.indexOf('u_')!==0)return;const u=JSON.parse(all[k]);
        rows.push({ph:k.slice(2),name:u.name||'',left:aiLeftOf_(u),req:u.aiReq||0})});
      rows.sort((x,y)=>((y.req>0)-(x.req>0))||String(x.name).localeCompare(String(y.name)));
      return odlJson_({ok:true,rows:rows});
    }
    if(a==='auth_delete'){                       // المدرب: حذف متدرب نهائياً (يسقط دخوله فوراً من كل الأجهزة ويختفي من التتبع عند المزامنة التالية)
      if(!isAdmin)return odlJson_({ok:false,err:'pin'});
      if(!odlGet_(ph))return odlJson_({ok:false,err:'الرقم غير مسجل'});
      odlRemoveUser_(ph);return odlJson_({ok:true});
    }
    if(a==='auth_ai_set'){                       // المدرب: تعيين عدد المحاولات لمتدرب
      if(!isAdmin)return odlJson_({ok:false,err:'pin'});
      const u=odlGet_(ph);if(!u)return odlJson_({ok:false,err:'الرقم غير مسجل'});
      u.aiLeft=Math.max(0,Math.min(AI_MAX_SET,parseInt(s.n,10)||0));u.aiDay=aiToday_();u.aiReq=0;odlPut_(ph,u);return odlJson_({ok:true,left:u.aiLeft});
    }
    if(a==='auth_ai_request'){                   // المتدرب: طلب تجديد
      const p=odlPhone_(s.tok);if(!p)return odlJson_({ok:false,err:'unauth',code:'SESSION_INVALID'});
      const u=odlGet_(p)||{};u.aiReq=now;odlPut_(p,u);return odlJson_({ok:true});
    }
    return odlJson_({ok:false,err:'unknown'});
  }finally{lock.releaseLock()}
}

/* ===================== أصوات الشخصيات (Gemini TTS) + نتائج المحاكاة ===================== */
/* مفتاح Gemini: Project Settings ← Script properties ← أضف GEMINI_KEY. الأصوات المولّدة تُحفظ في Drive فلا تُولَّد مرتين. */
const TTS_MODELS = ['gemini-3.1-flash-tts-preview', 'gemini-2.5-flash-preview-tts'];   // يُجرَّب بالترتيب
const TTS_VOICE  = 'Fenrir';          // صوت كل الشخصيات
const TTS_FOLDER = 'ODL_Voice';       // مجلد Google Drive لحفظ الأصوات
const TTS_DAILY  = 400;               // أقصى عدد أصوات جديدة تُولَّد يومياً (المحفوظ مسبقاً لا يُحتسب)
const TTS_MAX_CHARS = 700;
const TTS_STYLE = {                   // نبرة الأداء بحسب الشخصية (الصوت واحد: Fenrir)
  mkt:'in a polite, slightly worried and apologetic tone, like a department head reporting a delay',
  sta:'in a calm, analytical and confident tone',
  acc:'in a hurried, slightly tense tone, like someone under a deadline',
  qa:'in a friendly, efficient and clear tone',
  ctr:'in a formal, courteous and measured tone, like an official academic center',
  cou:'in a firm, questioning tone, like a college council member challenging a decision',
  coach:'in a warm, encouraging and wise mentoring tone'
};
function ttsFolder_(){ const it = DriveApp.getFoldersByName(TTS_FOLDER); return it.hasNext() ? it.next() : DriveApp.createFolder(TTS_FOLDER); }
function sgn_(v){ return v > 127 ? v-256 : v; }
function wavFromPcm_(pcm){          // PCM خام 24kHz أحادي 16-بت ← ملف WAV
  const n = pcm.length, rate = 24000, le32 = x => [x&255,(x>>8)&255,(x>>16)&255,(x>>24)&255].map(sgn_), le16 = x => [x&255,(x>>8)&255].map(sgn_), str = s => s.split('').map(c => c.charCodeAt(0));
  const hdr = [].concat(str('RIFF'), le32(36+n), str('WAVE'), str('fmt '), le32(16), le16(1), le16(1), le32(rate), le32(rate*2), le16(2), le16(16), str('data'), le32(n));
  return hdr.concat(Array.prototype.slice.call(pcm));
}
function ttsGenerate_(key, prompt){
  let last = '';
  for(const m of TTS_MODELS){
    try{
      const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+m+':generateContent', {method:'post', contentType:'application/json', headers:{'x-goog-api-key':key}, muteHttpExceptions:true,
        payload:JSON.stringify({contents:[{parts:[{text:prompt}]}], generationConfig:{responseModalities:['AUDIO'], speechConfig:{voiceConfig:{prebuiltVoiceConfig:{voiceName:TTS_VOICE}}}}})});
      const code = r.getResponseCode(); let j = {}; try{ j = JSON.parse(r.getContentText()||'{}'); }catch(e){}
      const part = j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts && j.candidates[0].content.parts[0];
      if(code === 200 && part && part.inlineData && part.inlineData.data) return part.inlineData.data;
      last = m+': '+code+' '+String((j.error && j.error.message) || '').slice(0,120);
    }catch(e){ last = m+': '+String(e).slice(0,120); }
  }
  throw new Error(last || 'tts_fail');
}
function ttsHandle_(b){
  const coach = b.pin !== undefined && String(b.pin) !== '' && b.pin === PIN;     // رمز المدرب صحيح (pinGate_ عدّ الخاطئ)
  if(!coach && !odlPhone_(b.tok)) return J_({ok:false, err:'unauth', code:'SESSION_INVALID'});
  const ch = TTS_STYLE[b.ch] ? String(b.ch) : 'coach', text = String(b.text||'').replace(/\s+/g,' ').trim().slice(0, TTS_MAX_CHARS);
  if(text.length < 2) return J_({ok:false, err:'empty'});
  const name = 'v_'+odlHash_('tts|'+TTS_VOICE+'|'+ch+'|'+text).slice(0,28)+'.wav', folder = ttsFolder_(), it = folder.getFilesByName(name);
  if(it.hasNext()) return J_({ok:true, cached:true, audio:Utilities.base64Encode(it.next().getBlob().getBytes())});
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if(!key) return J_({ok:false, err:'no_key'});
  const P = PropertiesService.getScriptProperties(), dk = 'tts_'+aiToday_();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{ const n = +(P.getProperty(dk)||0); if(n >= TTS_DAILY) return J_({ok:false, err:'tts_quota'}); P.setProperty(dk, String(n+1)); } finally { lock.releaseLock(); }
  let b64;
  try{ b64 = ttsGenerate_(key, 'Say '+TTS_STYLE[ch]+': '+text); }
  catch(e){ return J_({ok:false, err:'tts_fail', msg:String(e).slice(0,200)}); }
  const wav = wavFromPcm_(Utilities.base64Decode(b64));
  try{ folder.createFile(Utilities.newBlob(wav, 'audio/wav', name)); }catch(e){}
  return J_({ok:true, cached:false, audio:Utilities.base64Encode(wav)});
}

/* ---- نتائج محاكاة العمداء (الجلسة 3): المتدرب يرسل، والمدرب يقرأ ---- */
const SIM_HEADERS = ['ph','name','sim','score','best','tries','persona','comp','kpi','routine','overload','agent','hist','ts','college','mode','depts'];
function simHandle_(b){
  const a = b && b.action; if(a !== 'tts' && a !== 'sim_save' && a !== 'sim_list') return null;
  if(a === 'tts') return ttsHandle_(b);
  if(a === 'sim_list'){
    if(b.pin !== PIN) return J_({ok:false, err:'pin'});
    const s = sheet_('Sim', SIM_HEADERS), v = s.getDataRange().getValues(), rows = [];
    for(let i=1;i<v.length;i++){
      if(!v[i][0]) continue;
      const o = {}; SIM_HEADERS.forEach((k,j) => { o[k] = String(v[i][j]===undefined ? '' : v[i][j]); });
      ['comp','kpi','routine','agent','hist'].forEach(k => { o[k] = jparse_(o[k], k === 'hist' ? [] : null); });
      o.score = +o.score||0; o.best = +o.best||0; o.tries = +o.tries||0; o.overload = o.overload === '1'; o.ts = +o.ts||0;
      rows.push(o);
    }
    rows.sort((x,y) => y.ts - x.ts);
    return J_({ok:true, rows:rows});
  }
  /* sim_save */
  const ph = odlPhone_(b.tok); if(!ph) return J_({ok:false, err:'unauth', code:'SESSION_INVALID'});
  const d = b.data || {}, sim = String(b.sim||'s3').replace(/[^\w-]/g,'').slice(0,12) || 's3';
  const num = (x, lo, hi) => { x = +x; return isFinite(x) ? Math.max(lo, Math.min(hi, Math.round(x*100)/100)) : 0; };
  const comp = {}; Object.keys(d.comp||{}).slice(0,12).forEach(k => { if(/^\w{1,12}$/.test(k)) comp[k] = [num(d.comp[k],0,100), num((d.max||{})[k],0,100)]; });
  const kpi = {dpi:num(d.dpi,0,100), api:num(d.api,0,100), acc:num(d.acc,0,100)};
  const rt = (d.routine && typeof d.routine === 'object') ? {d1:String(d.routine.d1||'').slice(0,40), d2:String(d.routine.d2||'').slice(0,40), m3:String(d.routine.m3||'').slice(0,40), e4:String(d.routine.e4||'').slice(0,40)} : null;
  const ag = (d.agent && typeof d.agent === 'object') ? {src:d.agent.src === 'agents' ? 'agents' : 'local', score:num(d.agent.score,0,12), why:String(d.agent.why||'').slice(0,300)} : null;
  const hist = (Array.isArray(d.hist) ? d.hist : []).slice(0,20).map(x => ({title:String((x||{}).title||'').slice(0,60), earned:num((x||{}).earned,0,100), max:num((x||{}).max,0,100)}));
  const u = odlGet_(ph) || {}, score = num(d.score,0,100), now = Date.now();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const s = sheet_('Sim', SIM_HEADERS), v = s.getDataRange().getValues();
    let at = -1; for(let i=1;i<v.length;i++) if(String(v[i][0]) === ph && String(v[i][2]) === sim){ at = i; break; }
    const tries = (at > 0 ? (+v[at][5]||0) : 0) + 1, best = Math.max(score, at > 0 ? (+v[at][4]||0) : 0);
    const row = [ph, (u.name || String(d.name||'')).slice(0,80), sim, String(score), String(best), String(tries), String(d.persona||'').slice(0,60), JSON.stringify(comp), JSON.stringify(kpi), JSON.stringify(rt), d.overload ? '1' : '', JSON.stringify(ag), JSON.stringify(hist), String(now), String(d.college||'').replace(/[\u0000-\u001f<>]/g,'').slice(0,80), d.mode === 'team' ? 'team' : 'solo', String(d.depts||'').replace(/[\u0000-\u001f<>]/g,'').slice(0,200)];
    if(at > 0) s.getRange(at+1, 1, 1, row.length).setValues([row]); else s.appendRow(row);
    return J_({ok:true, tries:tries, best:best});
  } finally { lock.releaseLock(); }
}

/* شغّل هذه الدالة مرة واحدة من المحرر لمنح الأذونات، ثم أعد النشر بإصدار جديد */
function authorize(){
  SpreadsheetApp.openById(SS_ID).getName();
  DriveApp.getRootFolder().getName();
  UrlFetchApp.fetch('https://openrouter.ai', {muteHttpExceptions:true});
  PropertiesService.getScriptProperties().getKeys();
  CacheService.getScriptCache().get('x');
  ScriptApp.getProjectTriggers();                                   // للمشغّلات الزمنية (Tracking.gs)
  UrlFetchApp.fetch('https://graph.facebook.com', {muteHttpExceptions:true});   // لواتساب (Tracking.gs)
  UrlFetchApp.fetch('https://generativelanguage.googleapis.com', {muteHttpExceptions:true});   // أصوات Gemini TTS
  ttsFolder_();
  LockService.getScriptLock();
  Logger.log('تم منح الأذونات');
}
function testAI(){
  const key = PropertiesService.getScriptProperties().getProperty('OPENROUTER_KEY');
  const r = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    headers:{Authorization:'Bearer '+key},
    payload:JSON.stringify({model:'openrouter/free', messages:[{role:'user', content:'مرحبا'}]})
  });
  Logger.log(r.getResponseCode() + ' ' + r.getContentText().slice(0,500));
}
function testAI2(){
  const P = PropertiesService.getScriptProperties();
  Logger.log('الأسماء: ' + JSON.stringify(P.getKeys()));
  const raw = P.getProperty('OPENROUTER_KEY');
  Logger.log('موجود؟ ' + (raw !== null) + ' | الطول: ' + (raw ? raw.length : 0) + ' | البداية: ' + (raw ? JSON.stringify(raw.slice(0,9)) : '-') + ' | فيه مسافات؟ ' + (raw ? /\s/.test(raw) : '-'));
  const key = String(raw||'').trim();
  const r = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    headers:{Authorization:'Bearer ' + key},
    payload:JSON.stringify({model:'openrouter/free', messages:[{role:'user', content:'مرحبا'}]})
  });
  Logger.log(r.getResponseCode() + ' ' + r.getContentText().slice(0,300));
}
