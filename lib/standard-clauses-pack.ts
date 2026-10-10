import { NotaryClause, STRICT_FONT_FAMILY, STRICT_FONT_SIZE_PT } from './types';

const pStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;

export const OPTIONAL_STANDARD_CLAUSES_PACK: NotaryClause[] = [
  {
    id: 'std_clause_sale_designation',
    title: 'تعيين العقار المبيع',
    category: 'بيع عقار',
    enabled: true,
    order: 1,
    updatedAt: '2026-01-01T00:00:00.000Z',
    contentHtml: `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># تعيين العقار المبيع</span></p><p dir="rtl" style="${pStyle}">باع وأسقط وتنازل بموجب هذا العقد وبكافة الضمانات الفعلية والقانونية السيد(ة) <span class="smart-tag" data-var="البائع_1_الاسم">[البائع_1_الاسم]</span> لفائدة المشتري السيد(ة) <span class="smart-tag" data-var="المشتري_1_الاسم">[المشتري_1_الاسم]</span> القابل لذلك، العقار المعين أدناه:</p><p dir="rtl" style="${pStyle}">الحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> المتمثلة في <span class="smart-tag" data-var="طبيعة_الحصة">[طبيعة_الحصة]</span> الكائنة بـ <span class="smart-tag" data-var="الطابق">[الطابق]</span>، مساحتها <span class="smart-tag" data-var="المساحة">[المساحة]</span> م²، مع <span class="smart-tag" data-var="الأجزاء_المشتركة">[الأجزاء_المشتركة]</span> من الأجزاء المشتركة، وفق التعيين الكامل: <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span>.</p>`,
  },
  {
    id: 'std_clause_origin_ownership',
    title: 'أصل الملكية',
    category: 'بيع عقار',
    enabled: true,
    order: 2,
    updatedAt: '2026-01-01T00:00:00.000Z',
    contentHtml: `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># أصل الملكية</span></p><p dir="rtl" style="${pStyle}">إن العقار المبيع والمعين أعلاه يملكه البائع <span class="smart-tag" data-var="البائع_1_الاسم">[البائع_1_الاسم]</span> ملكية تامة بموجب عقد رسمي مشهر بالمحافظة العقارية تحت مراجع السند والوصف التقسيمي: <span class="smart-tag" data-var="مراجع_الوصف_التقسيمي">[مراجع_الوصف_التقسيمي]</span>.</p>`,
  },
  {
    id: 'std_clause_ownership_enjoyment',
    title: 'الملكية والانتفاع',
    category: 'بيع عقار',
    enabled: true,
    order: 3,
    updatedAt: '2026-01-01T00:00:00.000Z',
    contentHtml: `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># الملكية والانتفاع</span></p><p dir="rtl" style="${pStyle}">يصبح المشتري <span class="smart-tag" data-var="المشتري_1_الاسم">[المشتري_1_الاسم]</span> مالكاً للعقار المبيع ابتداءً من يوم إشهار هذا العقد بالمحافظة العقارية المختصة، وله حق الانتفاع والحيازة الفعلية والقانونية ابتداءً من تاريخ اليوم خالياً من كل شاغل أو رهن أو امتياز.</p>`,
  },
  {
    id: 'std_clause_price_payment',
    title: 'الثمن وكيفية الدفع',
    category: 'الشروط المالية',
    enabled: true,
    order: 4,
    updatedAt: '2026-01-01T00:00:00.000Z',
    contentHtml: `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># الثمن وكيفية الدفع</span></p><p dir="rtl" style="${pStyle}">تم هذا البيع ورضي به الطرفان مقابل ثمن إجمالي وجزافي قدره بالحروف <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span> دج)، دفعه المشتري للبائع عبر محاسبة الموثق الموقع أسفله، وبموجبه يمنحه البائع مخالصة تامة ونهائية وإبراءً شاملاً بالثمن.</p>`,
  },
  {
    id: 'std_clause_taxes_fees',
    title: 'الرسوم والمصاريف والتصاريح الجبائية',
    category: 'التصاريح الجبائية',
    enabled: true,
    order: 5,
    updatedAt: '2026-01-01T00:00:00.000Z',
    contentHtml: `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># الرسوم والمصاريف والتصاريح الجبائية</span></p><p dir="rtl" style="${pStyle}">يتحمل المشتري جميع مصاريف وحقوق التسجيل والطابع والشهر العقاري وأتعاب التوثيق المترتبة على هذا العقد. وقد ذكّر الموثق الطرفين بأحكام قانون التسجيل وقانون الضرائب المباشرة والرسوم المماثلة المتعلقة بصدق الثمن المصرح به والعقوبات المقررة في حالة إخفاء جزء من الثمن.</p>`,
  },
  {
    id: 'std_clause_power_of_attorney',
    title: 'موضوع الوكالة وحدود التفويض',
    category: 'وكالة',
    enabled: true,
    order: 6,
    updatedAt: '2026-01-01T00:00:00.000Z',
    contentHtml: `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># موضوع الوكالة وحدود التفويض</span></p><p dir="rtl" style="${pStyle}">وكّل الموكل السيد(ة) <span class="smart-tag" data-var="الموكل_1_الاسم">[الموكل_1_الاسم]</span> وأقام مقامه وعوضاً عن شخصه الوكيل السيد(ة) <span class="smart-tag" data-var="الوكيل_1_الاسم">[الوكيل_1_الاسم]</span> للقيام باسمه ولحسابه بـ <span class="smart-tag" data-var="موضوع_الوكالة">[موضوع_الوكالة]</span>، والتوقيع على كافة العقود والمحاضر والوثائق الإدارية اللازمة لذلك.</p>`,
  },
  {
    id: 'std_clause_donation_acceptance',
    title: 'الإيجاب والقبول في الهبة',
    category: 'هبة',
    enabled: true,
    order: 7,
    updatedAt: '2026-01-01T00:00:00.000Z',
    contentHtml: `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># الإيجاب والقبول في الهبة</span></p><p dir="rtl" style="${pStyle}">وهب الواهب السيد(ة) <span class="smart-tag" data-var="الواهب_1_الاسم">[الواهب_1_الاسم]</span> هبة صريحة لا رجعة فيها ودون عوض لفائدة الموهوب له السيد(ة) <span class="smart-tag" data-var="الموهوب_1_الاسم">[الموهوب_1_الاسم]</span> الحاضر في مجلس العقد والقابل لهذه الهبة صراحةً، العقار المعين أعلاه والمقوم مالياً لأجل التسجيل بمبلغ <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span> دج).</p>`,
  },
];
