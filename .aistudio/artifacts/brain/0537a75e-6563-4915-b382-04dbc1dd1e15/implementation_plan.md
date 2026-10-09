# خطة الإصلاح الجذري لاستبدال الوسوم وعرض النص العربي RTL (الإصدار 2.6.1)

تتضمن هذه الخطة التطبيق الحرفي والكامل لإصلاح **Bug #1 (Stale Closure)** وتقوية `mergePlaceholdersIntoHtml` عبر `DOMParser` مع ضمان تغليف الوسوم المكتوبة يدوياً (`decorateSmartTagsInDOM`) قبل الاستبدال وأثناء الكتابة.

---

## 1. تحديث `components/SmartModals.tsx` و `components/SidebarWorkspace.tsx`
- تغيير تعريف الـ prop إلى تمرير صريح للقيم:
  ```ts
  onBakeAllIntoDocument: (values: Record<string, string>) => void;
  ```
- عند الضغط على زر **«استبدال الوسوم نهائياً داخل نص العقد»** (أو «دمج الكل في النص» في الشريط الجانبي)، يتم جمع القيم الحالية من `fieldValues` ومن عناصر الإدخال الحية (`inputRefs.current`) وتمريرها صراحةً إلى `onBakeAllIntoDocument(mergedValues)`.

---

## 2. تحديث `components/NotaryEditorApp.tsx`
- مزامنة `fieldValuesRef` تلقائياً وفورياً:
  ```ts
  const fieldValuesRef = useRef<Record<string, string>>({});
  useEffect(() => {
    fieldValuesRef.current = fieldValues;
  }, [fieldValues]);
  ```
- استدعاء `decorateSmartTagsInDOM(bodyEditorRef.current)` عند كتابة `{{...}}` يدوياً (عند `onBlur` أو فور إغلاق `}}` وعند فتح نافذة الاستمارة وقبل تنفيذ الاستبدال) لضمان تغليف كل وسم بـ `<span class="smart-tag" data-var="...">`.
- تحديث دالة `handleBakeAllPlaceholdersIntoDocument(values?: Record<string, string>)` لتعتمد على `const activeValues = values ?? fieldValuesRef.current ?? fieldValues;` وتمررها إلى `mergePlaceholdersIntoHtml` لكل من `bodyEditorRef` و`headerEditorRef` و`footerEditorRef`.

---

## 3. تقوية `mergePlaceholdersIntoHtml` في `lib/docx-engine.ts` عبر `DOMParser`
- استخدام `DOMParser` لاستبدال جميع عناصر `.smart-tag, .smart-placeholder, [data-var]` بدقة تامة حتى بعد تطبيع خصائص `style`، متبوعاً باستبدال أي `{{...}}` نصي متبقٍ:
  ```ts
  export function mergePlaceholdersIntoHtml(
    html: string,
    fieldValues: Record<string, string> = {}
  ): string {
    if (!html) return '';
    if (typeof DOMParser !== 'undefined') {
      const parser = new DOMParser();
      const doc = parser.parseFromString(`<div>${html}</div>`, 'text/html');
      const root = doc.body.firstElementChild as HTMLElement | null;
      if (root) {
        root.querySelectorAll('.smart-tag, .smart-placeholder, [data-var]').forEach((el) => {
          const key = (
            el.getAttribute('data-var') ||
            (el.textContent || '').replace(/[{}]/g, '')
          ).trim();
          const val = fieldValues[key] ?? fieldValues[key.replace(/\s+/g, '_')];
          if (val !== undefined && val.trim() !== '') {
            el.replaceWith(doc.createTextNode(val.trim()));
          }
        });
        let out = root.innerHTML;
        out = out.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (full, rawKey) => {
          const key = String(rawKey).trim();
          const val = fieldValues[key] ?? fieldValues[key.replace(/\s+/g, '_')];
          return val !== undefined && val.trim() !== '' ? escapeXml(val.trim()) : full;
        });
        return out;
      }
    }
    // Fallback regex
    ...
  }
  ```

---

## 4. إصلاح عرض النص العربي (RTL Bidi) وتغليف الوسوم اليدوية
- إزالة `contenteditable="false"` و `user-select: all` من `.smart-tag` وإضافة `unicode-bidi: isolate` لمنع انعكاس ترتيب الكلمات العربية بصرياً في متصفح Chrome.
- توحيد سلوك زر `{{ }}` وزر `[ ] تحويل المحدد لوسم` ليحوّل أي نص محدد فوراً في مكانه إلى وسم ذكي.
