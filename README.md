# msoffice-transliterator

Microsoft ofis hujjatlari kontentini lotin va kiril alifbolari uchun transliteratsiya qilish

```php
<?php
include 'helper.php';
$alphabet = 'cyr';

//Ms Word
transliterate_docx('fayl.docx', $alphabet);

//Ms Excel
transliterate_xlsx('fayl.xlsx', $alphabet);

//Ms powerpoint
transliterate_pptx('fayl.pptx', $alphabet);

//Epub
transliterate_pptx('fayl.epub', $alphabet);

//HTML
transliterate_html('fayl.html', $alphabet);
```
