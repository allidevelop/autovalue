# Карта переменных финального отчета

Этот файл описывает переменные, которые нужно подставлять в Word-отчет оценки. Рабочий подход: один раз конвертировать существующую «козу» из `.doc` в `.docx`, заменить подсвеченные переменные фрагменты на плейсхолдеры и дальше генерировать готовые отчеты из `.docx`-шаблона.

## 1. Формат плейсхолдеров

Рекомендуемый формат:

```text
{{field_name}}
```

Для повторяющихся блоков:

```text
{{comparables_table}}
{{calculation_table}}
{{report_listing_images}}
{{appendix_full_page_screenshots}}
```

## 2. Источники данных

### 2.1. Витяг

Из витяга берем:

- `extract_index_number` - індексний номер витягу;
- `extract_formed_at` - дата і час формування витягу;
- `registry_object_number` - реєстраційний номер об'єкта нерухомого майна;
- `property_right_type` - тип речового права;
- `object_description` - опис об'єкта;
- `object_type` - тип об'єкта;
- `total_area_m2` - загальна площа;
- `living_area_m2` - житлова площа;
- `address_full` - повна адреса;
- `owners_from_extract` - власники з витяга для перевірки.

Полные реквизиты собственников не собираются из витяга как основной источник. Они берутся из отчета/шаблона клиента и сверяются с витягом.

### 2.2. Технический паспорт

Из техпаспорта берем:

- `rooms_count` - количество комнат;
- `technical_passport_inventory_number` - инвентаризационный номер, если нужен;
- `technical_passport_date` - дата техпаспорта, если нужна;
- `technical_passport_issuer` - орган/исполнитель, если нужен;
- `floor_or_level` - этаж/этажность, если указано;
- `explication_rows` - строки экспликации помещений.

Правило для комнат:

```text
rooms_count = количество строк в экспликации, где в столбце 6 заполнена жилая площадь
```

### 2.3. Excel-расчет

Из Excel-расчета берем:

- `market_value_uah` - итоговая рыночная стоимость, грн;
- `market_value_uah_rounded` - округленная стоимость, грн;
- `market_value_uah_words` - сумма прописью;
- `median_price_usd_m2` - медианное значение, $/кв.м;
- `average_price_usd_m2` - среднее значение, $/кв.м;
- `nbu_rate` - курс НБУ, если используется в тексте;
- `valuation_area_m2` - площадь объекта оценки для расчета;
- `comparables_table` - таблица объектов сравнения;
- `calculation_table` - расчетная таблица сравнительного подхода.

Строка `Корегування на торг` на этапе MVP не заполняется автоматически и должна сохраняться из шаблона/ручного ввода.

### 2.4. Аналоги и объявления

По каждому аналогу:

- `comparable_n_address`;
- `comparable_n_area_m2`;
- `comparable_n_floor_or_level`;
- `comparable_n_price_usd`;
- `comparable_n_price_usd_m2`;
- `comparable_n_location_quality`;
- `comparable_n_building_class`;
- `comparable_n_condition`;
- `comparable_n_delivery_date`;
- `comparable_n_source_url`;
- `comparable_n_full_page_screenshot`;
- `comparable_n_report_image`.

Для архива сохраняется full-page screenshot. В Word-отчет вставляется отдельная оптимизированная версия `report_image`: читабельная и с меньшим весом.

## 3. Основные плейсхолдеры отчета

### 3.1. Титульная часть

- `{{property_type_text}}` - например `однокімнатної квартири`;
- `{{rooms_text}}` - например `однокімнатної`;
- `{{apartment_number}}` - номер квартиры;
- `{{total_area_m2}}` - общая площадь;
- `{{address_city}}` - город;
- `{{address_street}}` - улица;
- `{{address_building}}` - дом;
- `{{address_apartment}}` - квартира;
- `{{address_full}}` - полная адреса одной строкой;
- `{{report_city}}` - город составления отчета;
- `{{report_year}}` - год отчета.

### 3.2. Стороны и документы

- `{{customer_full_name}}` - заказчик оценки;
- `{{owner_full_name}}` - собственник с полными реквизитами из отчета/шаблона;
- `{{owners_from_extract}}` - собственники из витяга для сверки;
- `{{ownership_document_text}}` - правоустанавливающий документ;
- `{{extract_index_number}}` - номер витяга;
- `{{extract_date}}` - дата витяга;
- `{{valuation_purpose}}` - цель оценки;
- `{{valuation_base}}` - база оценки;
- `{{valuation_currency}}` - валюта оценки.

### 3.3. Даты

- `{{valuation_date}}` - дата оценки;
- `{{report_date}}` - дата составления отчета;
- `{{extract_formed_at}}` - дата/время формирования витяга.

### 3.4. Объект оценки

- `{{object_type}}`;
- `{{object_description}}`;
- `{{rooms_count}}`;
- `{{total_area_m2}}`;
- `{{living_area_m2}}`;
- `{{floor_or_level}}`;
- `{{complex_name}}`;
- `{{building_class}}`;
- `{{condition}}`;
- `{{wall_material}}`;
- `{{delivery_date}}`.

### 3.5. Расчет и вывод

- `{{comparables_table}}`;
- `{{calculation_table}}`;
- `{{median_price_usd_m2}}`;
- `{{average_price_usd_m2}}`;
- `{{market_value_uah}}`;
- `{{market_value_uah_rounded}}`;
- `{{market_value_uah_words}}`;
- `{{valuation_conclusion_text}}`.

### 3.6. Приложения

- `{{report_listing_images}}` - читабельные изображения объявлений для отчета;
- `{{appendix_full_page_screenshots}}` - full-page screenshots для доказательной базы, если решим включать их в приложение;
- `{{extract_pages}}` - страницы витяга;
- `{{technical_passport_pages}}` - страницы техпаспорта.

## 4. Правила генерации Word-отчета

- Основной рабочий шаблон должен быть `.docx`.
- Исходные `.doc`-файлы используются как образцы и могут быть конвертированы в `.docx`.
- Подстановка должна идти по плейсхолдерам, а не по поиску подсвеченного текста.
- Перед финальной выдачей отчет должен сохраняться в папку периода/месяца.
- PDF-экспорт не обязателен для MVP.
- Изображения объявлений должны быть читабельными, но оптимизированными по весу.

