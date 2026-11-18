"""Generate the lab report (optional authoring script; not needed to run dashboard)."""
from __future__ import annotations
import csv
from collections import Counter
from pathlib import Path
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / 'data' / 'raw'
FIG = ROOT / 'figures'
REPORTS = ROOT / 'reports'
REPORTS.mkdir(exist_ok=True)
FIG.mkdir(exist_ok=True)

def read(name):
    with (RAW/name).open(encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), fill); tcPr.append(shd)

def set_cell_text(cell, text, bold=False, color='222222', size=9):
    cell.text=''; p=cell.paragraphs[0]; p.paragraph_format.space_after=Pt(2)
    r=p.add_run(str(text)); r.bold=bold; r.font.name='Arial'; r.font.size=Pt(size); r.font.color.rgb=RGBColor.from_string(color)
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def add_table(doc, headers, rows, widths=None):
    t=doc.add_table(rows=1, cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
    for i,h in enumerate(headers):
        set_cell_text(t.rows[0].cells[i],h,True,'FFFFFF',8); set_cell_shading(t.rows[0].cells[i],'263B49')
    for ridx,row in enumerate(rows):
        cells=t.add_row().cells
        for i,val in enumerate(row):
            set_cell_text(cells[i],val,False,'222222',8)
            if ridx%2==1:set_cell_shading(cells[i],'EEF2F4')
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths): row.cells[i].width=Cm(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(1)
    return t

def add_bullet(doc, text):
    p=doc.add_paragraph(style='List Bullet'); p.paragraph_format.space_after=Pt(2); p.add_run(text)
    return p

def add_heading(doc,text,level=1):
    p=doc.add_heading(text,level=level); p.paragraph_format.space_before=Pt(7); p.paragraph_format.space_after=Pt(4); return p

def add_page_field(paragraph):
    r=paragraph.add_run('Страница '); r.font.name='Arial'; r.font.size=Pt(8)
    fld=OxmlElement('w:fldSimple'); fld.set(qn('w:instr'),'PAGE'); paragraph._p.append(fld)

def prepare_figures(deaths,preds,battles):
    plt.rcParams.update({'font.family':'DejaVu Sans','font.size':9,'axes.titlesize':12,'axes.labelsize':9})
    # Death year distribution including missing values as explicit category.
    dcounts=Counter((r['Death Year'].strip() or 'Нет года') for r in deaths)
    order=['297','298','299','300','Нет года']
    fig,ax=plt.subplots(figsize=(7.4,3.0)); bars=ax.bar(order,[dcounts[x] for x in order],color=['#D8B46A','#69C6B5','#E28E7D','#A69ADC','#77A8D5'])
    ax.set_title('Распределение записей по году смерти'); ax.set_ylabel('Количество записей'); ax.spines[['top','right']].set_visible(False); ax.grid(axis='y',alpha=.2)
    ax.bar_label(bars,padding=3); fig.tight_layout(); fig.savefig(FIG/'report_deaths_year.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # Confusion matrix using actual and pred values from raw source.
    labels=['actual 0','actual 1']; m=[[0,0],[0,0]]
    for r in preds:
        if r['actual'] in ('0','1') and r['pred'] in ('0','1'): m[int(r['actual'])][int(r['pred'])]+=1
    fig,ax=plt.subplots(figsize=(5.2,3.3)); im=ax.imshow(m,cmap='YlGnBu')
    ax.set_xticks([0,1],labels=['pred 0','pred 1']); ax.set_yticks([0,1],labels=labels); ax.set_title('Матрица actual × pred')
    for i in range(2):
        for j in range(2): ax.text(j,i,str(m[i][j]),ha='center',va='center',fontsize=12,fontweight='bold',color='black')
    fig.colorbar(im,ax=ax,fraction=.046,pad=.04,label='Количество'); fig.tight_layout(); fig.savefig(FIG/'report_prediction_matrix.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # Battles by region.
    bcounts=Counter(r['region'].strip() or 'Не указано' for r in battles); ordered=bcounts.most_common()
    fig,ax=plt.subplots(figsize=(7.4,3.4)); labels=[x[0] for x in ordered][::-1]; vals=[x[1] for x in ordered][::-1]
    bars=ax.barh(labels,vals,color='#69AFA5'); ax.set_title('Распределение сражений по регионам'); ax.set_xlabel('Количество сражений'); ax.spines[['top','right']].set_visible(False); ax.grid(axis='x',alpha=.2); ax.bar_label(bars,padding=3); fig.tight_layout(); fig.savefig(FIG/'report_battles_region.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # Battle types.
    tcounts=Counter(r['battle_type'].strip() or 'Не указано' for r in battles)
    fig,ax=plt.subplots(figsize=(5.6,3.0)); items=tcounts.most_common(); bars=ax.bar([x[0] for x in items],[x[1] for x in items],color=['#D8B46A','#69C6B5','#E28E7D','#A69ADC','#77A8D5'][:len(items)])
    ax.set_title('Сражения по типам'); ax.set_ylabel('Количество'); ax.spines[['top','right']].set_visible(False); ax.grid(axis='y',alpha=.2); ax.bar_label(bars,padding=3); fig.tight_layout(); fig.savefig(FIG/'report_battles_type.png',dpi=180,bbox_inches='tight'); plt.close(fig)

def build():
    deaths=read('character-deaths.csv'); preds=read('character-predictions.csv'); battles=read('battles.csv')
    prepare_figures(deaths,preds,battles)
    conf=Counter((r['actual'],r['pred']) for r in preds)
    total=sum(conf.values()); correct=sum(v for (a,b),v in conf.items() if a==b)
    years=Counter(r['Death Year'].strip() or 'Нет года' for r in deaths)
    outcome=Counter(r['attacker_outcome'].strip() or 'Не указано' for r in battles)
    regions=Counter(r['region'].strip() or 'Не указано' for r in battles)
    types=Counter(r['battle_type'].strip() or 'Не указано' for r in battles)
    known_year=sum(bool(r['Death Year'].strip()) for r in deaths)
    known_book=sum(bool(r['Book of Death'].strip()) for r in deaths)
    both_sizes=sum(bool(r['attacker_size'].strip()) and bool(r['defender_size'].strip()) for r in battles)
    doc=Document(); sec=doc.sections[0]
    sec.top_margin=Cm(1.8); sec.bottom_margin=Cm(1.7); sec.left_margin=Cm(2.1); sec.right_margin=Cm(2.1)
    styles=doc.styles
    styles['Normal'].font.name='Times New Roman'; styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman'); styles['Normal'].font.size=Pt(11)
    styles['Normal'].paragraph_format.space_after=Pt(5); styles['Normal'].paragraph_format.line_spacing=1.08
    for stname,size,color in [('Title',24,'263B49'),('Heading 1',16,'263B49'),('Heading 2',12,'6A5430')]:
        st=styles[stname]; st.font.name='Arial'; st._element.rPr.rFonts.set(qn('w:eastAsia'),'Arial'); st.font.size=Pt(size); st.font.bold=True; st.font.color.rgb=RGBColor.from_string(color)
    header=sec.header.paragraphs[0]; header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    rr=header.add_run('Лабораторная работа № 10 · Интерактивные дашборды'); rr.font.name='Arial'; rr.font.size=Pt(8); rr.font.color.rgb=RGBColor(110,120,128)
    footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_page_field(footer)
    # Cover page
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(45)
    r=p.add_run('ЛАБОРАТОРНАЯ РАБОТА № 10'); r.bold=True; r.font.name='Arial'; r.font.size=Pt(18)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(20)
    r=p.add_run('ИНТЕРАКТИВНЫЕ ДАШБОРДЫ'); r.bold=True; r.font.name='Arial'; r.font.size=Pt(25); r.font.color.rgb=RGBColor(38,59,73)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('Анализ данных вселенной «Игры престолов»'); r.italic=True; r.font.size=Pt(14)
    for _ in range(4): doc.add_paragraph('')
    add_table(doc,['Поле','Значение'],[
        ['Выполнил(а)','____________________________________________'],['Группа','____________________________________________'],['Проверил(а)','____________________________________________'],['Образовательная организация','____________________________________________']],[5.2,11.0])
    for _ in range(3): doc.add_paragraph('')
    p=doc.add_paragraph('2026'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.runs[0].font.size=Pt(14)
    doc.add_page_break()
    # Page 2
    add_heading(doc,'Содержание',1)
    for line in ['1. Цель и задачи работы','2. Исходные данные','3. Подготовка данных','4. Концепция и структура дашборда','5. Фильтры, KPI и визуализации','6. Результаты анализа','7. Тестирование и оптимизация','8. Выводы','Источники']:
        p=doc.add_paragraph(line); p.paragraph_format.space_after=Pt(2)
    add_heading(doc,'1. Цель и задачи работы',1)
    doc.add_paragraph('Цель лабораторной работы — разработать интерактивный дашборд, позволяющий исследовать наборы данных с помощью фильтров, KPI, таблиц и динамических визуализаций. Реализация адаптирована к фактическому составу трех CSV-файлов из набора Game of Thrones.')
    for x in ['изучить структуру исходных CSV и типы полей;','проверить пропуски, повторяющиеся записи и возможность объединения таблиц;','подготовить вычисляемые представления для анализа персонажей, прогнозов и сражений;','создать интерактивные фильтры, KPI и графики;','проверить пересчет показателей, экспорт данных и работу интерфейса.']:
        add_bullet(doc,x)
    add_heading(doc,'2. Исходные данные',1)
    add_table(doc,['Файл','Размер','Содержание'],[
        ['character-deaths.csv','917 строк × 13 столбцов','Имя, принадлежность, год/книга/глава смерти, Gender, Nobility, признаки появления в книгах'],
        ['character-predictions.csv','1 946 строк × 33 столбца','actual, pred, alive, isAlive, имя, дом, культура, возраст, популярность и родственные связи'],
        ['battles.csv','38 строк × 25 столбцов','Год, название и тип сражения, атакующие/защитники, исход, место и регион']
    ],[4.2,3.2,9.0])
    doc.add_paragraph('Всего анализируются 2 901 строка в трех исходных таблицах. Таблица смертей содержит 916 уникальных имен: имя Myles встречается дважды с разными атрибутами и поэтому не удаляется как дубликат имени.')
    doc.add_page_break()
    # Page 3
    add_heading(doc,'3. Подготовка и проверка данных',1)
    doc.add_paragraph('Для сборки автономной страницы используется отдельный скрипт scripts/build_dashboard.py на стандартной библиотеке Python. Скрипт читает CSV в кодировке UTF-8 с BOM, проверяет наличие обязательных столбцов, убирает только полностью совпадающие строки и встраивает записи в JSON внутри index.html. Готовый HTML не требует Python при запуске.')
    add_table(doc,['Проверка','Результат','Как обработано'],[
        ['Строки смертей с заполненным Death Year',f'{known_year} из {len(deaths)}','612 записей без года оставлены отдельной категорией'],
        ['Заполненный Book of Death',f'{known_book} из {len(deaths)}','Пропуски не заменены нулевым номером книги'],
        ['actual и pred',f'{total} из {len(preds)}','Обе колонки заполнены; расчет совпадения возможен для всех строк'],
        ['Известен исход сражения',f'{sum(outcome[x] for x in ["win","loss"])} из {len(battles)}','Одна запись с неизвестным исходом показана отдельно'],
        ['Оценки численности обеих сторон',f'{both_sizes} из {len(battles)}','Scatter plot строится только по строкам с двумя числами']
    ],[5.0,3.0,8.4])
    doc.add_paragraph('В character-deaths.csv отсутствует поле метода смерти, а battles.csv не содержит координат. Поэтому фильтр типа смерти и географическая карта не реализуются: вместо них используются доступные поля книги, принадлежности, региона и места. Номер книги не подменяется сезоном сериала.')
    add_heading(doc,'4. Концепция и структура дашборда',1)
    doc.add_paragraph('Создан тематический аналитический интерфейс «Летопись Вестероса» с тремя независимыми разделами: смертность персонажей, прогнозы классификации и сражения. Четвертый раздел описывает методику, качество и ограничения данных.')
    doc.add_paragraph('Техническая архитектура: один HTML-файл со встроенными CSV-данными в формате JSON; HTML отвечает за структуру, CSS — за оформление и адаптивную сетку, JavaScript — за фильтрацию и пересчет показателей, SVG — за отрисовку графиков. Сторонние библиотеки, CDN, сервер и сетевое подключение не используются.')
    if (FIG/'report_dashboard_top.png').exists():
        doc.add_picture(str(FIG/'report_dashboard_top.png'),width=Cm(16.0))
        p=doc.add_paragraph('Рисунок 1 — Интерфейс дашборда: навигация, фильтры, KPI и визуализации'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    # Section 5 follows the interface figure on the same page where space allows.
    # Page 4
    add_heading(doc,'5. Фильтры, KPI и визуализации',1)
    add_table(doc,['Раздел','Фильтры','Ключевые KPI'],[
        ['Смертность','Год смерти, принадлежность, книга смерти, Gender, Nobility','Число записей; строки с годом; средний год смерти; самая частая принадлежность'],
        ['Прогнозы','actual, pred, дом, культура, male, isNoble','Число профилей; число сравнимых строк; доля совпадений; среднее поле alive'],
        ['Сражения','Год, регион, тип, первый атакующий дом, исход, место','Число сражений; известные исходы; доля побед атакующих; крупнейший по числу записей регион']
    ],[2.4,7.2,6.8])
    doc.add_paragraph('В каждом разделе элементы управления поддерживают множественный выбор. Реализованы каскадные списки: год ограничивает список домов/регионов, выбранная категория уточняет последующие доступные значения. После изменения фильтра выборка пересчитывается в JavaScript, вслед за ней обновляются KPI, диаграммы и таблица. Кнопка «Сбросить фильтры» возвращает исходное состояние; кнопка экспорта сохраняет текущую выборку в CSV.')
    add_heading(doc,'Реализованные типы визуализаций',2)
    for x in ['столбчатые и горизонтальные диаграммы распределения;','кольцевая диаграмма долей категорий;','линейный график последовательности / накопительного количества;','тепловая матрица категорий;','точечная диаграмма числовых полей;','гистограмма числового показателя;','таблица с ограниченным просмотром строк и CSV-экспортом.']:
        add_bullet(doc,x)
    add_heading(doc,'Примечание к фильтрам',2)
    doc.add_paragraph('Категории «Не указано» и «Нет данных» сохраняют смысл пропусков исходного файла. Пустой список в многострочном фильтре трактуется как отсутствие ограничения. Фильтры внутри каждого тематического раздела действуют на все KPI и графики этого раздела.')
    # Page 5 results
    add_heading(doc,'6. Результаты анализа',1)
    doc.add_paragraph(f'В таблице смертей заполненный год имеется у {known_year} строк (33,3%). Чаще всего встречается 299 год — {years["299"]} записей; без года в источнике остаются {years["Нет года"]} строк. Это существенно влияет на полноту временного анализа и явно показано отдельной категорией.')
    doc.add_picture(str(FIG/'report_deaths_year.png'),width=Cm(14.5)); p=doc.add_paragraph('Рисунок 2 — Распределение записей по году смерти'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f'В таблице прогнозов совпадение actual и pred составляет {correct} из {total} строк ({correct/total*100:.1f}%). Наибольшая ячейка матрицы — actual = 1 и pred = 1: {conf[("1","1")]} записей. Доля совпадений является описательной метрикой на предоставленной таблице и не заменяет независимую оценку качества модели.')
    doc.add_picture(str(FIG/'report_prediction_matrix.png'),width=Cm(10.5)); p=doc.add_paragraph('Рисунок 3 — Матрица actual × pred'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    # Page 6
    add_heading(doc,'6. Результаты анализа (продолжение)',1)
    doc.add_paragraph(f'В таблице сражений 38 записей. Больше всего сражений отмечено в регионе {regions.most_common(1)[0][0]} ({regions.most_common(1)[0][1]} записей), затем в регионе {regions.most_common(2)[1][0]} ({regions.most_common(2)[1][1]}). У 37 сражений известен исход: {outcome["win"]} побед и {outcome["loss"]} поражений атакующей стороны. У 16 сражений заполнены обе оценки численности сторон.')
    doc.add_picture(str(FIG/'report_battles_region.png'),width=Cm(14.0)); p=doc.add_paragraph('Рисунок 4 — Распределение сражений по регионам'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    add_table(doc,['Показатель','Значение'],[
        ['Всего сражений',len(battles)],['Известен исход',37],['Победа атакующей стороны',f'{outcome["win"]} ({outcome["win"]/37*100:.1f}%)'],['Полевые сражения',types['pitched battle']],['Осады',types['siege']],['Засады',types['ambush']],['Разорения',types['razing']]
    ],[7.3,7.3])
    # Testing and conclusions continue after the battle summary where space allows.
    # Page 7
    add_heading(doc,'7. Тестирование и оптимизация',1)
    add_paragraph=doc.add_paragraph
    add_paragraph('Проверены структура и формирование файла HTML, наличие встроенного JSON, чтение трех CSV и сохранение контрольных размеров таблиц. Автоматические тесты стандартной библиотеки Python завершились успешно: 3 из 3.')
    add_table(doc,['Проверка интерфейса','Ожидаемое поведение','Результат'],[
        ['Открытие HTML','Страница загружается без внешних ресурсов','Успешно'],
        ['Фильтр года в смертности','KPI и графики перестраиваются','917 → 156 записей для года 299'],
        ['Каскадный список','Год ограничивает доступные дома / регионы','Проверено: 297 AC → Night’s Watch; 298 AC → 2 региона'],
        ['Фильтр региона в сражениях','Остаются записи указанного региона','38 → 17 записей The Riverlands'],
        ['Фильтр actual = 0','Обновляется таблица и метрики прогноза','1 946 → 495 профилей'],
        ['Переключение вкладок','Разделы отображаются без ошибок JavaScript','Успешно'],
        ['Сброс фильтров','Возвращается полная выборка','Успешно'],
        ['Экспорт CSV','Скачивается текущая фильтрованная таблица','Проверена команда экспорта']
    ],[4.0,6.4,4.8])
    doc.add_paragraph('Для оптимизации используется статический HTML с локально встроенными данными, поэтому отсутствуют повторные запросы к серверу и загрузка библиотек. Таблицы ограничивают количество отображаемых строк, а исходная CSV-выборка остается доступной через экспорт.')
    add_heading(doc,'8. Выводы',1)
    doc.add_paragraph('Разработан автономный интерактивный дашборд по трем наборам данных Game of Thrones. Реализованы тематические разделы с фильтрами, динамическими KPI, несколькими типами визуализаций, таблицами и экспортом. Файл открывается в браузере напрямую и не зависит от установки Python-пакетов, что устраняет ошибки запуска, связанные с отсутствующими Streamlit и Plotly.')
    doc.add_paragraph('Работа показала, что качество интерактивной аналитики зависит от реального состава источника: значительная доля записей смертей не имеет года или книги, а географические координаты и метод смерти отсутствуют. Эти ограничения отражены в интерфейсе и не маскируются искусственными значениями.')
    add_heading(doc,'Источники',1)
    doc.add_paragraph('1. Методические рекомендации для выполнения лабораторной работы № 10 «Интерактивные дашборды», предоставленные в задании.')
    doc.add_paragraph('2. character-deaths.csv, character-predictions.csv, battles.csv — исходные наборы данных, предоставленные для выполнения лабораторной работы.')
    doc.add_paragraph('3. Документация HTML, CSS и JavaScript (стандартные браузерные API); внешние runtime-зависимости в готовом дашборде отсутствуют.')
    out=REPORTS/'Лабораторная_работа_10_Интерактивные_дашборды.docx'; doc.save(out); print(out)

if __name__=='__main__':
    build()
