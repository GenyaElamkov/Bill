"""
Генерируем документы (накладные) согласно шаблонам:
- report_tpl.docx - рапорт,
- statement_tpl.docx - ведомость.

Считываем шаблон, вносим данные и выводим новые документы.
Планируется:
- GUI для удобного ввода данных.
- распечатывать документы.

"""
from pathlib import Path
from docxtpl import DocxTemplate


def checking_templates():
    """Проверяем наличие шаблонов в директории template."""
    template = ['report_tpl.docx', 'statement_tpl.docx']

    path = Path('./template/')
    all_files = list(path.iterdir())

    if not all_files:
        raise FileNotFoundError('[!] Отсутствуют шаблоны.')

    all_files = [p.name for p in all_files]

    if len(set(all_files).intersection(template)) < 2:
        raise FileNotFoundError(f'[!] Отсутствует шаблоны. {set(template).difference(all_files)}')


def statement():
    """
    Генерируем документ (ведомость) с данными.
    Считываем шаблон и вводим данные.
    name_item_1 - наименование предмета.
    unit_measurement_1 - единица измерения.
    recipient's name - имя получателя.
    quantity - количество.
    total - итого.
    """
    # Ввод.
    doc = DocxTemplate('template/statement_tpl.docx')

    # Логика.
    contex = {f'name_item_{i}': '' for i in range(11)}

    # Вывод.
    doc.render(contex)
    doc.save('statement_tpl_final.docx')


def set_report():
    """Генерирует отчет по шаблону report_tpl.docx"""

    doc = DocxTemplate('template/report_tpl.docx')
    contex = {
        'post': 'Начальник',
        'title': 'полковник',
        'full_name': 'С.Д. Кольтяпин',
        'position_person': 'Начальник ОМТиХО',
        'title_person': 'майор полиции',
        'full_name_person': 'Е.Ю. Еламков',
        'evnt_date': '05.12.2022'
    }
    doc.render(contex)
    doc.save('report_tpl_final.docx')


def main():
    checking_templates()
    # set_report()
    # statement()


if __name__ == '__main__':
    main()
