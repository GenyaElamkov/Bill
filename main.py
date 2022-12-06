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
from random import randint

def checking_templates() -> None:
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
    В РАЗРАБОТКЕ
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


def set_report(data: dict) -> None:
    """Генерирует отчет по шаблону report_tpl.docx"""

    doc = DocxTemplate('template/report_tpl.docx')

    name_item = []

    # Сюда данные приходят извне.
    contex = {
        'post': data['post'],
        'title': data['title'],
        'full_name': data['full_name'],
        'position_person': data['position_person'],
        'title_person': data['title_person'],
        'full_name_person': data['full_name_person'],
        'evnt_date': '05.12.2022',
        'name_item': name_item
    }
    doc.render(contex)
    doc.save('report_tpl_final.docx')


def set_data() -> dict:
    """Ввод тестовых данных"""
    post = ['Начальник', 'Заместитель начальника', 'Врио начальника']
    title = ['полковник полиции', 'подполковник полиции']
    full_name = ['И.И. Иванов', 'П.П. Петров']
    position_person = ['Начальник ОМТиХО']
    title_person = ['майор полиции', 'капитан полиции']
    full_name_person = ['С.С. Сидоров', 'А.А. Александров']
    dict_data = {
        'post': post,
        'title': title,
        'full_name': full_name,
        'position_person': position_person,
        'title_person': title_person,
        'full_name_person': full_name_person,
    }
    return dict_data


def main():
    # print(set_data())
    checking_templates()
    data = set_data()
    set_report(data)
    # statement()


if __name__ == '__main__':
    main()
