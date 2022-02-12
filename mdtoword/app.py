import sys
import argparse
import marko

from h2d import HtmlToDocx

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor, Cm


def html_to_docx(html) -> Document:
    """
    Создаем Document с html строки
    Используем для этого скрипт с библиотеки html2docx
    С некоторыми доработками
    """

    html = html.replace('<pre>', '<p><pre>')
    html = html.replace('</pre>', '</p></pre>')

    document = Document()
    new_parser = HtmlToDocx()
    new_parser.add_html_to_document(html, document)
    return document


def docx_set_margins(docx):
    """
    Выставляем отступы страниц для документа:
    Верхний - 2см
    Нижний - 2см
    Левый - 2.5см
    Правый - 1.5см
    """
    sections = docx.sections

    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.5)


def docx_styling(docx):
    """
    Выбираем все стили документа, передаем их в base_style_settings
    Для заголовка первого уровня:
    Выравнивание по центру
    Все буквы большие
    Остальное, как для стандартных стилей
    """
    for style in docx.styles:
        if style.type == WD_STYLE_TYPE.PARAGRAPH:
            base_style_settings(style)

    h1 = docx.styles['Heading 1']
    h1.font.all_caps = True
    h1.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


def base_style_settings(style):
    """
    Выставляем базовые стили для абзацев документа
    Шрифт 'Times New Roman', 14пт, черный
    Параграф без лишних отступов, только красная строка с 1см отступа
    Междустрочный интервал 1.5 строки.
    Выравнивание текста - всегда по ширине
    """
    font = style.font
    font.size = Pt(14)
    font.name = 'Times New Roman'
    font.color.rgb = RGBColor(0, 0, 0)
    font.bold = False
    font.italic = False
    font.cs_bold = False
    font.cs_italic = False

    paragraph_format = style.paragraph_format
    paragraph_format.first_line_indent = Cm(1)
    paragraph_format.left_indent = Cm(0)
    paragraph_format.space_before = Cm(0)
    paragraph_format.widow_control = True
    paragraph_format.line_spacing = 1.5
    paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY


def _parse_cmd_args() -> argparse.Namespace:
    """
    Используем встроенную библиотеку argparse для того, чтобы
    проверить наличие аргументов командной строки необходимых для запуска
    и собрать их в один объект
    Затем проверяем пути к файлам на правильность/доступность
    """
    parser = argparse.ArgumentParser()
    parser.add_argument('md_path', help='Path to md file for compile')
    parser.add_argument('-o', '--output', help='Path to compiled docx file')
    args = parser.parse_args()
    errors = list()

    try:
        open(args.md_path, 'r')
    except IOError as err:
        errors.append(str(err))

    if args.output:
        try:
            open(args.output, 'w')
        except IOError as err:
            errors.append(str(err))

    if errors:
        for err in errors:
            print(err)
            exit(1)

    return args


def md_file_to_html(md_text: str) -> str:
    """
    Конвертируем md разметку в виде строки в html
    Через библиотеку marko
    """
    html = marko.convert(md_text)
    return html


def _main():
    """
    Собираем аргументы,
    Читаем md файл, создаем из него документ,
    Результат сохраняем по указанному пути
    """
    args = _parse_cmd_args()

    with open(args.md_path) as file:
        html = md_file_to_html(file.read())

    document = html_to_docx(html)

    docx_set_margins(document)
    docx_styling(document)

    if args.output:
        document.save(args.output)
    else:
        document.save(sys.stdout.buffer)


if __name__ == '__main__':
    _main()
