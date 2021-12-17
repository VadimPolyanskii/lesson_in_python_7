import datetime

from docxtpl import DocxTemplate
from docxtpl import InlineImage
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage


def get_context(brand, model, fuel_consumption, price):
    return {
        'brand': brand,
        'model': model,
        'fuel_consumption': fuel_consumption,
        'price': price
    }


def from_template(brand, model, fuel_consumption, price, template, signature):
    template = DocxTemplate(template)
    context = get_context(brand, model, fuel_consumption, price)

    # Задаём параметры картинки
    img_size = Cm(15)
    acc = InlineImage(template, signature, img_size)

    # Насыщаем шаблон передаваемой информацией
    context['acc'] = acc
    template.render(context)

    # Сохраняем получившийся файл с информацией
    template.save(brand + '_' + str(datetime.datetime.now().date()) + '_data.docx')


def generate_report(brand, model, fuel_consumption, price):
    template = 'report.docx'
    signature = 'skoda.jpeg'
    document = from_template(brand, model, fuel_consumption, price, template, signature)


def to_fixed(num_obj, digits=0):
    return f"{num_obj:.{digits}f}"


generate_report('Skoda', 'Octavia', '9 l/100 km', '1 500 000 RUB')








