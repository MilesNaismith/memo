from docx import Document
from docx.shared import Inches, Cm
import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH
from shutil import copyfile

def get_note(company_name, budget_item, pay_reason, pay_check, pay_sum):
    tomorrow = datetime.date.today() + datetime.timedelta(days=1)
    document = Document()
    ###Шапка###
    p = document.add_paragraph('СОГЛАСОВАНО \nГенеральному директору \nАО «Газпромбанк Лизинг» \nМ.А. Агаджанову') 
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p = document.add_paragraph('')
    p.add_run('Служебная записка на согласование \nадминистративно-хозяйственных расходов от {}'.format(datetime.date.today().strftime('%d.%m.%Y'))).italic = True
    document.add_paragraph('')
    p = document.add_paragraph('Уважаемый Максим Анатольевич!')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_paragraph('Прошу Вас согласовать расход на административно-хозяйственные нужды:')

    ###Тело###

    ### Название компании
    company = document.add_paragraph(
        'Получатель/контрагент ', style='List Bullet'
    )
    company.add_run(company_name).bold = True

    ### Сумма оплаты
    price = document.add_paragraph(
        'Сумма расхода', style='List Bullet'
    )
    price.add_run('{} RUB'.format(pay_sum)).bold = True

    ### Дата оплаты
    pay_date = document.add_paragraph(
        'Ожидаемый срок расхода ', style='List Bullet'
    )
    pay_date.add_run(tomorrow.strftime('%d.%m.%Y')).bold = True

    ### Содержимое счета
    pay_explanation = document.add_paragraph(
        'Обоснование расхода', style='List Bullet'
    )
    pay_explanation.add_run(pay_reason).bold = True

    ### Номер счета
    pay_doc = document.add_paragraph(
        'Документ, на основании которого осуществляется расход (приложить счет/договор, если применимо)', style='List Bullet'
    )
    pay_doc.add_run('Счет № {}'.format(pay_check)).bold = True

    ### Статья бюджета
    budget = document.add_paragraph(
        'Статья бюджета* ', style='List Bullet'
    )
    budget.add_run('ИТ расходы/{}'.format(budget_item)).bold = True

    ###Подвал###
    document.add_paragraph('')
    document.add_paragraph('Инициатор расхода                                                                                                                Трусков А.А.')
    document.add_paragraph('Инициатор расхода                                                                                                              Карякин М.Б.')   
    document.add_paragraph('Инициатор расхода                                                                                                             Богданов Д.Т.')
    document.add_paragraph('')
    document.add_paragraph('СОГЛАСОВАНО:')
    document.add_paragraph('Инициатор расхода                                                                                                            Клочкова Е.В.')
    document.add_paragraph('Инициатор расхода                                                                                                          Ястребова Н.Н.')
    table = document.add_table(rows=1, cols=2, style='Table Grid')
    table.height = Cm(0.8)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Информация о бюджете'
    hdr_cells[0].width = Cm(5.0)
    hdr_cells[1].width = Cm(12.0)
    document.save('note.docx')
    copyfile('note.docx', 'note/media/note.docx')
    
    
if __name__ == "__main__":   
    get_note()    