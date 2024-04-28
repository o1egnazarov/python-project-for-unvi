import pandas as pd
from pptx import Presentation

excel_file_path = 'tasks.xlsx'

num_questions = int(input('Enter number of questions: '))


df = pd.read_excel(excel_file_path)
prs = Presentation()

# хуярим по строкам df для создания слайдов
for index, row in df.head(num_questions).iterrows():
    game_title = row['Название игры']
    question = row['Вопрос']
    answer = row['Ответ']
    game_class = row['Класс']

    # cоздаем слайд с вопросом
    question_slide = prs.slides.add_slide(prs.slide_layouts[1])
    question_title = question_slide.shapes.title
    question_content = question_slide.placeholders[1]
    question_title.text = f"{game_title}"
    question_content.text = f"Вопрос: {question}\nКласс: {game_class}"

    # cоздаем слайд с ответом
    answer_slide = prs.slides.add_slide(prs.slide_layouts[1])
    answer_title = answer_slide.shapes.title
    answer_content = answer_slide.placeholders[1]
    answer_content.text = f"Ответ: {answer}"

prs.save('output_presentation.pptx')

