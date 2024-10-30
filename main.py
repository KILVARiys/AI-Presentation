import os
from transformers import pipeline
from pptx import Presentation

# Указываем модель и ревизию
model_name = "facebook/bart-large-cnn"
revision = "main"


# Функция для генерации сводки
def generate_summary(input_text):
    summarizer = pipeline("summarization", model=model_name, revision=revision)
    summary = summarizer(input_text, max_length=80, min_length=20, do_sample=False)
    return summary[0]['summary_text']


# Функция для создания презентации
def create_presentation(summary_text):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # Заголовок и контент
    title = slide.shapes.title
    content = slide.placeholders[1]

    title.text = "Заголовок"
    content.text = summary_text

    # Путь к папке Загрузки
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "presentation.pptx")
    prs.save(downloads_path)

    return downloads_path  # Возвращаем путь к сохраненному файлу


# Основной код
if __name__ == "__main__":
    # Входной текст
    input_text = "Basilio the cat checks the presentation for functionality"

    # Генерация сводки
    summary = generate_summary(input_text)
    print("Сводка:", summary)

    # Создание презентации и получение пути к файлу
    saved_path = create_presentation(summary)
    print(f"Презентация сохранена в: {saved_path}")