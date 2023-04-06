import sys
import docx
import asyncio
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QLineEdit, QVBoxLayout, QHBoxLayout, QFileDialog


async def format_docx(input_path, output_path):
    # Открываем файл docx
    doc = docx.Document(input_path)

    # Проходим по всем параграфам в документе
    for para in doc.paragraphs:
        # Убираем лишние отступы
        para.paragraph_format.left_indent = 0
        para.paragraph_format.right_indent = 0

        # Форматируем абзацы
        para_format = para.paragraph_format
        para_format.space_before = docx.shared.Pt(12)
        para_format.space_after = docx.shared.Pt(12)
        para_format.line_spacing = docx.shared.Pt(14)

    # Экспортируем отформатированный текст обратно в .docx
    doc.save(output_path)


class App(QWidget):
    def __init__(self):
        super().__init__()

        # Создаем элементы интерфейса
        self.input_label = QLabel('Input file:')
        self.input_field = QLineEdit()
        self.input_button = QPushButton('Browse')
        self.output_label = QLabel('Output file:')
        self.output_field = QLineEdit()
        self.output_button = QPushButton('Browse')
        self.format_button = QPushButton('Format')

        # Организуем элементы интерфейса в компоновщики
        input_layout = QHBoxLayout()
        input_layout.addWidget(self.input_label)
        input_layout.addWidget(self.input_field)
        input_layout.addWidget(self.input_button)
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_field)
        output_layout.addWidget(self.output_button)
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.format_button)
        main_layout = QVBoxLayout()
        main_layout.addLayout(input_layout)
        main_layout.addLayout(output_layout)
        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

        # Подключаем обработчики событий
        self.input_button.clicked.connect(self.browse_input)
        self.output_button.clicked.connect(self.browse_output)
        self.format_button.clicked.connect(self.format_doc)

    def browse_input(self):
        # Открываем диалог выбора файла для выбора входного файла
        file_path, _ = QFileDialog.getOpenFileName(self, 'Open Document', '', 'Word Document (*.docx)')
        if file_path:
            self.input_field.setText(file_path)

    def browse_output(self):
        # Открываем диалог выбора файла для выбора выходного файла
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save Document', '', 'Word Document (*.docx)')
        if file_path:
            self.output_field.setText(file_path)

    def format_doc(self):
        # Получаем пути к входному и выходному файлам
        input_path = self.input_field.text()
        output_path = self.output_field.text()

        # Запускаем функцию format_docx в асинхронном режиме
        loop = asyncio.get_event_loop()
        loop.run_until_complete(format_docx(input_path, output_path))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
