import subprocess
import sys
import webbrowser

import pandas as pd
import speech_recognition as sr
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit


class TreeRecognitionApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Распознавание текста")
        self.setGeometry(100, 100, 1050, 600)

        self.text_edit = QTextEdit(self)
        self.text_edit.setGeometry(20, 20, 760, 300)

        self.run_button = QPushButton("выбрать .txt файл", self)
        self.run_button.setGeometry(50, 350, 200, 75)
        self.run_button.clicked.connect(self.run_txt_recognition)

        self.run_button = QPushButton("выбрать аудио файл", self)
        self.run_button.setGeometry(50, 450, 200, 75)
        self.run_button.clicked.connect(self.run_audio_recognition)

        self.open_website_button = QPushButton("Открыть сайт", self)
        self.open_website_button.setGeometry(550, 350, 200, 50)
        self.open_website_button.clicked.connect(self.open_website)

        self.open_website_button = QPushButton("Открыть таблицу Excel", self)
        self.open_website_button.setGeometry(800, 350, 200, 50)
        self.open_website_button.clicked.connect(self.open_excel)

        self.edited_text = ""  # Переменная для хранения отредактированного текста

    def open_excel(self):
        try:
            excel_file = "output.xlsx"  # Укажите путь к файлу "output.xlsx" здесь
            subprocess.Popen(['start', 'excel', excel_file], shell=True)
        except Exception as e:
            print(f"Ошибка при открытии Excel-файла: {e}")

    def open_website(self):
        webbrowser.open('result.html', new=2)  # Открываем HTML-файл в браузере

    def run_audio_recognition(self):
        audio_file, _ = QFileDialog.getOpenFileName(self, "Open Audio File", "", "Audio Files (*.mp3 *.wav)")

        recognizer = sr.Recognizer()

        with sr.AudioFile(audio_file) as source:
            audio_data = recognizer.record(source)
            try:
                text = recognizer.recognize_google(audio_data, language="ru-RU")
            except sr.UnknownValueError:
                print("Речь не распознана")
            except sr.RequestError as e:
                print("Ошибка при выполнении запроса к сервису Google: {0}".format(e))

        result_text = "Распознанный текст: " + text + '\n' + recognize(text)
        self.text_edit.clear()
        self.text_edit.insertPlainText(result_text)

        # Сохраните отредактированный текст в переменную self.edited_text
        self.edited_text = self.text_edit.toPlainText()

    def run_txt_recognition(self):
        text_file, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Files (*.txt)")

        text = ""
        with open(text_file, "r", encoding="ANSI") as file:
            text = file.read()

        result_text = "Распознанный текст: " + text + '\n' + recognize(text)

        self.text_edit.clear()
        self.text_edit.insertPlainText(result_text)


def recognize(text):
    types = 'осина,береза,елка,сосна'

    text3 = text
    text = text.lower() + ' конец текста точно точно точно'
    text = text.replace(',', "")
    text = text.replace('.', "")

    data_list = []
    words = text.split()

    while len(words) >= 4:
        tree_type = ""
        word = words.pop(0)
        print(text)

        for tree in types.split(','):
            if word == tree and words[0].isnumeric():
                if words[1] == "дрова":
                    tree_type = word + ' дрова'
                    height = words[0]
                    text = ' '.join(words[3:])
                else:
                    tree_type = word
                    height = words[0]
                    text = ' '.join(words[2:])
                data_list.append({'Tree Type': tree_type, 'Height': height})

    df = pd.DataFrame(data_list)
    df['Count'] = df.groupby(['Tree Type', 'Height'])['Tree Type'].transform('count')
    df = df.sort_values(by='Tree Type')
    df = df.drop_duplicates(['Tree Type', 'Height'])

    excel_file = 'output.xlsx'
    df.to_excel(excel_file, index=False)

    second_digit = -1
    prev_tree_type = None

    def generate_code(row):
        nonlocal second_digit, prev_tree_type

        if 'дрова' in row['Tree Type']:
            digit_1 = '3'
        else:
            digit_1 = '1'

        if 'дрова' in row['Tree Type']:
            current_tree_type = row['Tree Type'].replace(' дрова', '')
        else:
            current_tree_type = row['Tree Type']

        if current_tree_type != prev_tree_type:
            prev_tree_type = current_tree_type
            second_digit += 1

        digit_2 = str(second_digit)
        code = f"v{digit_1}{digit_2}_{row['Height']}"

        return code

    df['Code'] = df.apply(generate_code, axis=1)

    codes = df['Code'].tolist()
    codes_str = ', '.join([f'"{code}":{row.Count}' for code, row in zip(codes, df.itertuples())])
    codes_str = '{' + codes_str + ','

    tree_counts = df.groupby(['Tree Type', 'Height']).agg({'Count': 'sum'}).reset_index()

    with open("9.html", "r") as file:
        data = file.read()

    data = data[1:]

    with open("result.html", "w") as file:
        file.write(codes_str + data)

    return codes_str


def main():
    app = QApplication(sys.argv)
    window = TreeRecognitionApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
