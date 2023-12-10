import tkinter as tk
from tkinter import ttk, filedialog
from pdf2docx import parse
import pathlib
import fitz


class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()

    def init_main(self):
        toolbar = tk.Frame(bg='white', bd=2)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        self.img_add = tk.PhotoImage(file="Spell Check32.png")
        btn_add_new = ttk.Button(toolbar, text="Распознать", command=self.open_dialog,
                                 compound=tk.TOP, image=self.img_add)
        btn_add_new.pack(side=tk.LEFT)

        self.img_upd = tk.PhotoImage(file="New Edit32.png")
        btn_edit_dialog = ttk.Button(toolbar, text="Добавить текст", command=self.open_update_dialog,
                                     compound=tk.TOP, image=self.img_upd)
        btn_edit_dialog.pack(side=tk.LEFT)
    def open_dialog(self):
        Child()

    def open_update_dialog(self):
        Update()

class Child(tk.Toplevel):
    def __init__(self):
        super().__init__(root)
        self.init_child()
        self.view = app

    def init_child(self):
        icon = tk.PhotoImage(file="Spell Check32.png")
        self.iconphoto(False, icon)
        self.title('Конвертер PDF в Word')
        self.geometry('400x300+300+300')
        self.resizable(False, False)

        def callback():
            name = filedialog.askopenfilename()
            self.ePath.config(state='normal')
            self.ePath.delete('1', -1)
            self.ePath.insert('1', name)
            self.ePath.config(state='readonly')

        def convert():
            pdf_file = self.ePath.get()
            word_file = pathlib.Path(pdf_file)
            word_file = word_file.stem + '.docx'
            parse(pdf_file, word_file)
            self.message_done = ttk.Label(self, text='Конвертация завершена')
            self.message_done.pack(pady=10)

        self.entry_file_path = ttk.Button(self, text='Выбрать PDF файл', command=callback)
        self.entry_file_path.pack(pady=10)
        # self.entry_file_path.place(x=200, y=230)

        self.lbPath = ttk.Label(self, text='Путь к файлу:')
        self.lbPath.pack()

        self.ePath = ttk.Entry(self, width=50, state='readonly')
        self.ePath.pack(pady=10)

        self.btnConvert = ttk.Button(self, text='Конвертировать', command=convert)
        self.btnConvert.pack(pady=10)

        btn_cancel = ttk.Button(self, text='Закрыть', command=self.destroy)
        btn_cancel.place(x=300, y=270)

        self.grab_set()
        self.focus_set()

class Update(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.init_edit()
        self.view = app

    def init_edit(self):
        icon = tk.PhotoImage(file="New Edit32.png")
        self.iconphoto(False, icon)
        self.title('Добавление текста в PDF')
        self.geometry('400x300+300+300')
        self.resizable(False, False)

        def choose_pdf():
            name = filedialog.askopenfilename()
            self.pdfePath.config(state='normal')
            self.pdfePath.delete('1', -1)
            self.pdfePath.insert('1', name)
            self.pdfePath.config(state='readonly')

        def choose_txt():
            name = filedialog.askopenfilename()
            self.txtePath.config(state='normal')
            self.txtePath.delete('1', -1)
            self.txtePath.insert('1', name)
            self.txtePath.config(state='readonly')

            input_txt_path = self.txtePath.get()

            with open(input_txt_path) as text_source:
                text = text_source.readlines()
                self.text_to_add = [line.rstrip() for line in text]
                # print(text)

        def add_text_to_pdf():
            # Open the existing PDF file
            input_pdf_path = self.pdfePath.get()

            pdf_document = fitz.open(input_pdf_path)
            i = 0
            # Iterate through all pages of the PDF
            for page_num in range(pdf_document.page_count):
                # Get the page from the PDF document
                page = pdf_document[page_num]

                # Create a new text object to add to the page
                text_rectangle = fitz.Rect(100, 100, 500, 500)  # Adjust the rectangle coordinates as needed
                text_point = fitz.Point(50, 50)  # Set the starting point of the text
                text_to_page = self.text_to_add[0 + i]
                page.insert_text(text_point, text_to_page, fontsize=15, fontname="Helvetica")
                i += 1
            # Save the modified PDF to a new file
            pdf_document.save('pdf_added_text.pdf')
            pdf_document.close()
            self.message_done = ttk.Label(self, text='Текст добавлен')
            self.message_done.pack(pady=10)

        self.entry_pdf_path = ttk.Button(self, text='Выбрать PDF файл', command=choose_pdf)
        self.entry_pdf_path.pack(pady=10)
        # self.entry_file_path.place(x=200, y=230)

        self.pdfPath = ttk.Label(self, text='Путь к файлу:')
        self.pdfPath.pack()

        self.pdfePath = ttk.Entry(self, width=50, state='readonly')
        self.pdfePath.pack(pady=10)

        self.entry_txt_path = ttk.Button(self, text='Выбрать txt файл', command=choose_txt)
        self.entry_txt_path.pack(pady=10)
        # self.entry_file_path.place(x=200, y=230)

        self.txtPath = ttk.Label(self, text='Путь к файлу:')
        self.txtPath.pack()

        self.txtePath = ttk.Entry(self, width=50, state='readonly')
        self.txtePath.pack(pady=10)

        self.btnConvert = ttk.Button(self, text='Добавить текс', command=add_text_to_pdf)
        self.btnConvert.pack(pady=10)

        btn_cancel = ttk.Button(self, text='Закрыть', command=self.destroy)
        btn_cancel.place(x=300, y=270)

        self.grab_set()
        self.focus_set()
















if __name__ == '__main__':
    root = tk.Tk()
    app = Main(root)
    app.pack(fill="both", expand=1, pady=10, padx=10, side=tk.LEFT)
    # Название
    root.title("Распознавание и добавление записей ")
    # Иконка
    icon = tk.PhotoImage(file="pdf.png")
    root.iconphoto(False, icon)
    # Размеры окна
    root.geometry("400x65+300+300")
    root.resizable(True, True)
    # Запуск
    root.mainloop()
    root.mainloop()

