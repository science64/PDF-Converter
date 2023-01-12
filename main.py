__author__ = "Süleyman Bozkurt and ChatGPT"
__version__ = "v1.6"
__maintainer__ = "Süleyman Bozkurt"
__email__ = "sbozkurt.mbg@gmail.com"
__date__ = '28.12.2022'
__update__ = '12.01.2023'

from tkinter import *
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import os
import threading
import traceback
import tkinter.messagebox
import pytesseract
from docx import Document
from pdf2image import convert_from_path
from pdf2docx import Converter
from tempfile import TemporaryDirectory
from PIL import ImageTk, Image
import webbrowser
from tkinter.font import Font

class about(ttk.Frame):
    def __init__(self, parent):

        textAboutMe = """
        Hello! My name is Süleyman Bozkurt and I am a senior PhD student in the field of biochemistry. 
        My research focuses on mitochondrial protein import and proteomics, 
        and I am particularly interested in understanding how it's impact on cellular function.
        In addition to my research, I am also interested in building my own apps and sharing them with others. 
        I am a huge fan of Python and enjoy using R for visualizations. I believe that these tools 
        can help make complex scientific data more accessible and easier to understand.
        I am passionate about using my skills and knowledge to make a positive impact in the scientific community. 
        I am always looking for new ways to learn and grow, and I enjoy collaborating with others 
        to find innovative solutions to complex problems.
        I hope this gives you a sense of who I am and what I am interested in. 
        If you have any questions or would like to learn more about my work, please don't hesitate to contact me.
        """

        ttk.Frame.__init__(self, parent)
        self.font = Font(family="Times New Roman", size=20)
        label2 = Label(self, text="")
        label2.grid(padx=1000, pady=800)
        label2 = Label(self, text="ABOUT ME!",font=self.font)
        label2.place(x=420, y=35)

        label = tk.Label(self, font = Font(family="Times New Roman", size=14),
                         text = textAboutMe,
                         )

        label.place(x=40, y=75)

        # Open the image using PIL
        self.github_logo = Image.open("./files/github.png")
        self.linkedin_logo = Image.open("./files/linkedin.png")
        self.twitter_logo = Image.open("./files/twitter.png")
        self.ibc2_logo = Image.open("./files/ibc2logo.png")

        # Resize the image
        self.github_logo = self.github_logo.resize((100, 100), resample=Image.Resampling.LANCZOS)
        self.linkedin_logo = self.linkedin_logo.resize((100, 100), resample=Image.Resampling.LANCZOS)
        self.twitter_logo = self.twitter_logo.resize((100, 100), resample=Image.Resampling.LANCZOS)
        self.ibc2_logo = self.ibc2_logo.resize((100, 100), resample=Image.Resampling.LANCZOS)

        # Create a PhotoImage object from the resized image
        self.github_logo = ImageTk.PhotoImage(self.github_logo)
        self.linkedin_logo = ImageTk.PhotoImage(self.linkedin_logo)
        self.twitter_logo = ImageTk.PhotoImage(self.twitter_logo)
        self.ibc2_logo = ImageTk.PhotoImage(self.ibc2_logo)

        # Create clickable link to your website
        website_link = tk.Label(self, image=self.ibc2_logo, text="Website", fg="blue", cursor="hand2")
        website_link.place(x=220, y=370)
        website_link.bind("<Button-1>", self.open_website)

        # Create clickable link to your GitHub page with logo
        github_link = tk.Button(self, image=self.github_logo, cursor="hand2", bd=0, highlightthickness=0,
                                activebackground="white")
        github_link.place(x=370, y=370)
        github_link.bind("<Button-1>", self.open_github)

        # Create clickable link to your LinkedIn profile with logo
        linkedin_link = tk.Button(self, image=self.linkedin_logo, cursor="hand2", bd=0, highlightthickness=0,
                                  activebackground="white")
        linkedin_link.place(x=520, y=370)
        linkedin_link.bind("<Button-1>", self.open_linkedin)

        # Create clickable link to your Twitter profile with logo
        twitter_link = tk.Button(self, image=self.twitter_logo, cursor="hand2", bd=0, highlightthickness=0,
                                 activebackground="white")
        twitter_link.place(x=670, y=370)
        twitter_link.bind("<Button-1>", self.open_twitter)

        label2 = Label(self, text="You can reach me one of these profiles!",font=self.font)
        label2.place(x=270, y=500)

    def open_github(self, event):
        # Open your GitHub page in the default web browser
        import webbrowser
        webbrowser.open("https://github.com/science64")

    def open_website(self, event):
        # Open your website in the default web browser
        import webbrowser
        webbrowser.open("https://biochem2.com/people/bozkurt-sueleyman")

    def open_linkedin(self, event):
        # Open your LinkedIn profile in the default web browser
        import webbrowser
        webbrowser.open("https://www.linkedin.com/in/s%C3%BCleyman-bozkurt-51080b60/")

    def open_twitter(self, event):
        # Open your LinkedIn profile in the default web browser
        import webbrowser
        webbrowser.open("https://twitter.com/science_mbg")

    def callback(self, url):
        webbrowser.open_new(url)

class support(ttk.Frame):
    def __init__(self, parent):
        ttk.Frame.__init__(self, parent)
        self.font = Font(family="Times New Roman", size=16)

        img = Image.open("files/buymecofee.jpeg")
        img = img.resize((800, 400), Image.Resampling.LANCZOS)
        img = ImageTk.PhotoImage(img)
        panel = Label(self, image=img)
        panel.image = img
        panel.place(x=80, y=20, width=800, height=400)

        label2 = Label(self, text="You can support by buying a coffee ☕️ here ", font=self.font)
        label2.grid(padx=1000, pady=800)
        label2.place(x=300, y=450)

        self.clickme = Button(self, text='https://www.buymeacoffee.com/science64', fg='black', bg='#FFC19E',
                                font=Font(family="Times New Roman", size=18, weight='bold'), command=self.callback)
        self.clickme.place(x=270, y=500)

    def callback(self):
        webbrowser.open_new('https://www.buymeacoffee.com/science64')

class PDF2Doc(ttk.Frame):
    def __init__(self, parent):
        ttk.Frame.__init__(self, parent)

        self.pdf_label = tk.Label(self, text="Select a PDF file:", font=("Helvetica", 14), pady=10)
        self.pdf_label.pack()

        self.pdf_button = tk.Button(self, text="Browse", font=("Helvetica", 14), pady=10, command=self.select_pdf)
        self.pdf_button.pack()

        self.conversion_method_label = tk.Label(self, text="Select a conversion method:", font=("Helvetica", 14), pady=10)
        self.conversion_method_label.pack()

        self.conversion_method_var = tk.IntVar()

        # Set the default value of the IntVar to 1
        self.conversion_method_var.set(1)

        self.pdf_to_word_radio = tk.Radiobutton(self, text="PDF to Word Converter", value=1, font=("Helvetica", 14), pady=10,
                                                variable=self.conversion_method_var)
        self.pdf_to_word_radio.pack()

        self.ocr_radio = tk.Radiobutton(self, text="OCR (Optical Character Recognition)", value=2, font=("Helvetica", 14), pady=10,
                                        variable=self.conversion_method_var)
        self.ocr_radio.pack()

        self.language = StringVar()
        self.language.set("English")  # default value

        language_formats = ["English", "German", "Turkish",
                            "French", "Spanish", "Italian",
                            "Polish", "Russian", "Ukrainian", "Croatian",
                            "Arabic", "Chinese", "Hindi"] # 13 different language format so far supported!

        self.options = OptionMenu(self, self.language, *language_formats)
        self.options.pack()

        if self.language.get() == 'English':
            self.LangPref = 'eng'
        elif self.language.get() == 'French':
            self.LangPref = 'fra'
        elif self.language.get() == 'German':
            self.LangPref = 'deu'
        elif self.language.get() == 'Turkish':
            self.LangPref = 'tur'
        elif self.language.get() == 'Polish':
            self.LangPref = 'pol'
        elif self.language.get() == 'Spanish':
            self.LangPref = 'spa'
        elif self.language.get() == 'Italian':
            self.LangPref = 'ita'
        elif self.language.get() == 'Arabic':
            self.LangPref = 'ara'
        elif self.language.get() == 'Chinese':
            self.LangPref = 'chi_sim'
        elif self.language.get() == 'Hindi':
            self.LangPref = 'hin'
        elif self.language.get() == 'Croatian':
            self.LangPref = 'hrv'
        elif self.language.get() == 'Russian':
            self.LangPref = 'rus'
        elif self.language.get() == 'Ukrainian':
            self.LangPref = 'ukr'

        # self.output_text = tk.Text(root, height=40, width=30, font=("Helvetica", 14))
        # self.output_text.pack(padx=10, pady=10)
        self.output_frame = tk.Frame(self)
        self.output_frame.pack(padx=10, pady=10)

        # Create a Scrollbar widget and associate it with the Text widget
        self.scrollbar = tk.Scrollbar(self.output_frame, orient="vertical")
        self.scrollbar.pack(side="right", fill="y")

        # Create the Text widget and associate it with the Scrollbar widget
        self.output_text = tk.Text(self.output_frame, height=10, width=60, font=("Helvetica", 14),
                                   yscrollcommand=self.scrollbar.set)
        self.output_text.pack(side="left", fill="both", expand=True)

        # Set the command of the Scrollbar widget to the yview method of the Text widget
        self.scrollbar.config(command=self.output_text.yview)

        self.button_frame = tk.Frame(self)
        self.button_frame.pack()

        self.convert_button = tk.Button(self.button_frame, text="Convert", font=("Helvetica", 14), pady=10,
                                        command=self.convert,
                                        state=tk.DISABLED)
        self.convert_button.pack(side=tk.LEFT)

        self.open_button = tk.Button(self.button_frame, text="Open", font=("Helvetica", 14), pady=10,
                                        command=self.open_docx,)
        self.open_button.pack(side=tk.RIGHT)

    def select_pdf(self):
        try:
            self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if self.pdf_path:
                self.output_text.insert(tk.END, "- Selected PDF: " + self.pdf_path + "\n")
                self.convert_button.config(state=tk.NORMAL)
        except Exception:
            # Display an error message in a pop-up window
            tkinter.messagebox.showerror("Error",
                                         "An error occurred while selecting the PDF file.\n\n" + traceback.format_exc())

    def convert(self):
        # Create a new thread to perform the conversion
        conversion_thread = threading.Thread(target=self.conversion_worker)
        conversion_thread.start()

    def open_docx(self):
        try:
            # Open the docx file with the default program
            os.startfile(self.docx_path)
        except Exception:
            import tkinter.messagebox

            tkinter.messagebox.showerror("Error",
                                         "An error occurred while opening the docx file.\n\n" + traceback.format_exc())

    def conversion_worker(self):
        pdf_name, pdf_ext = os.path.splitext(os.path.basename(self.pdf_path))

        # Create the Word file path with the same file name and a .docx extension
        self.docx_path = os.path.join(os.path.dirname(self.pdf_path), pdf_name + ".docx")
        self.conversion_methodResult = self.conversion_method_var.get()

        try:
            if self.conversion_methodResult == 1:
                self.output_text.insert(tk.END, "- Convertion method: PDF to Word Converter!\n")
                self.output_text.insert(tk.END, "- Converting PDF...\n")
                # Use a PDF to Word converter to convert the PDF
                cv = Converter(self.pdf_path)
                cv.convert(self.docx_path)  # all pages by default
                cv.close()
            else:
                # Use OCR to convert the PDF
                self.output_text.insert(tk.END, "- Convertion method: OCR Converter!\n")
                self.output_text.insert(tk.END, "- Converting PDF...\n")

                with TemporaryDirectory() as temp_dir:
                    # Convert the PDF to a sequence of images

                    images = convert_from_path(self.pdf_path, poppler_path = "./files/poppler-22.04.0/Library/bin", dpi=300)

                    # Create a new Word document
                    document = Document()

                    pytesseract.pytesseract.tesseract_cmd = "./files/Tesseract-OCR/tesseract.exe"

                    # Extract the text from each image and add it to the Word document as a paragraph
                    for image in images:
                        text = pytesseract.image_to_string(image, lang=self.LangPref)
                        document.add_paragraph(text)

                core_properties = document.core_properties
                core_properties.author = 'PDF Converter v1.6 by Suleyman Bozkurt'
                core_properties.comments = 'PDF converted into DOCX with PDF Converter v1.6'
                core_properties.title = pdf_name

                # Save the Word document
                document.save(self.docx_path)
        except Exception as e:
            # Print an error message if the conversion fails
            print(f"An error occurred while converting the PDF to Word: {e}")
            tkinter.messagebox.showerror("Error",
                                         f"An error occurred, contact to the programmer. Error is {e}.\n\n" + traceback.format_exc())

        self.output_text.insert(tk.END, f"- Saved here: {self.docx_path}\n")
        self.output_text.insert(tk.END, "- Finished!\n")
        tkinter.messagebox.showinfo("Task complete", "The task is complete!")

class PDF2Image(ttk.Frame):
    def __init__(self, parent):
        ttk.Frame.__init__(self, parent)

        self.pdf_label = tk.Label(self, text="Select a PDF file:", font=("Helvetica", 14), pady=10)
        self.pdf_label.pack()

        self.pdf_button = tk.Button(self, text="Browse", font=("Helvetica", 14), pady=10, command=self.select_pdf)
        self.pdf_button.pack()

        self.conversion_method_label = tk.Label(self, text="Select a image format:", font=("Helvetica", 14), pady=10)
        self.conversion_method_label.pack()

        self.variable = StringVar()
        self.variable.set("tiff")  # default value
        image_formats = ["tiff", "jpg", "png", "eps", "bmp", "ico"] # 6 different file format so far

        self.options = OptionMenu(self, self.variable, *image_formats)
        self.options.pack()

        # self.output_text = tk.Text(root, height=40, width=30, font=("Helvetica", 14))
        # self.output_text.pack(padx=10, pady=10)
        self.output_frame = tk.Frame(self)
        self.output_frame.pack(padx=10, pady=10)

        # Create a Scrollbar widget and associate it with the Text widget
        self.scrollbar = tk.Scrollbar(self.output_frame, orient="vertical")
        self.scrollbar.pack(side="right", fill="y")

        # Create the Text widget and associate it with the Scrollbar widget
        self.output_text = tk.Text(self.output_frame, height=13, width=60, font=("Helvetica", 14),
                                   yscrollcommand=self.scrollbar.set)
        self.output_text.pack(side="left", fill="both", expand=True)

        # Set the command of the Scrollbar widget to the yview method of the Text widget
        self.scrollbar.config(command=self.output_text.yview)

        self.button_frame = tk.Frame(self)
        self.button_frame.pack()

        self.convert_button = tk.Button(self.button_frame, text="Convert", font=("Helvetica", 14), pady=10,
                                        command=self.convert,
                                        state=tk.DISABLED)
        self.convert_button.pack(side=tk.LEFT)

        self.open_button = tk.Button(self.button_frame, text="Open", font=("Helvetica", 14), pady=10,
                                        command=self.open_file,)
        self.open_button.pack(side=tk.RIGHT)

    def pdf_to_image(self, path, out_format):
        pdf_name, pdf_ext = os.path.splitext(os.path.basename(path))
        self.image_path = os.path.join(os.path.dirname(self.pdf_path), pdf_name)
        try:
            img = convert_from_path(path, poppler_path = "./files/poppler-22.04.0/Library/bin", dpi=300)
            num = len(img)

            for i, page in enumerate(img):
                if num == 1:
                    self.output_file = f'{self.image_path}.{out_format}'
                else:
                    self.output_file = f'{self.image_path}_page_{i+1}.{out_format}'

                page.save(self.output_file, out_format)
                self.output_text.insert(tk.END, f"- Saved here: {self.output_file}\n")
        except Exception as e:
            tkinter.messagebox.showerror("Error",
                                     "An error occurred while converting the PDF file.\n\n" + traceback.format_exc())

    def select_pdf(self):
        try:
            self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if self.pdf_path:
                self.output_text.insert(tk.END, "- Selected PDF: " + self.pdf_path + "\n")
                self.convert_button.config(state=tk.NORMAL)
        except Exception:
            # Display an error message in a pop-up window
            tkinter.messagebox.showerror("Error",
                                         "An error occurred while selecting the PDF file.\n\n" + traceback.format_exc())

    def convert(self):
        # Create a new thread to perform the conversion
        conversion_thread = threading.Thread(target=self.conversion_worker)
        conversion_thread.start()

    def open_file(self):
        try:
            # Open the docx file with the default program
            os.startfile(self.output_file)
        except Exception:
            import tkinter.messagebox

            tkinter.messagebox.showerror("Error",
                                         "An error occurred while opening the docx file.\n\n" + traceback.format_exc())

    def conversion_worker(self):

        self.options_method = self.variable.get()
        self.output_text.insert(tk.END, f"- Convertion format: PDF to {self.options_method}!\n")

        try:
            self.output_text.insert(tk.END, "- Converting PDF...\n")
            if self.options_method == 'tiff':
                self.pdf_to_image(self.pdf_path, 'tiff')  # Convert PDF to TIFF
            elif self.options_method == 'jpg':
                self.pdf_to_image(self.pdf_path, 'JPEG') # Convert PDF to JPEG
            elif self.options_method == 'png':
                self.pdf_to_image(self.pdf_path, 'png')  # Convert PDF to PNG
            elif self.options_method == 'ico':
                self.pdf_to_image(self.pdf_path, 'ICO')  # Convert PDF to ICO
            elif self.options_method == 'bmp':
                self.pdf_to_image(self.pdf_path, 'BMP')  # Convert PDF to SVG
            elif self.options_method == 'eps':
                self.pdf_to_image(self.pdf_path, 'eps') # Convert PDF to EPS

        except Exception as e:
            # Print an error message if the conversion fails
            print(f"An error occurred while converting the PDF to Word: {e}")
            tkinter.messagebox.showerror("Error",
                                         f"An error occurred, contact to the programmer. Error is {e}.\n\n" + traceback.format_exc())

        self.output_text.insert(tk.END, "- Finished!\n")
        tkinter.messagebox.showinfo("Task complete", "The task is complete!")

class MyWindow():

    def __init__(self, root):
        self.root = root
        self.notebook = ttk.Notebook(self.root)

        self.notebook.pack(expand=1, fill="both")
        self.PDFConverterFrame = PDF2Doc(self.notebook)
        self.PDFConverterFrame.bind("<<NotebookTabChanged>>", self.on_tab_selected)
        self.notebook.add(self.PDFConverterFrame, text="PDF to Docx")

        self.PDFimageFrame = PDF2Image(self.notebook)
        self.notebook.add(self.PDFimageFrame, text="PDF to Image")

        self.aboutFrame = about(self.notebook)
        self.notebook.add(self.aboutFrame, text="About")

        self.supportFrame = support(self.notebook)
        self.notebook.add(self.supportFrame, text="Support")

        #self.notebook.grid()
    def on_tab_selected(self):
        selected_tab = self.widget.select()
        tab_text = self.widget.tab(selected_tab, "text")

if __name__ == '__main__':
    root = Tk()
    root.title("PDF Converter v1.6 @2023", )
    root.geometry("960x600+480+250")
    root.resizable(0, 0)
    root.wm_iconbitmap('./files/icon.ico')
    MyWindow(root)
    root.mainloop()