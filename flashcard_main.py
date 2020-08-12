from tkinter import *
from tkinter import filedialog,scrolledtext
from PIL import ImageTk, Image
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import xlsxwriter

root = Tk()
root.title("Flashcards")
root.iconbitmap('C:/Users/zacha/PycharmProjects/Flashcards/imgs/pacman_ghost.ico')
root.minsize(1200,800)

file_indexer = 0
end_of_file = False
fonttype = "Franklin Gothic Book"

def importfile(value):
   global file_indexer
   global end_of_file
   global df
   global filename
   global question
   file_indexer=1
   end_of_file=False
   filename = filedialog.askopenfilename(title="Select file",filetypes =(("XLSX","*.xlsx"),("ALL","*.*")))
   if filename:
      df = pd.read_excel(filename, index_col="id")
      if value == 2:
         df = df.sample(frac=1).reset_index(drop=True)
      question_txt = scrolledtext.ScrolledText(question_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
      question_txt.grid(row=0, column=0, sticky=E + W + N + S)
      question_txt.insert(INSERT, df.loc[file_indexer, "question"])
      question = question_txt.get('1.0', 'end-1c')
   else:
      file_indexer == 0
      return
   return

def show(events=None):
   if file_indexer > 0 and end_of_file!=True:
      answer_txt = scrolledtext.ScrolledText(answer_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
      answer_txt.grid(row=0, column=0, sticky=E + W + N + S)
      answer_txt.insert(INSERT, df.loc[file_indexer, "answer"])
   else:
      return

def next(event=None):
   global file_indexer
   global df
   global end_of_file
   global question
   if file_indexer !=0 and end_of_file !=True:
      try:
         file_indexer +=1
         question_txt = scrolledtext.ScrolledText(question_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
         question_txt.grid(row=0, column=0, sticky=E + W + N + S)
         question_txt.insert(INSERT, df.loc[file_indexer, "question"])
         answer_txt = scrolledtext.ScrolledText(answer_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
         answer_txt.grid(row=0, column=0, sticky=E + W + N + S)
         answer_txt.insert(INSERT, "")
      except KeyError:
         question_txt = scrolledtext.ScrolledText(question_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
         question_txt.grid(row=0, column=0, sticky=E + W + N + S)
         question_txt.insert(INSERT, "End of file reached \nPlease import a new file.")
         end_of_file = True
         return
      except NameError:
         question_txt = scrolledtext.ScrolledText(question_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
         question_txt.grid(row=0, column=0, sticky=E + W + N + S)
         question_txt.insert(INSERT, "No file loaded.  Please load file.")
   try:
      question = question_txt.get('1.0', 'end-1c')
   except NameError:
      return
   return

def previous(event=None):
   global file_indexer
   global df
   global end_of_file
   global question
   end_of_file = False
   if file_indexer <= 1:
      return
   try:
      file_indexer -=1
      question_txt = scrolledtext.ScrolledText(question_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
      question_txt.grid(row=0, column=0, sticky=E + W + N + S)
      question_txt.insert(INSERT, df.loc[file_indexer, "question"])
      answer_txt = scrolledtext.ScrolledText(answer_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
      answer_txt.grid(row=0, column=0, sticky=E + W + N + S)
      answer_txt.insert(INSERT, "")
   except NameError:
      question_txt = scrolledtext.ScrolledText(question_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
      question_txt.grid(row=0, column=0, sticky=E + W + N + S)
      question_txt.insert(INSERT, "No file loaded.  Please load file.")
   try:
      question = question_txt.get('1.0', 'end-1c')
   except NameError:
      return
   return

def notes():
   #Initialize the notes pop-up window
   try:
      global notes_txt
      global notesWindow
      notes = pd.read_excel(filename, index_col="id")
      note_index = notes[notes['question']==question].index.values.astype(int)[0]
      notes = notes.loc[note_index].at['notes']
      notesWindow = Toplevel(root)
      notesWindow.title("Add notes to flashcard")
      notesWindow.geometry("800x483")
      notesWindow.iconbitmap('C:/Users/zacha/PycharmProjects/Flashcards/imgs/pacman_ghost.ico')
      notes_frame = LabelFrame(notesWindow, text="Type your note", padx=5, pady=5)
      notes_frame.grid(row=0, column=0, columnspan=3, padx=10, pady=0, sticky=E+W+N+S)
      notes_txt = scrolledtext.ScrolledText(notes_frame, pady=5, width=90)
      notes_txt.grid(row=0, column=0, sticky=E + W + N + S, padx=10, pady=10)
      if pd.isnull(notes):
         notes_txt.insert(INSERT, "")
      else:
         notes_txt.insert(INSERT, notes)
      save_button = Button(notesWindow, text="Save Note", command=lambda: note_save(notes_txt), padx=360)
      save_button.grid(row=1, column=0, sticky=W+N, padx=10, pady=5)
   except NameError:
      return

def get_notes():
   #Used to get the notes from window for the save_notes function
   global notes_received
   notes_received = notes_txt.get('1.0', 'end-1c')
   return notes_received

def note_save(input):
   #Save the notes to original filename
   note_df = pd.read_excel(filename, index_col="id")
   note_indexer = note_df[note_df['question']==question].index.values.astype(int)[0]
   note_df.loc[[note_indexer], ["notes"]] = get_notes()
   writer = pd.ExcelWriter(filename, engine="xlsxwriter")
   note_df.to_excel(writer, sheet_name="Sheet1")
   workbook = writer.book
   worksheet = writer.sheets['Sheet1']
   format1= workbook.add_format()
   format2= workbook.add_format()
   format3= workbook.add_format()
   worksheet.set_column('B:B', 100, format1)
   worksheet.set_column('C:C', 100, format2)
   worksheet.set_column('D:D', 100, format3)
   writer.save()
   return notesWindow.destroy()

def CommandReference():
   CommandReferenceWindow = Toplevel(root)
   CommandReferenceWindow.title("Command Reference")
   CommandReferenceWindow.geometry("300x100")
   CommandReferenceWindow.iconbitmap('C:/Users/zacha/PycharmProjects/Flashcards/imgs/pacman_ghost.ico')
   leftTextLabel = Label(CommandReferenceWindow, text="[Left Arrow]")
   leftTextLabel.grid(row=0, column=0, sticky=W)
   rightTextLabel = Label(CommandReferenceWindow, text="[Right Arrow]")
   rightTextLabel.grid(row=1, column=0, sticky=W)
   downTextLabel = Label(CommandReferenceWindow, text="[Down Arrow]")
   downTextLabel.grid(row=2, column=0, sticky=W)
   spaceTextLabel = Label(CommandReferenceWindow, text="[Spacebar]")
   spaceTextLabel.grid(row=3, column=0, sticky=W)
   leftresultLabel = Label(CommandReferenceWindow, text="Previous Card")
   leftresultLabel.grid(row=0, column=1, sticky=W)
   rightresultLabel = Label(CommandReferenceWindow, text="Next Card")
   rightresultLabel.grid(row=1, column=1, sticky=W)
   downresultLabel = Label(CommandReferenceWindow, text="Show Answer")
   downresultLabel.grid(row=2, column=1, sticky=W)
   spaceresultLabel = Label(CommandReferenceWindow, text="Show Answer")
   spaceresultLabel.grid(row=3, column=1, sticky=W)
   return


def about():
   aboutWindow = Toplevel(root)
   aboutWindow.title("Command Reference")
   aboutWindow.geometry("400x400")
   aboutWindow.iconbitmap('C:/Users/zacha/PycharmProjects/Flashcards/imgs/pacman_ghost.ico')
   f= open("readme.txt", "r")
   if f.mode =="r":
      readme = f.read()
   f.close()
   aboutLabel = Label(aboutWindow, text=readme)
   aboutLabel.grid(row=0, column=0, sticky=W)

# --- Creates the menu bar
menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Import Ordered", command=lambda: importfile(1))
filemenu.add_command(label="Import Randomized", command=lambda: importfile(2))
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=filemenu)
helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Command Ref", command=CommandReference)
helpmenu.add_command(label="About", command=about)
menubar.add_cascade(label="Help", menu=helpmenu)
root.config(menu=menubar)
# --- End of menu bar

# Creation and placement of buttons ------------------------------
buttons_frame = Frame(root)
buttons_frame.grid(row=0, column=0, sticky=W+E)

btn_previous = Button(buttons_frame, text='Previous', command=previous)
btn_previous.pack(side="left", padx=20, pady=10)

btn_answer = Button(buttons_frame, text='Show Answer', command=show)
btn_answer.pack(side="left", padx=20, pady=10)

btn_next = Button(buttons_frame, text='Next', command=next)
btn_next.pack(side="left", padx=20, pady=10)

btn_notes = Button(buttons_frame, text='Notes', command=notes)
btn_notes.pack(side="right", padx=20, pady=10)

# Creation of shortcut keys ---------------------------------------
root.bind("<Right>", next)
root.bind("<Left>", previous)
root.bind("<Down>", show)
root.bind("<space>", show)

# Group1 Frame ----------------------------------------------------
question_frame = LabelFrame(root, text="Question", padx=5, pady=5)
question_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=0, sticky=E+W+N+S)
answer_frame = LabelFrame(root, text="Answer", padx=5, pady=5)
answer_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky=E+W+N+S)

# Frame Formatting -----------------------------------------------
root.columnconfigure(0, weight=1)
root.rowconfigure(1, weight=2)
root.rowconfigure(2, weight=2)
question_frame.rowconfigure(0, weight=1)
question_frame.columnconfigure(0, weight=1)
answer_frame.rowconfigure(0, weight=1)
answer_frame.columnconfigure(0, weight=1)

# Text box creation, placement, and formatting -------------------
question_txt = scrolledtext.ScrolledText(question_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
question_txt.grid(row=0, column=0,   sticky=E+W+N+S)
answer_txt = scrolledtext.ScrolledText(answer_frame, width=40, height=10, font=(fonttype, 20), wrap="word")
answer_txt.grid(row=0, column=0,   sticky=E+W+N+S)

root.mainloop()