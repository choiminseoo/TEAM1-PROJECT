from tkinter import *
from tkinter import messagebox

root = Tk()
root.title("정렬 방식 선택")
root.geometry("300x200")

mainMenu = Menu(root)
root.config(menu=mainMenu)

fileMenu = Menu(mainMenu)
mainMenu.add_cascade(label="정렬", menu=fileMenu)
fileMenu.add_command(label="넷플릭스")
fileMenu.add_separator()
fileMenu.add_command(label="왓챠")
fileMenu.add_separator()
fileMenu.add_command(label="티빙")
fileMenu.add_separator()
fileMenu.add_command(label="평점순")
fileMenu.add_separator()

root.mainloop()