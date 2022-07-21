from tkinter import Button
from tkinter import Label
from tkinter import Entry
from sorte import sorte_
from tkinter import Tk

def _label(text:str="", row:int=0, column:int=0, pady:int=0, rowspan:int=1, colspan:int=1) -> None:
    Label(text=text, bg="lightgray", font=("Arial", 15), padx=20, pady=pady) \
        .grid(row=row, column=column, rowspan=rowspan, columnspan=colspan)

def _entry(row:int=0, column:int=0, rowspan:int=1, colspan:int=1) -> Entry:
    _input = Entry(bg="white", font=("Arial", 15), width=10)
    _input.grid(row=row, column=column, rowspan=rowspan, columnspan=colspan)
    return _input

def _button(text:str, script, what:str, row:int=0, column:int=0, rowspan:int=1, colspan:int=1) -> None:
    Button(text=text, command=lambda: script(what, value.get()), font=("Arial", 15), bg="lightblue", width=20) \
        .grid(row=row, column=column, rowspan=rowspan, columnspan=colspan)

window = Tk()
window.title("DeskRobot")
window.config(bg="lightgray")
window.resizable(False, False)

_label()
_label("valor por jogo:", 1, 1)

_label(row=2)
value = _entry(1, 2)

_label(row=2, column=3)
_button("realizar jogo", sorte_, "net", 3, column=1, colspan=2)

_label(row=4)
window.mainloop()