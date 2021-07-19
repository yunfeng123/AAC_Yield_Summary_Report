from tkinter import *


def txt_print(tk_text, tag_name, print_info, length, BG, Font_Name, Font_Size):
    tk_text.insert(END, print_info)
    tk_text_line = tk_text.index(END).split('.')[0]
    if tag_name != '':
        tk_text.tag_add(tag_name, str(int(tk_text_line) - 1) + '.0', str(int(tk_text_line) - 1) + '.' + str(length))
        tk_text.tag_config(tag_name, background=BG, font=(Font_Name, Font_Size))
    tk_text.insert(END, '\n')
    tk_text.see(END)
    tk_text.update()
