import csv
from tkinter import *
from tkinter import ttk
from tkinter import filedialog


def decimal_to_float(dec_num):
    if len(dec_num) > 0:
        dec_num_parts = dec_num.split('E')
        number = float(dec_num_parts[0])
        power = float(dec_num_parts[1])
        return number * (10**power)


def point_to_comma(number_str):
    if number_str.count(',') == 1:
        number_str = number_str.replace(',', '.', 1)
    return number_str


def im_re_Z_to_sigma(re_Z, im_Z):
    sigma = 400 / (re_Z + (im_Z**2)/re_Z)
    return sigma


def open_file():
    global data
    filepath = filedialog.askopenfilename()
    print(filepath)
    if filepath != "":
        with open(filepath, "r") as file:
            text = file.read()
            text_editor.delete("1.0", END)
            text_editor.insert(END, text)

        with open(filepath, "r") as file:
            rows = file.readlines()

            for i in range(len(rows)):
                data.append([])
            
            for i in range(len(rows)):
                row_list = rows[i].strip().split(';')
                if 'freq' in row_list[0]:
                    if len(data[i]) == 0:
                        data[i].append("freq[Hz]")
                        data[i].append("Sigma")
                    continue

                re_Z = decimal_to_float(point_to_comma(row_list[3]))
                im_Z = decimal_to_float(point_to_comma(row_list[4]))
                sigma = im_re_Z_to_sigma(re_Z, im_Z)

                sigma_excel = str(sigma).replace('.', ',', 1)

                if len(data[i]) == 0:
                    freq_excel = str(decimal_to_float(
                        point_to_comma(row_list[0]))).replace('.', ',', 1)
                    data[i].append(freq_excel)
                data[i].append(sigma_excel)


def save_file():
    global data
    filepath = filedialog.asksaveasfilename()
    if filepath != "":
        with open(filepath, "w", newline='') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerows(data)

        with open(filepath, "r") as file:
            text = file.read()

    text_editor.delete("1.0", END)
    text_editor.insert(END, text)



root = Tk()
root.title("Impedance spectrometry")
root.geometry('650x450')

def close():
    root.destroy()
    root.quit()

text_editor = Text()
text_editor.grid(row=0, column=0)

data = []

button_open = ttk.Button(root, text="Open and process", command=open_file)
button_open.grid(row=1, column=0)

button_save = ttk.Button(root, text="Get the spectrum and save .csv", command=save_file)
button_save.grid(row=2, column=0)


root.protocol('WM_DELETE_WINDOW', close)
root.mainloop()