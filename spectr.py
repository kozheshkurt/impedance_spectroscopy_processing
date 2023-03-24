import csv
import os

folder_name = input('Enter folder name: ') + '/'

if len(folder_name) > 1:
    files_list = os.listdir(folder_name)
else:
    files_list = os.listdir()


data = []
empty = True


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


for file_name in files_list:

    if file_name[-7:] != 'new.dat' or file_name == 'data_new.csv' or file_name == folder_name + '.csv':
        continue

    try:
        if len(folder_name) > 1:
            handler = open(folder_name + file_name)
        else:
            handler = open(file_name)
    except:
        print("Error opening ", file_name)
        continue
    print(file_name, "processed")

    rows = handler.readlines()

    if len(data) > 0:
        empty = False
    else:
        for i in range(len(rows)):
            data.append([])

    for i in range(len(rows)):
        row_list = rows[i].strip().split(';')
        if 'freq' in row_list[0]:
            if len(data[i]) == 0:
                data[i].append("freq[Hz]")
            data[i].append("'"+file_name[7:-8])
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
#        print(data[i])

#    print(data)

if len(folder_name) > 1:
    new_file = folder_name + folder_name[:-1] + '.csv'
else:
    new_file = 'data_new.csv'

with open(new_file, 'w', newline='') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerows(data)
