from itertools import combinations as com
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import numpy as np

# hlper_fnc function
def hlper_fnc(test_list):
    return tuple(zip(*test_list))

def seed_combination(first_col, sec_col) :
    combi = []
    for first_cell in all [first_col][1:] :  
        for sec_cell in all[sec_col][1:] :
            if (first_cell.value is not None) and (sec_cell.value is not None) :
                #print(first_cell.value + " " + sec_cell.value)
                combi.append(first_cell.value + " " + sec_cell.value)
            else : 
                break
    return combi

def create_comb_col_indices(indices):
    comb_indices = com(indices, 2)
    comb_indices_arr = [k for k in comb_indices]
    #print(comb_indices_arr)
    return comb_indices_arr


def create_sheet(sheet_name, first_col_ind, second_col_ind):
    ws = wb.create_sheet(str(sheet_name), 0)
    ws.title = str(sheet_name)
    col_indice = col_indices[ind]
    combi = seed_combination(col_indice[0], col_indice[1])
    print(type(combi))

    for row in combi:
        print(row)
        row_list = list()
        row_list.append(row)
        ws.append(row_list)
    

title_col = ["Brand","Topic", "Level"] # <-modify here
x = com(title_col, 2)
col_names = [i for i in x]
print(col_names)



#wb = Workbook()
wb = load_workbook(filename="sample_book.xlsx")
#for z in y:

contentSeed = wb.get_sheet_by_name("Seeds")
all = contentSeed["A:C"] # <-modify here

col_indices = create_comb_col_indices([0, 1, 2]) # <-modify here
ind = 0
for col_name in col_names :
    print(col_name)
    sheet_name = col_name[0] + '_' + col_name[1]
    ws = wb.create_sheet(str(sheet_name), 0)
    ws.title = str(sheet_name)
    col_indice = col_indices[ind]
    combi = seed_combination(col_indice[0], col_indice[1])
    print(type(combi))

    for row in combi:
        print(row)
        row_list = list()
        row_list.append(row)
        ws.append(row_list)
    ind = ind + 1


#col_indices = create_comb_col_indices([0, 1, 2])
#for col_indice in col_indices :
    #combi = seed_comobination(col_indice[0],col_indice[1])
    #print(combi)


wb.save(filename = 'sample_book.xlsx')
