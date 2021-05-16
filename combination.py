from itertools import combinations as com
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import numpy as np
import spacy
nlp = spacy.load("en_core_web_sm") # <--change-here


# hlper_fnc function
def hlper_fnc(test_list):
    return tuple(zip(*test_list))

def seed_combination(first_col, sec_col) :
    combi = []
    for first_cell in all [first_col][1:] :  
        for sec_cell in all[sec_col][1:] :
            if (first_cell.value is not None) and (sec_cell.value is not None) :
                combi.append(first_cell.value + " " + sec_cell.value)
            else : 
                break
    return combi

def create_comb_col_indices(indices):
    comb_indices = com(indices, 2)
    comb_indices_arr = [k for k in comb_indices]
    return comb_indices_arr


def create_sheet(sheet_name, first_col_ind, second_col_ind):
    ws = wb.create_sheet(str(sheet_name), 0)
    ws.title = str(sheet_name)
    combi = seed_combination(first_col_ind, second_col_ind)

    for row in combi:
        #print(row)
        row_list = list()
        row_list.append(row)
        entities = nlp(row)
        for entity in entities.ents:
            row_list.append(entity.text)
        ws.append(row_list)
    

title_col = ["Brand","Topic", "Level"] # <-modify here
x = com(title_col, 2)
col_names = [i for i in x]
print(col_names)

wb = load_workbook(filename="sample_book.xlsx")

contentSeed = wb.get_sheet_by_name("Seeds")
all = contentSeed["A:C"] # <-modify here

col_indices = create_comb_col_indices([0, 1, 2]) # <-modify here
ind = 0
for col_name in col_names :
    print(col_name)
    col_indice = col_indices[ind]
    sheet_name = col_name[0] + '_' + col_name[1]
    create_sheet(sheet_name, col_indice[0] , col_indice[1] )
    sheet_name = col_name[1] + '_' + col_name[0]
    create_sheet(sheet_name, col_indice[1] , col_indice[0] )
    ind = ind + 1

wb.save(filename = 'sample_book.xlsx')
