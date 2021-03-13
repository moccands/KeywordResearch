from itertools import combinations as com
from openpyxl.workbook import Workbook
from openpyxl import load_workbook

def seed_comobination(first_col, sec_col) :
    #shit_name = all[first_col]
    for first_cell in all [first_col][1:] :  
        for sec_cell in all[sec_col][1:] : 
            if (first_cell.value is not None) and (sec_cell.value is not None) :
                print(first_cell.value + " " + sec_cell.value)
            else : 
                break

def create_comb_col_indices(indices):
    comb_indices = com(indices, 2)
    comb_indices_arr = [k for k in comb_indices]
    print(comb_indices_arr)
    return comb_indices_arr

letter = ["Brand","Topic", "ENtiry"]
x = com(letter,2)
sheet_names = [i for i in x]
print(sheet_names)

#wb = Workbook()
wb = load_workbook(filename="sample_book.xlsx")
#for z in y:
for name in sheet_names :
    print(name)
    ws = wb.create_sheet(str(name), 0)
    ws.title = str(name)

contentSeed = wb.get_sheet_by_name("Seeds")
all = contentSeed["A:C"]

ws = wb.create_sheet('mexicano', 0)
wb.save(filename = 'sample_book.xlsx')


col_indices = create_comb_col_indices([0, 1, 2])
for col_indice in col_indices :
    seed_comobination(col_indice[0],col_indice[1])


#for col in ws.iter_rows(min_row=1, max_row=10, min_col=0, max_col=3, values_only=False) :
    #print(col)



wb.save(filename = 'sample_book.xlsx')
