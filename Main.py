import os
from tkinter import Tk, Button, Label, Listbox, END, filedialog, messagebox
from openpyxl import load_workbook

from Collaborator import Collaborator

def getDailyEntries(worksheet, period):
    general_list = []
    collaborators_infos = {}
    days_observed = []
    
    for i, row in enumerate(worksheet.iter_rows(min_row=3)):
        if row[0].value == "Colaborador":
            collab_name = row[1].value
            collaborators_infos[collab_name] = []
        elif (
            ("Seg" in row[0].value or
            "Ter" in row[0].value or
            "Qua" in row[0].value or
            "Qui" in row[0].value or
            "Sex" in row[0].value or
            "Sáb" in row[0].value or
            "Dom" in row[0].value) and
            (i > 1 and i <= period + 3)
        ):
            days_observed.append(row[0].value[5:10].replace("/", "."))                
        elif row[0].value not in ["Data", "Resumo", "TOTAIS"]:            
            first_entry = row[5].value
            first_exit = row[6].value
            second_entry = row[7].value
            second_exit = row[8].value
            collaborators_infos[collab_name].append(
                Collaborator(collab_name, first_entry, first_exit, second_entry, second_exit)
            )
    
    max_entries = max(len(entries) for entries in collaborators_infos.values())
    for _ in range(max_entries):
        general_list.append([])
        
    for i in range(len(days_observed)):
        general_list[i].append(days_observed[i])

    for _, entries in collaborators_infos.items():
        for j, entry in enumerate(entries):
            general_list[j].append(entry)

    return general_list

def readTemplateWorksheet(file_name):
    wb = load_workbook(file_name)
    template_ws = wb.active
    new_ws = wb.copy_worksheet(template_ws)
    new_ws.title = "02.05"
    wb.save(file_name)

def selectFiles(listbox):
    file_paths = filedialog.askopenfilenames(title="Selecione os arquivos .xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_paths:
        for file in file_paths:
            listbox.insert(END, file)

def confirmSelection():
    print("Enviou")
    
def exitProgram(root):
    root.destroy()
    
def main():
    root = Tk()
    root.title("Relatório Digital de Registro de Ponto")
    root.geometry("400x400")
    
    label = Label(root, text='Selecione os arquivos ".xlsx"')
    label.pack(pady=10)
    
    listbox = Listbox(root, width=50, height=15)
    listbox.pack(pady=10)
    
    select_button = Button(root, text="Adicionar arquivos", command=lambda: selectFiles(listbox))
    select_button.pack(pady=5)
    
    confirm_button = Button(root, text="Confirmar", command=confirmSelection)
    confirm_button.pack(pady=5)

    exit_button = Button(root, text="Sair", command=lambda: exitProgram(root))
    exit_button.pack(pady=5)

    root.mainloop()
    
    """ pontoMais_files = selectFiles()
    
    if not pontoMais_files:
        print("Nenhum arquivo selecionado. O programa será encerrado")
        return
    
    print("Arquivos adicionados:")
    for file in pontoMais_files:
        print(f" - {file}") """
    
    #template_wb_name = "./Relatório de Registro de ponto - Maio 2024.xlsx"
    #app_wb_name = "./Pontomais_-_Jornada_(01.05.2024_-_14.05.2024)_-_44082604.xlsx"
    #app_workbook = load_workbook(app_wb_name)
    #app_worksheet = app_workbook.active
    #period_observed = 14
    
    #daily_entries_per_collab = getDailyEntries(app_worksheet, period_observed)
    
    #readTemplateWorksheet(template_wb_name)
    
    """ print("[\n")
    for _, daily_list in enumerate(daily_entries_per_collab):
        print(f"\t[ # dia: ")
        for collaborator in daily_list:
            print(collaborator.__str__() + "\n")
        print("\t],\n")
    print("]") """
    
    ...

if __name__ == "__main__":
    main()