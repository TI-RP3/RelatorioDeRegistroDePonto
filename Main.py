from openpyxl import load_workbook

from Collaborator import Collaborator

""" def createCollaboratorsList(worksheet):
    
    collaborators = []
    
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value == "Colaborador":
                nomeColaborador = cell.offset(column=1).value # Pega o valor da célula na coluna seguinte (ao lado de "Colaborador")
                
                colaborador = Collaborator(nomeColaborador,firstEntry="",firstExit="",secondEntry="",secondExit="")
                collaborators.append(colaborador) # Insere o novo colaborador ao final da lista de colaboradores
    
    return collaborators """

def createDailyLists(worksheet):
    general_list = []
    collaborators_info = {}
    
    for i, row in enumerate(worksheet.iter_rows(min_row=3)):
        if row[0].value == "Colaborador":
            nomeColaborador = row[1].value
            collaborators_info[nomeColaborador] = [] # 
        elif row[0].value not in ["Data", "Resumo", "TOTAIS"]:
            firstEntry = row[5].value
            firstExit = row[6].value
            secondEntry = row[7].value
            secondExit = row[8].value
            collaborators_info[nomeColaborador].append(
                Collaborator(nomeColaborador, firstEntry, firstExit, secondEntry, secondExit)
            )
    
    max_entries = max(len(entries) for entries in collaborators_info.values())
    for _ in range(max_entries):
        general_list.append([])

    for collaborator, entries in collaborators_info.items():
        for i, entry in enumerate(entries):
            general_list[i].append(entry)

    return general_list

def main():
    workbook = load_workbook("./Pontomais_-_Jornada_(01.05.2024_-_14.05.2024)_-_44082604.xlsx")
    worksheet = workbook.active
    
    # colaboradores = createCollaboratorsList(worksheet)
    
    general_list = createDailyLists(worksheet)
    
    print("[\n")
    for i, daily_list in enumerate(general_list):
        print(f"\t[ # dia: {i+1}:\n")
        for collaborator in daily_list:
            print(collaborator.__str__() + "\n")
        print("\t],\n")
    print("]")
    
    """ for c in range(0, colaboradores.__len__()):
        print(colaboradores[c].__str__()) """

if __name__ == "__main__":
    main()