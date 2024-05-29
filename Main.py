from openpyxl import load_workbook

from Collaborator import Collaborator

def createCollaboratorsList(worksheet):
    
    collaborators = []
    
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value == "Colaborador":
                nomeColaborador = cell.offset(column=1).value # Pega o valor da célula na coluna seguinte (ao lado de "Colaborador")
                
                colaborador = Collaborator(nomeColaborador,firstEntry="",firstExit="",secondEntry="",secondExit="")
                collaborators.append(colaborador) # Insere o novo colaborador ao final da lista de colaboradores
    
    return collaborators

def createDailyLists(worksheet):
    workbook = load_workbook("./Pontomais_-_Jornada_(01.05.2024_-_14.05.2024)_-_44082604.xlsx")
    worksheet = workbook.active
    for i,row in enumerate(worksheet.iter_rows(min_row=3)):
        if (row.__getitem__(0).value == "Colaborador"):
            nomeColaborador = row.__getitem__(1).value
            print(nomeColaborador)
            skipRow = i + 2 + 3
            
        if (row.__getitem__(0).value == "Data" or 
            row.__getitem__(0).value == "Resumo"):
            skipRow = i = 1  + 3

        if (row.__getitem__(0).value == "TOTAIS"):
            skipRow = i + 2 + 3
        
        for cell in row:            
            """ if cell.value == "Colaborador":
                nomeColaborador = cell.offset(column=1).value
                skipRow = cell.row + 1
                print(nomeColaborador) """
                
            if (
                (cell.column == 6 or 
                cell.column == 7 or 
                cell.column == 8 or 
                cell.column == 9) and
                cell.row > skipRow
            ):
                print(cell.value)

def main():
    workbook = load_workbook("./Pontomais_-_Jornada_(01.05.2024_-_14.05.2024)_-_44082604.xlsx")
    worksheet = workbook.active
    
    colaboradores = createCollaboratorsList(worksheet)
    
    createDailyLists(worksheet)
    
    """ for c in range(0, colaboradores.__len__()):
        print(colaboradores[c].__str__()) """

if __name__ == "__main__":
    main()