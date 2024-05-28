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
    for row in worksheet.iter_rows(min_row=3):
        for cell in row:            
            if cell.value == "Colaborador":
                nomeColaborador = cell.offset(column=1).value
                skipCell = cell.row
                print(nomeColaborador)
            elif (
                cell.value != "Data" and
                cell.value != "TOTAIS" and
                cell.value != "Resumo"
            ):
                if (
                    (cell.column == 6 or 
                    cell.column == 7 or 
                    cell.column == 8 or 
                    cell.column == 9) and
                    cell.row > skipCell + 1
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