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

def createDailyLists(worksheet,collaborators):
    for row in worksheet.iter_rows():
        for cell in row:
            if (
                cell.value == "Data" or 
                cell.value == "TOTAIS" or 
                cell.value == "Resumo"
            ): continue
            
            if cell.value == "Colaborador":
                nomeColaborador = cell.offset(column=1).value 
            
            
            
            
            
                    
                    
                

def main():
    workbook = load_workbook("./Pontomais_-_Jornada_(01.05.2024_-_14.05.2024)_-_44082604.xlsx")
    worksheet = workbook.active
    
    colaboradores = createCollaboratorsList(worksheet)
    
    createDailyLists(worksheet,colaboradores)
    
    for c in range(0, colaboradores.__len__()):
        print(colaboradores[c].__str__())

if __name__ == "__main__":
    main()