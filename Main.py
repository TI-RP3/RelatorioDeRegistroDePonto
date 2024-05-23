from openpyxl import load_workbook

from Collaborator import Collaborator

def createCollaboratorsList(table_path):
    workbook = load_workbook(table_path)
    worksheet = workbook.active
    
    collaborators = []
    
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value == "Colaborador":
                nomeColaborador = cell.offset(column=1).value # Pega o valor da célula na coluna seguinte (ao lado de "Colaborador")
                
                colaborador = Collaborator(nomeColaborador,firstEntry="",firstExit="",secondEntry="",secondExit="")
                collaborators.append(colaborador) # Insere o novo colaborador ao final da lista de colaboradores
    
    return collaborators

def main():
    
    colaboradores = createCollaboratorsList( # Caminho para a tabela excel
        "./Pontomais_-_Jornada_(01.05.2024_-_14.05.2024)_-_44082604.xlsx"
    )
    
    for c in range(0, colaboradores.__len__()):
        print(colaboradores[c].__str__())

if __name__ == "__main__":
    main()