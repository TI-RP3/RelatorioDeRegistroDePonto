from openpyxl import load_workbook

from Collaborator import Collaborator

def createCollaboratorsList(table_path):
    workbook = load_workbook(table_path)
    worksheet = workbook.active
    
    colaboradores = []
    
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value == "Colaborador":
                nomeColaborador = cell.offset(column=1).value
                
                colaborador = Collaborator(nomeColaborador,"","","","")
                colaboradores.append(colaborador)
    
    return colaboradores

def main():
    
    createCollaboratorsList( # Caminho para a tabela excel
        "./Pontomais_-_Jornada_(01.05.2024_-_14.05.2024)_-_44082604.xlsx"
    )

if __name__ == "__main__":
    main()