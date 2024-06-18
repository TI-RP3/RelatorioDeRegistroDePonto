from openpyxl import load_workbook

from Collaborator import Collaborator

def createDailyLists(worksheet, period):
    general_list = []
    collaborators_info = {}
    days = []
    
    for i, row in enumerate(worksheet.iter_rows(min_row=3)):
        if row[0].value == "Colaborador":
            nomeColaborador = row[1].value
            collaborators_info[nomeColaborador] = []
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
            days.append(row[0].value[5:10].replace("/", "."))                
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
        
    for j in range(len(days)):
        general_list[j].append(days[j])

    for collaborators, entries in collaborators_info.items():
        for j, entry in enumerate(entries):
            general_list[j].append(entry)

    return general_list

def main():
    periodoObservado = 14
    pontoMaisWorkbook = load_workbook("./Pontomais_-_Jornada_(01.05.2024_-_14.05.2024)_-_44082604.xlsx")
    pontoMaisWorksheet = pontoMaisWorkbook.active
    
    listaDeHorariosPorColaborador = createDailyLists(pontoMaisWorksheet, periodoObservado)
    
    print("[\n")
    for i, daily_list in enumerate(listaDeHorariosPorColaborador):
        print(f"\t[ # dia: ")
        for collaborator in daily_list:
            print(collaborator.__str__() + "\n")
        print("\t],\n")
    print("]")

if __name__ == "__main__":
    main()