class Collaborator:
    def __init__(self, fullName, firstEntry, firstExit, secondEntry, secondExit):
        self.fullName = fullName
        self.firstEntry = firstEntry
        self.firstExit = firstExit
        self.secondEntry = secondEntry
        self.secondExit = secondExit
    
    def __str__(self):
        return f"Nome do colaborador: {self.fullName}: \n1ª Entrada: {self.firstEntry}, \n1ª Saída: {self.firstExit}, \n2ª Entrada: {self.secondEntry}, \n2ª Saída: {self.secondExit}"
        