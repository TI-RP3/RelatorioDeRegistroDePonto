class Collaborator:
    def __init__(self, fullName, firstEntry, firstExit, secondEntry, secondExit):
        self.fullName = fullName
        self.firstEntry = firstEntry
        self.firstExit = firstExit
        self.secondEntry = secondEntry
        self.secondExit = secondExit
    
    def __str__(self):
        return f"\t\tNome do colaborador: {self.fullName}: \n\t\t1ª Entrada: {self.firstEntry}, \n\t\t1ª Saída: {self.firstExit}, \n\t\t2ª Entrada: {self.secondEntry}, \n\t\t2ª Saída: {self.secondExit}"
        