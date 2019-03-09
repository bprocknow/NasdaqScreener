class Person:
    def __init__(self):
        pass
    def setName(self, name):
        self.name = name
    def printName(self):
        print(self.name)
def main():
    newPerson = Person()
    newPerson.setName("Ben")
    newPerson.printName()
if __name__ == "__main__":
    main()
