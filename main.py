import sys
from PyQt6.QtWidgets import QApplication, QColorDialog, QMainWindow, QFileDialog
from PyQt6 import uic
from Section import Section
from Generator import Generator
from FinalSection import FinalSection
from JsonHelper import JsonHelper

class UI(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("window.ui", self)
        
        self.sections = []
        self.finalSection = FinalSection()

        self.ConnectButtons()
        self.SetUpComboBox()
        
    def ConnectButtons(self):
        self.carregarButton.clicked.connect(self.UploadFile)

        self.nomeTextBox.textChanged.connect(self.UpdateTextBox)
        self.subseccoesSpinBox.valueChanged.connect(self.UpdateSubSections)
        self.percentagemSpinBox.valueChanged.connect(self.UpdatePercentage)

        self.adicionarButton.clicked.connect(self.AddSection)
        self.removerButton.clicked.connect(self.RemoveSection)
        self.cabecalhoButton.clicked.connect(lambda: self.GetColorFromDialog(self.cabecalhoFrame))
        self.textoButton.clicked.connect(lambda: self.GetColorFromDialog(self.textoFrame))

        self.cabecalhoAvaliacaoButton.clicked.connect(lambda: self.GetColorFromDialog(self.cabecalhoAvaliacaoFrame))
        self.textoAvaliacaoButton.clicked.connect(lambda: self.GetColorFromDialog(self.textoAvaliacaoFrame))
        self.gerarButton.clicked.connect(self.Generate)

    def UploadFile(self):
        fname = QFileDialog.getOpenFileName(self, 'Abrir ficheiro', './',"Ficheiro json (*.json)")

        dictionary = JsonHelper.ParseJson(fname[0])

        self.ficheiroTextBox.setText(dictionary["fileName"])
        self.alunosSpinBox.setValue(dictionary["numberStudents"])
        self.periodosSpinBox.setValue(dictionary["numberPeriods"])

        self.sections = dictionary["sections"]
        self.RefreshComboBox()

        self.finalSection = dictionary["finalSection"][0]
        self.UpdateColorFrames(self.cabecalhoAvaliacaoFrame, self.textoAvaliacaoFrame, self.finalSection)

    def UpdateColorFrames(self, backgroundFrame, textFrame, section):
        style = f"background-color:rgb({section.headerColor[0]}, {section.headerColor[1]}, {section.headerColor[2]}); "
        backgroundFrame.setStyleSheet(style)

        style = f"background-color:rgb({section.textColor[0]}, {section.textColor[1]}, {section.textColor[2]}); "
        textFrame.setStyleSheet(style)

    def SetUpComboBox(self):
        section1 = Section("Secção 1")

        self.sections.append(section1)
        self.seccaoComboBox.addItem(section1.name)
        self.seccaoComboBox.currentIndexChanged.connect(self.ComboBoxChanged)

    def UpdateTextBox(self):
        newName = self.nomeTextBox.text()
        self.sections[self.seccaoComboBox.currentIndex()].name = newName
        
        if not self.changingComboBox:
            self.RefreshComboBox()

    def UpdateSubSections(self, i):
        self.sections[self.seccaoComboBox.currentIndex()].subSections = i

    def UpdatePercentage(self, i):
        self.sections[self.seccaoComboBox.currentIndex()].percentage = i

    def GetColorFromDialog(self, frame):
        color = QColorDialog.getColor()
        style = f"background-color:rgb({color.red()}, {color.green()}, {color.blue()}); "
        frame.setStyleSheet(style)

        match frame:
            case self.cabecalhoFrame:
                self.sections[self.seccaoComboBox.currentIndex()].headerColor = [color.red(), color.green(), color.blue()]
            case self.textoFrame:
                self.sections[self.seccaoComboBox.currentIndex()].textColor = [color.red(), color.green(), color.blue()]
            case self.cabecalhoAvaliacaoFrame:
                self.finalSection.headerColor = [color.red(), color.green(), color.blue()]
            case self.textoAvaliacaoFrame:
                self.finalSection.textColor = [color.red(), color.green(), color.blue()]

    def RefreshComboBox(self):
        index = self.seccaoComboBox.currentIndex() 
        
        self.seccaoComboBox.clear()
        self.seccaoComboBox.addItems([s.name for s in self.sections])
        self.seccaoComboBox.setCurrentIndex(index)

    def AddSection(self):
        sectionIndex = self.seccaoComboBox.count()
        name = "Secção " + str(sectionIndex + 1)
        section = Section(name)

        self.sections.append(section)
        self.seccaoComboBox.addItem(section.name)
        self.seccaoComboBox.setCurrentIndex(sectionIndex)

    def RemoveSection(self):
        if self.seccaoComboBox.count() == 1:
            return

        index = self.seccaoComboBox.currentIndex() 
        self.seccaoComboBox.removeItem(index)
        del self.sections[index]
        self.ComboBoxChanged(self.seccaoComboBox.currentIndex())

    def ComboBoxChanged(self, index):
        section = self.sections[index]
        self.changingComboBox = True

        self.nomeTextBox.setText(section.name)
        self.subseccoesSpinBox.setValue(section.subSections)
        self.percentagemSpinBox.setValue(section.percentage)

        style = f"background-color:rgb({section.headerColor[0]}, {section.headerColor[1]}, {section.headerColor[2]}); "
        self.cabecalhoFrame.setStyleSheet(style)

        style = f"background-color:rgb({section.textColor[0]}, {section.textColor[1]}, {section.textColor[2]}); "
        self.textoFrame.setStyleSheet(style)

        self.changingComboBox = False

    def Generate(self):
        if not self.ValidadeInputs():
            return

        self.infoLabel.setText('A gerar...')

        fileName = self.ficheiroTextBox.text()
        numberPeriods = self.periodosSpinBox.value()
        numberStudents = self.alunosSpinBox.value()
        sections = self.sections
        finalSection = self.finalSection

        generator = Generator(fileName, numberPeriods, numberStudents, sections, finalSection)
        generator.Generate()
        generator.GenerateJson()
        self.infoLabel.setText('Criado na pasta \"FicheirosGerados\"')

    def ValidadeInputs(self):
        totalPercentage = 0
        for s in self.sections:
            totalPercentage += s.percentage
        
        if totalPercentage != 100:
            self.infoLabel.setText('Soma de percentagens não é 100%')
            return False

        return True

def main():
    app = QApplication(sys.argv)
    window = UI()
    window.setWindowTitle("Gerador de Excel")
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()