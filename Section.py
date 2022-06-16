class Section:
    def __init__(self,
                 name,
                 subSections = 3,
                 percentage = 20,
                 headerColor = [255,255,255],
                 textColor = [0,0,0]):
        self.name = name
        self.subSections = subSections
        self.percentage = percentage
        self.headerColor = headerColor
        self.textColor = textColor

    def ColorToHex(self):
        headerRed = "0x{:02x}".format(int(self.headerColor[0]))[2:]
        headerGreen = "0x{:02x}".format(int(self.headerColor[1]))[2:]
        headerBlue = "0x{:02x}".format(int(self.headerColor[2]))[2:]

        textRed = "0x{:02x}".format(int(self.textColor[0]))[2:]
        textGreen = "0x{:02x}".format(int(self.textColor[1]))[2:]
        textBlue = "0x{:02x}".format(int(self.textColor[2]))[2:]
        
        headerHexcode = f"#{headerRed}{headerGreen}{headerBlue}"
        textHexcode = f"#{textRed}{textGreen}{textBlue}"

        return {
            'Header': headerHexcode, 
            'Text': textHexcode}
