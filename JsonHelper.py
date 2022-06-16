import json
from FinalSection import FinalSection
from Section import Section

class JsonHelper:
    def ConvertTo(properties):
        dictionary = {
            "fileName": properties["fileName"],
            "numberStudents": properties["numberStudents"],
            "numberPeriods": properties["numberPeriods"],
        }

        sections = []

        for s in properties["sections"]:
            sections.append({
                "name": s.name,
                "subSections": s.subSections,
                "percentage": s.percentage,
                "backgroundColor": JsonHelper.GetColorToJson(s.headerColor),
                "textColor": JsonHelper.GetColorToJson(s.textColor),
            })

        dictionary["sections"] = sections
        dictionary["finalSection"] = {
            "backgroundColor": JsonHelper.GetColorToJson(properties["finalSection"].headerColor),
            "textColor": JsonHelper.GetColorToJson(properties["finalSection"].textColor)
        }

        return dictionary

    def ParseJson(file):
        f = open('FicheirosGerados/Grelha.json')
        data = json.load(f)

        dictionary = {
            "fileName": data["fileName"],
            "numberStudents": data["numberStudents"],
            "numberPeriods": data["numberPeriods"]
        }

        sections = []

        for s in data["sections"]:
            sections.append(Section(
                s["name"],
                s["subSections"],
                s["percentage"],
                JsonHelper.GetColorToArray(s["backgroundColor"]),
                JsonHelper.GetColorToArray(s["textColor"])
            ))

        dictionary["sections"] = sections

        dictionary["finalSection"] = FinalSection(
            JsonHelper.GetColorToArray(data["finalSection"]["backgroundColor"]),
            JsonHelper.GetColorToArray(data["finalSection"]["textColor"])),

        return dictionary

    def GetColorToJson(color):
        return{
            "red": color[0],
            "green": color[1],
            "blue": color[2],
        }

    def GetColorToArray(color):
        return [color["red"], color["green"], color["blue"]]