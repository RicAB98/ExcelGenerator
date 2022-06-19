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
        f = open(file)
        data = json.load(f)

        dictionary = {
            "fileName": data["fileName"],
            "numberStudents": data["numberStudents"],
            "numberPeriods": data["numberPeriods"]
        }

        sections = []

        for s in data["sections"]:
            try:
                sections.append(Section(
                    s["name"],
                    s["subSections"],
                    s["percentage"],
                    JsonHelper.GetColorToArray(s["backgroundColor"]),
                    JsonHelper.GetColorToArray(s["textColor"])
                ))
            except Exception as e:
                return (False, s["name"] + ":  \n" + str(e))

        dictionary["sections"] = sections

        try:
            dictionary["finalSection"] = FinalSection(
                JsonHelper.GetColorToArray(data["finalSection"]["backgroundColor"]),
                JsonHelper.GetColorToArray(data["finalSection"]["textColor"])),
        except Exception as e:
            return (False, "Avaliação:  \n" + str(e))

        return (True, dictionary)

    def GetColorToJson(color):
        return{
            "red": color[0],
            "green": color[1],
            "blue": color[2],
        }

    def GetColorToArray(color):
        if color["red"] < 0 or color["red"] > 255 or color["green"] < 0 or color["green"] > 255 or color["blue"] < 0 or color["blue"] > 255:
            raise Exception("Cores têm que ter valores entre 0 e 255")
        return [color["red"], color["green"], color["blue"]]