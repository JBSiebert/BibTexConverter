import pandas as pd

class BibTexConverter():
        def __init__(self, path, file, output_name):
                self.path = path
                self.file = file
                self.output_name = output_name


        def loadBiblio(self):
                bib = pd.read_excel((self.path + self.file), engine="odf")
                return bib



        def convert_row_to_bibtex(self, bib, row):
                textDict = dict(zip(bib.columns.values, [value for value in bib.loc[row,:]]))

                notna = [value for value in textDict.values() if not pd.isna(value)]
                names = [list(textDict.keys())[list(textDict.values()).index(item)] for item in notna]

                authorcode = textDict['author'].split()[1].title() + str(textDict['year'])
                bibentry = "@" + textDict['type'] + "{" + authorcode
                for name in names:
                        if name in ("url", "doi"):
                                bibstring = f'{name} = {{{textDict[name]}}}'
                        elif type(textDict[name]) is str:
                                bibstring = f'{name} = {{{textDict[name].title()}}}'
                        elif type(textDict[name]).__module__ == "numpy":
                                bibstring = f'{name} = {{{int(textDict[name])}}}'
                        bibentry = bibentry + ",\n" +  bibstring
                bibentry = bibentry + "\n}"
                return bibentry
        
        def convert_to_bibtex(self):
                f = open((self.path + self.output_name), "w")
                bib = self.loadBiblio()
                for row in range(0, len(bib)):
                        f.write("\n" + self.convert_row_to_bibtex(bib, row))



