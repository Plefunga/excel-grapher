import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
import os

class Box:
    def __init__(self, width):
        self._width = width
        self._text = "#"*width + "########" # for the padding and border

    def add_text(self, text, center=False):
        if center:
            self._text += "\n"
            self._text += "##  "
            lpad = (self._width - len(text))//2
            rpad = self._width - lpad - len(text)
            self._text += " " * lpad + text + " " * rpad
            self._text += "  ##"

        else:
            lines = []
            words = text.split(" ")
            line = []
            for word in words:
                if len(" ".join(line)) + len(" " + word) <= self._width:
                    line.append(word)
                else:
                    lines.append(line)
                    line = [word]

            # get the last line
            if len(line) != 0:
                lines.append(line)

            for line in lines:
                self._text += "\n##  " + " ".join(line) + " " * (self._width - len(" ".join(line))) + "  ##"

    def add_list(self, items):
        for i in range(len(items)):
            lines = []
            words = items[i].split(" ")
            line = [f"{i + 1}."]
            indent = 0
            for word in words:
                if len(" ".join(line)) + len(" " + word) <= self._width - indent:
                    line.append(word)
                else:
                    lines.append(line)
                    line = [word]
                indent = 3
            # get the last line
            if len(line) != 0:
                lines.append(line)

            indent = 0
            for line in lines:
                self._text += "\n##  " + " " * indent + " ".join(line) + " " * (self._width - len(" ".join(line)) - indent) + "  ##"
                indent = 3


    def __str__(self):
        return self._text + "\n" + "#"*self._width + "########"

    def __repr__(self):
        return self._text + "\n" + "#"*self._width + "########"

class Spreadsheet:
    def __init__(self, dataframe, title):
        self._data = dataframe
        self._data = self._data.reset_index()
        self._title = title

        self.index = 0

        self._rows = {}
        self._x_axis = self._data.columns.values.tolist()[2:]
        for index, row in self._data.iterrows():
            row = row.values.tolist()[1:]
            self._rows[row[0]] = row[1:]

        # create dictionary of columns with the header as keys
        #self._columns = {}
        #for series_name, series in self._data.items():
        #    self._columns[series_name] = series.to_list()

    def __repr__(self):
        return str(self)

    def __str__(self):
        out = ""
        for i in self._rows.keys():
            out += str(i) + "\t" + "\t".join(self._rows[i])
            out += "\n"
        return out
        #for i in self._columns.keys():
        #    out += str(i) + "\t"

        #out += "\n"
        #for i in range(np.nanmax([len(self._columns[j]) for j in self._columns.keys()])):
        #    for j in self._columns.keys():
        #        try:
        #            out += str(self._columns[j][i]) + "\t"
        #        except:
        #            out += "na\t"
        #    out += "\n"
        #return out

    def __getitem__(self, key):
        return self._rows[key]
        #return self._columns[key]

    def __iter__(self):
        return self

    def __next__(self):
        if index < len(self._columns.keys()):
            index += 1
            #return self._columns[self._columns.keys()[index - 1]]
            return self._rows[self._rows.keys()[index - 1]]
        else:
            raise StopIteration

    def plot(self, plot_dir="."):
        # determine scales
        scales = []
        for i in self._rows.keys():
            row = self._rows[i]
            if type(row[0]) == type(0) or type(row[0]) == type(1.1):
                close_enough = False
                for scale_index in range(len(scales)):
                    scale = scales[scale_index]
                    if not math.isnan(np.nanmin(row)) and not math.isnan(np.nanmax(row)):
                        if abs(np.nanmin(row) / scale[0]) < 10 and abs(scale[0] / np.nanmin(row)) and abs(np.nanmax(row) / scale[1]) < 10 and abs(scale[1] / np.nanmax(row)) < 10:
                            close_enough = True
                            scales[scale_index][2].append(i)
                if not close_enough:
                    all_nans = True
                    for element in row:
                        if not math.isnan(element):
                           all_nans = False
                    if not all_nans:
                        if not math.isnan(np.nanmin(row)) and not math.isnan(np.nanmax(row)):
                            scales.append([np.nanmin(row), np.nanmax(row), [i]])
            
        # if length of scales is 2, create 1 plot with 2 axes
        if len(scales) == 2:
            fig, ax1 = plt.subplots()
            ax2 = ax1.twinx()
                
            for i in self._rows.keys():
                row = self._rows[i]
                if type(row[0]) == type(0) or type(row[0]) == type(1.1):
                    if i in scales[0][2]:
                        ax1.plot(self._x_axis, row, label=i)
                    if i in scales[1][2]:
                        ax2.plot(self._x_axis, row, label=i)
            ax1.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05),
                        fancybox=True, shadow=True, ncol=1)
            ax2.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05),
                        fancybox=True, shadow=True, ncol=1)
            fig.tight_layout()
            fig.suptitle(self._title)
            
        elif len(scales) > 0:
            fig, axes = plt.subplots(len(scales),1)

            for i in self._rows.keys():
                row = self._rows[i]
                if type(row[0]) == type(0) or type(row[0]) == type(1.1):
                    for j in range(len(scales)):
                        if i in scales[j][2]:
                            axes[j].plot(self._x_axis, row, label=i)

            for i in axes:
                i.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05),
                          fancybox=True, shadow=True, ncol=1)
            fig.tight_layout()
            fig.suptitle(self._title)
        
        plt.savefig(f"{plot_dir}/{self._title}.png")
        plt.clf()
                

class Data:
    def __init__(self, fname):
        self.fname = fname
        self.dataframe = pd.read_excel(self.fname, sheet_name=None)
        self.spreadsheets = [Spreadsheet(self.dataframe[i], i) for i in self.dataframe]
        self.index = 0

    def __getitem__(self, key):
        if(type(key) == type(0)):
            return self.spreadsheets[key]
        else:
            return self.dataframe[key]

    def __iter__(self):
        return self

    def __next__(self):
        if self.index < len(self.spreadsheets):
            self.index += 1
            return self.spreadsheets[self.index - 1]
        else:
            raise StopIteration

    INFORMATION = Box(47)
    INFORMATION.add_text("SPREADSHEET GRAPHER", center=True)
    INFORMATION.add_text("AUTHOR: NATHAN CARTER", center=True)
    INFORMATION.add_text("VERSION 0.1.1", center=True)
    INFORMATION.add_text("")
    INFORMATION.add_text("")
    INFORMATION.add_text("ABOUT", center=True)
    INFORMATION.add_text("This program will create a graph for each sheet in an Excel spreadsheet. It assumes that your spreadsheet uses the first column as data labels, and the top row as the x-axis.")
    INFORMATION.add_text("")
    INFORMATION.add_text("The name that you give each sheet will be the name of the generated graph.")
    INFORMATION.add_text("")
    INFORMATION.add_text("For each sheet, the program will try and fit all the data onto a single y-axis. However, if there is a difference of more than a factor of 10 in the range of data, it will place that data onto a new plot (or have two y-axis on the one plot if it can).")
    INFORMATION.add_text("")
    INFORMATION.add_text("")
    INFORMATION.add_text("INSRUCTIONS:", center=True)
    INFORMATION.add_list(["Enter the name (or index) of the Excel spreadsheet. If you enter a name, bear in mind it has to be relative to the current directory.","Enter the RELATIVE (meaning relative to the folder that you ran this from) path to the folder where you want the graphs. If you want them in this folder, just leave it blank."])
#######################################################
##                SPREADSHEET GRAPHER                ##
##               AUTHOR: NATHAN CARTER               ##
##                   VERSION 0.1.1                   ##
##                                                   ##
##                                                   ##
##                       ABOUT                       ##
##  This program will create a graph for each sheet  ##
##  in an Excel spreadsheet. It assumes that your    ##
##  spreadsheet uses the first column as data        ##
##  labels, and the top row as the x-axis.           ##
##                                                   ##
##  The name that you give each sheet will be the    ##
##  name of the generated graph.                     ##
##                                                   ##
##  For each sheet, the program will try and fit     ##
##  all the data onto a single y-axis. However, if   ##
##  there is a difference of more than a factor of   ##
##  10 in the range of data, it will place that      ##
##  data onto a new plot (or have two y-axis on the  ##
##  one plot if it can).                             ##
##                                                   ##
##                                                   ##
##                   INSTRUCTIONS:                   ##
##  1. Enter the name (or index) of the Excel        ##
##     spreadsheet.                                  ##
##       -- has to be in this directory.             ##
##  2. Enter the RELATIVE (meaning relative to the   ##
##     folder that you ran this from) path to the    ##
##     folder where you want the graphs. If you      ##
##     want them in this folder, just leave it       ##
##     blank.                                        ##
#######################################################


if __name__ == "__main__":
    print(Data.INFORMATION)

    print("Excel spreadsheets in this directory:")
    
    index = 0
    files = {}
    for file in os.listdir("."):
        if file.endswith(".xlsx"):
            print(f"{index}. {file}")
            files[index] = file
    selection = input("Which spreadsheet: ")
    fname = ""
    try:
        fname = files[int(selection)]
    except:
        fname = selection
    data = Data(fname)
    img_dir = input("Image Directory (leave blank for current directory): ")
    if len(img_dir) == 0:
        img_dir = "."
    
    print("Found sheets:")
    print(", ".join([i for i in data.dataframe]))
    for d in data:
        d.plot(plot_dir=img_dir)
    print("Done.")
