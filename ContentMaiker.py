import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from tkinter import ttk
import codecs
import csv
import random
from copy import deepcopy
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import PageBreak
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import os
from pydub import AudioSegment
import numpy as np

class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        # Initialyze the frame
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title(u'Мамусе От Димуси! Делатель Арифметуси!')
        self.parent.geometry('500x500')

        # Some predifined data
        ThemeList = {
            1: u'Прямое сложение и вычитание',
            2: u'Младшие товарищи',
            3: u'Старшие товарищи (+)',
            4: u'Микс формулы (+)',
            5: u'Старшие товарищи (-)',
            6: u'Микс формулы (-)',
        }
        self.Themes = {
            1: {
                    'FD': {0:[],1:[],2:[],3:[],4:[],5:[],6:[],7:[],8:[],9:[]},
                    'OD': {0:[],1:[],2:[],3:[],4:[],5:[],6:[],7:[],8:[],9:[]},
                    'Subthemes': [1,2,3,4,5,6,7,8,9],
                    'SubthemesOD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[-1,1,2,3,5,6,7,8],
                        2:[-2,-1,1,2,5,6,7],
                        3:[-3,-2,-1,1,5,6],
                        4:[-4,-3,-2,-1,5],
                        5:[-5,1,2,3,4],
                        6:[-6,-5,-1,1,2,3],
                        7:[-7,-6,-5,-2,-1,1,2],
                        8:[-8,-7,-6,-5,-3,-2,-1,1],
                        9:[-9,-8,-7,-6,-5,-4,-3,-2,-1],
                    },
                    'SubthemesFD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[1,2,3,5,6,7,8],
                        2:[-1,1,2,5,6,7],
                        3:[-2,-1,1,5,6],
                        4:[-3,-2,-1,5],
                        5:[1,2,3,4],
                        6:[-5,-1,1,2,3],
                        7:[-6,-5,-2,-1,1,2],
                        8:[-7,-6,-5,-3,-2,-1,1],
                        9:[-8,-7,-6,-5,-4,-3,-2,-1],
                    }
                },
            2: {
                    'FD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[1,2,3,5,6,7,8],
                        2:[-1,1,2,5,6,7],
                        3:[-2,-1,1,5,6],
                        4:[-3,-2,-1,5],
                        5:[1,2,3,4],
                        6:[-5,-1,1,2,3],
                        7:[-6,-5,-2,-1,1,2],
                        8:[-7,-6,-5,-3,-2,-1,1],
                        9:[-8,-7,-6,-5,-4,-3,-2,-1],
                    },
                    'OD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[-1,1,2,3,5,6,7,8],
                        2:[-2,-1,1,2,5,6,7],
                        3:[-3,-2,-1,1,5,6],
                        4:[-4,-3,-2,-1,5],
                        5:[-5,1,2,3,4],
                        6:[-6,-5,-1,1,2,3],
                        7:[-7,-6,-5,-2,-1,1,2],
                        8:[-8,-7,-6,-5,-3,-2,-1,1],
                        9:[-9,-8,-7,-6,-5,-4,-3,-2,-1],
                    },
                    'Subthemes': [1,2,3,4],
                    'SubthemesOD': {
                        0:[],
                        1:[4],
                        2:[3,4],
                        3:[2,3,4],
                        4:[1,2,3,4],
                        5:[-4,-3,-2,-1],
                        6:[-4,-3,-2],
                        7:[-4,-3],
                        8:[-4],
                        9:[],
                    },
                    'SubthemesFD': {
                        0:[],
                        1:[4],
                        2:[3,4],
                        3:[2,3,4],
                        4:[1,2,3,4],
                        5:[-4,-3,-2,-1],
                        6:[-4,-3,-2],
                        7:[-4,-3],
                        8:[-4],
                        9:[],
                    }
                },
            3: {
                    'FD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[1,2,3,4,5,6,7,8],
                        2:[-1,1,2,3,4,5,6,7],
                        3:[-2,-1,1,2,3,4,5,6],
                        4:[-3,-2,-1,1,2,3,4,5],
                        5:[-4,-3,-2,-1,1,2,3,4],
                        6:[-5,-4,-3,-2,-1,1,2,3],
                        7:[-6,-5,-4,-3,-2,-1,1,2],
                        8:[-7,-6,-5,-4,-3,-2,-1,1],
                        9:[-8,-7,-6,-5,-4,-3,-2,-1],
                    },
                    'OD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[-1,1,2,3,4,5,6,7,8],
                        2:[-2,-1,1,2,3,4,5,6,7],
                        3:[-3,-2,-1,1,2,3,4,5,6],
                        4:[-4,-3,-2,-1,1,2,3,4,5],
                        5:[-5,-4,-3,-2,-1,1,2,3,4],
                        6:[-6,-5,-4,-3,-2,-1,1,2,3],
                        7:[-7,-6,-5,-4,-3,-2,-1,1,2],
                        8:[-8,-7,-6,-5,-4,-3,-2,-1,1],
                        9:[-9,-8,-7,-6,-5,-4,-3,-2,-1],
                    },
                    'Subthemes': [1,2,3,4,5,6,7,8,9],
                    'SubthemesOD': {
                        0:[],
                        1:[9],
                        2:[8,9],
                        3:[7,8,9],
                        4:[6,7,8,9],
                        5:[5],
                        6:[4,5,9],
                        7:[3,4,5,8,9],
                        8:[2,3,4,5,7,8,9],
                        9:[1,2,3,4,5,6,7,8,9],
                    },
                    'SubthemesFD': {
                        0:[],
                        1:[],
                        2:[],
                        3:[],
                        4:[],
                        5:[],
                        6:[],
                        7:[],
                        8:[],
                        9:[],
                    }
                },
            4: {
                    'FD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[1,2,3,4,5,6,7,8],
                        2:[-1,1,2,3,4,5,6,7],
                        3:[-2,-1,1,2,3,4,5,6],
                        4:[-3,-2,-1,1,2,3,4,5],
                        5:[-4,-3,-2,-1,1,2,3,4],
                        6:[-5,-4,-3,-2,-1,1,2,3],
                        7:[-6,-5,-4,-3,-2,-1,1,2],
                        8:[-7,-6,-5,-4,-3,-2,-1,1],
                        9:[-8,-7,-6,-5,-4,-3,-2,-1],
                    },
                    'OD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[-1,1,2,3,4,5,6,7,8,9],
                        2:[-2,-1,1,2,3,4,5,6,7,8,9],
                        3:[-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        4:[-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        5:[-5,-4,-3,-2,-1,1,2,3,4,5],
                        6:[-6,-5,-4,-3,-2,-1,1,2,3,4,5,9],
                        7:[-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,8,9],
                        8:[-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,7,8,9],
                        9:[-9,-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                    },
                    'Subthemes': [6,7,8,9],
                    'SubthemesOD': {
                        0:[],
                        1:[],
                        2:[],
                        3:[],
                        4:[],
                        5:[6,7,8,9],
                        6:[6,7,8],
                        7:[6,7],
                        8:[6],
                        9:[],
                    },
                    'SubthemesFD': {
                        0:[],
                        1:[],
                        2:[],
                        3:[],
                        4:[],
                        5:[],
                        6:[],
                        7:[],
                        8:[],
                        9:[],
                    }
                },
            5: {
                    'FD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[1,2,3,4,5,6,7,8],
                        2:[-1,1,2,3,4,5,6,7],
                        3:[-2,-1,1,2,3,4,5,6],
                        4:[-3,-2,-1,1,2,3,4,5],
                        5:[-4,-3,-2,-1,1,2,3,4],
                        6:[-5,-4,-3,-2,-1,1,2,3],
                        7:[-6,-5,-4,-3,-2,-1,1,2],
                        8:[-7,-6,-5,-4,-3,-2,-1,1],
                        9:[-8,-7,-6,-5,-4,-3,-2,-1],
                    },
                    'OD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[-1,1,2,3,4,5,6,7,8,9],
                        2:[-2,-1,1,2,3,4,5,6,7,8,9],
                        3:[-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        4:[-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        5:[-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        6:[-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        7:[-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        8:[-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        9:[-9,-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                    },
                    'Subthemes': [-1,-2,-3,-4,-5,-6,-7,-8,-9],
                    'SubthemesOD': {
                        0:[-9,-8,-7,-6,-5,-4,-3,-2,-1],
                        1:[-9,-8,-7,-5,-4,-3,-2],
                        2:[-9,-8,-5,-4,-3],
                        3:[-9,-5,-4],
                        4:[-5],
                        5:[-9,-8,-7,-6],
                        6:[-9,-8,-7],
                        7:[-9,-8],
                        8:[-9],
                        9:[],
                    },
                    'SubthemesFD': {
                        0:[],
                        1:[],
                        2:[],
                        3:[],
                        4:[],
                        5:[],
                        6:[],
                        7:[],
                        8:[],
                        9:[],
                    }
                },
            6: {
                    'FD': {
                        0:[1,2,3,4,5,6,7,8,9],
                        1:[1,2,3,4,5,6,7,8],
                        2:[-1,1,2,3,4,5,6,7],
                        3:[-2,-1,1,2,3,4,5,6],
                        4:[-3,-2,-1,1,2,3,4,5],
                        5:[-4,-3,-2,-1,1,2,3,4],
                        6:[-5,-4,-3,-2,-1,1,2,3],
                        7:[-6,-5,-4,-3,-2,-1,1,2],
                        8:[-7,-6,-5,-4,-3,-2,-1,1],
                        9:[-8,-7,-6,-5,-4,-3,-2,-1],
                    },
                    'OD': {
                        0:[-9,-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        1:[-9,-8,-7,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        2:[-9,-8,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        3:[-9,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        4:[-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        5:[-9,-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        6:[-9,-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        7:[-9,-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        8:[-9,-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                        9:[-9,-8,-7,-6,-5,-4,-3,-2,-1,1,2,3,4,5,6,7,8,9],
                    },
                    'Subthemes': [-6,-7,-8,-9],
                    'SubthemesOD': {
                        0:[],
                        1:[-6],
                        2:[-7,-6],
                        3:[-8,-7,-6],
                        4:[-9,-8,-7,-6],
                        5:[],
                        6:[],
                        7:[],
                        8:[],
                        9:[],
                    },
                    'SubthemesFD': {
                        0:[],
                        1:[],
                        2:[],
                        3:[],
                        4:[],
                        5:[],
                        6:[],
                        7:[],
                        8:[],
                        9:[],
                    }
                }
        }
        self.PauseListC = []
        self.Intro = self.MakeIntro()

        # SELECT THEME SECTION
        self.Theme = tk.StringVar(self.parent)
        self.Theme.trace('w', self.ThemeChanged)
        choices = [u'',u'Прямое сложение и вычитание',u'Младшие товарищи',u'Старшие товарищи (+)',u'Микс формулы (+)',u'Старшие товарищи (-)',u'Микс формулы (-)']
        self.ThemeID = {}
        for ID in range(1,7):
            self.ThemeID[ThemeList[ID]] = ID
        tk.Label(self.parent, text="Выбери тему:").pack(side = 'top')
        popupMenu = ttk.OptionMenu(self.parent, self.Theme, *choices)
        popupMenu.pack(side = 'top')

        # SPECIFY THEME
        tk.Label(self.parent, text="Выбери подтему:").pack(side = 'top')
        self.subtheme = tk.Frame(self.parent)
        self.subtheme.pack(side = 'top')
        self.subtheme_container = tk.Frame(self.subtheme)
        self.subtheme_container.pack(side = 'left', fill='x')
        # Restore tool container and fill it with tools

        # SELECT DIMENSION
        tk.Label(self.parent, text="Выбери количество десятичных знаков (1..5):").pack(side = 'top')
        self.DIM = tk.StringVar()
        self.DIM.set('1')
        Entry = tk.Entry(self.parent, textvariable=self.DIM)
        Entry.pack(side = 'top')

        # SELECT LENGTH
        tk.Label(self.parent, text="Выбери количество слогаемых (2..25):").pack(side = 'top')
        self.LENGTH = tk.StringVar()
        self.LENGTH.set('5')
        Entry = tk.Entry(self.parent, textvariable=self.LENGTH)
        Entry.pack(side = 'top')

        # SELECT NTABLES
        tk.Label(self.parent, text="Выбери количество таблиц (1..50):").pack(side = 'top')
        self.NTABLES = tk.StringVar()
        self.NTABLES.set('1')
        Entry = tk.Entry(self.parent, textvariable=self.NTABLES)
        Entry.pack(side = 'top')

        tk.Label(self.parent, text="----------------------- ПАРАМЕТРЫ ДЛЯ ОЗВУЧКИ ------------------").pack(side = 'top')

        # правило первого знака
        self.OZ = tk.IntVar(value=1)
        CheckBox = tk.Checkbutton(self.parent,
                                  text = 'Делать озвучку?',
                                  variable = self.OZ)
        CheckBox.pack(side = 'top')

        # правило первого знака
        self.PPZ = tk.IntVar(value=1)
        CheckBox = tk.Checkbutton(self.parent,
                                  text = 'Правило первого знака',
                                  variable = self.PPZ)
        CheckBox.pack(side = 'top')

        # SELECT PACE
        tk.Label(self.parent, text="Выбери длину паузы (0..5000):").pack(side = 'top')
        AddPauseSection = tk.Frame(self.parent)
        AddPauseSection.pack(side = 'top')
        self.PACE = tk.StringVar()
        self.PACE.set('1000')
        Entry = tk.Entry(AddPauseSection, textvariable=self.PACE)
        Entry.pack(side = 'left')
        AddPauseButton = tk.Button(AddPauseSection,
                            text='Добавить',
                            command= self.AddPause)
        AddPauseButton.pack(side = 'left')
        ResetButton = tk.Button(AddPauseSection,
                            text='Очистить',
                            command= self.ResetPause)
        ResetButton.pack(side = 'left')

        # PACE LIST
        tk.Label(self.parent, text="Cписок пауз:").pack(side = 'top')
        self.PauseList = tk.Frame(self.parent)
        self.PauseList.pack(side = 'top')
        self.PauseListContainer = tk.Frame(self.PauseList)
        tk.Label(self.PauseListContainer, text="Пока что пусто").pack(side = 'top')
        self.PauseListContainer.pack(side = 'top')

        # MAKE TABLE
        self.btn_text = tk.StringVar()
        self.btn_text.set('Сгенерировать таблицу')
        self.MakeTableButton = tk.Button(self.parent,
                            textvariable=self.btn_text,
                            command= self.StartAction)
        self.MakeTableButton.pack(side = 'top')

    def MakeIntro(self):
        VOICEFOLDER = 'voice/'
        S1 = AudioSegment.silent(100)
        S2 = AudioSegment.silent(1000)
        S3 = AudioSegment.silent(3000)
        Z = AudioSegment.from_mp3(VOICEFOLDER + 'zapisali.mp3')
        P1 = AudioSegment.from_mp3(VOICEFOLDER + 'P1.mp3')
        P2 = AudioSegment.from_mp3(VOICEFOLDER + 'P2.mp3')
        P3 = AudioSegment.from_mp3(VOICEFOLDER + 'P3.mp3')
        P4 = AudioSegment.from_mp3(VOICEFOLDER + 'P4.mp3')
        P5 = AudioSegment.from_mp3(VOICEFOLDER + 'P5.mp3')
        P6 = AudioSegment.from_mp3(VOICEFOLDER + 'P6.mp3')
        P7 = AudioSegment.from_mp3(VOICEFOLDER + 'P7.mp3')
        P8 = AudioSegment.from_mp3(VOICEFOLDER + 'P8.mp3')
        P9 = AudioSegment.from_mp3(VOICEFOLDER + 'P9.mp3')
        P10 = AudioSegment.from_mp3(VOICEFOLDER + 'P10.mp3')
        I1 = S1 + P1 + S2
        I2 = Z + S3 + P2 + S2
        I3 = Z + S3 + P3 + S2
        I4 = Z + S3 + P4 + S2
        I5 = Z + S3 + P5 + S2
        I6 = Z + S3 + P6 + S2
        I7 = Z + S3 + P7 + S2
        I8 = Z + S3 + P8 + S2
        I9 = Z + S3 + P9 + S2
        I10 = Z + S3 + P10 + S2
        I11 = Z + S3
        Intro = [I1,I2,I3,I4,I5,I6,I7,I8,I9,I10,I11]
        return Intro

    def AddPause(self):
        try:
            pace = int(self.PACE.get())
            if not((pace<0) or (pace>5000)):
                self.PauseListC.append(pace)
                self.PauseListContainer.destroy()
                self.PauseListContainer = tk.Frame(self.PauseList)
                for pace in self.PauseListC:
                    tk.Label(self.PauseListContainer, text=str(pace)).pack(side = 'top')
                self.PauseListContainer.pack(side = 'top')
            else:
                self.PACE.set('1000')
        except:
            self.PACE.set('1000')

    def ResetPause(self):
        self.PauseListC = []
        self.PauseListContainer.destroy()
        self.PauseListContainer = tk.Frame(self.PauseList)
        tk.Label(self.PauseListContainer, text="Пока что пусто").pack(side = 'top')
        self.PauseListContainer.pack(side = 'top')

    def ThemeChanged(self, var, indx, mode):
        self.subtheme_container.destroy()
        self.subtheme_container = tk.Frame(self.subtheme)
        self.subtheme_container.pack(side = 'left', fill='x')
        theme = self.Theme.get()
        themeNumber = self.ThemeID[theme]
        self.SubThemeID = {
            1: tk.IntVar(value=1),
            2: tk.IntVar(value=1),
            3: tk.IntVar(value=1),
            4: tk.IntVar(value=1),
            5: tk.IntVar(value=1),
            6: tk.IntVar(value=1),
            7: tk.IntVar(value=1),
            8: tk.IntVar(value=1),
            9: tk.IntVar(value=1),
            -1: tk.IntVar(value=1),
            -2: tk.IntVar(value=1),
            -3: tk.IntVar(value=1),
            -4: tk.IntVar(value=1),
            -5: tk.IntVar(value=1),
            -6: tk.IntVar(value=1),
            -7: tk.IntVar(value=1),
            -8: tk.IntVar(value=1),
            -9: tk.IntVar(value=1),
        }
        # Constract checkboxes for all subthemes
        for subtheme in self.Themes[themeNumber]['Subthemes']:
            CheckBox = tk.Checkbutton(self.subtheme_container,
                                  text = str(subtheme),
                                  variable = self.SubThemeID[subtheme])
            CheckBox.pack(side = 'left')


    def CheckInputValidity(self):

        Validity = True
        self.results = {}

        ############# Theme ##############
        try:
            theme = self.ThemeID[self.Theme.get()]
            self.results['theme'] = theme
        except:
            Validity = False


        ############# Dimension ##############
        try:
            dim = int(self.DIM.get())
            self.results['dim'] = dim
            if (dim<1) or (dim>5):
                Validity = False
        except:
            Validity = False


        ############# Length ##############
        try:
            length = int(self.LENGTH.get())
            self.results['length'] = length
            if (length<2) or (length>25):
                Validity = False
        except:
            Validity = False


        ############# Ntables ##############
        try:
            nt = int(self.NTABLES.get())
            self.results['ntables'] = nt
            if (nt<1) or (nt>50):
                Validity = False
        except:
            Validity = False

        self.results['PPZ'] = self.PPZ.get()
        self.results['OZ'] = self.OZ.get()
        if not Validity:
            self.results = {}

        return Validity


    def StartAction(self):
        check = self.CheckInputValidity()
        if check:
            self.makeTable()
        else:
            messagebox.showerror('Ошибка!', 'Проверь входные данные!')


    def digitize(self, value):
        digits = []
        sum = value
        for i in range(self.results['dim']):
            digits.append(sum // 10**(self.results['dim']-1-i))
            sum = sum % 10**(self.results['dim']-1-i)
        return digits


    def allowedSum(self, value):
        if value >= 0:
            allowed = True
            onlyPositive = False
            onlyNegative = False
            digits = self.digitize(value)
            # check all possible additives for sign conflict
            for i in range(len(digits)):
                if i == 0:
                    possibleAdd = self.AllowedFD[digits[i]]
                else:
                    possibleAdd = self.AllowedOD[digits[i]]
                positive = False
                negative = False
                for j in possibleAdd:
                    if j<0: negative = True
                    else: positive = True
                if not onlyPositive and not negative:
                    onlyPositive = True
                if not onlyNegative and not positive:
                    onlyNegative = True
            if onlyPositive and onlyNegative:
                allowed = False
        else:
            allowed = False
        return allowed


    def computeAdd(self, value, oldAdd):
        digits = self.digitize(value)
        # collect for each digit possible additives
        PosAddList = []
        NegAddList = []
        onlyPositive = False
        onlyNegative = False
        for i in range(len(digits)):
            if i == 0:
                possibleAdd = self.AllowedFD[digits[i]]
            else:
                possibleAdd = self.AllowedOD[digits[i]]
            posBuffer = []
            negBuffer = []
            for j in possibleAdd:
                if j < 0:
                    negBuffer.append(j)
                else:
                    posBuffer.append(j)
            PosAddList.append(posBuffer)
            NegAddList.append(negBuffer)
        # generate additive
        add = -999999
        while ((len(str(value + add)) != self.results['dim']) or (not self.allowedSum(value + add)) or (add == oldAdd)):
            error = False
            # If negative
            AddVector = []
            add = 0
            sign = random.randint(1,2)
            if sign == 1:
                # construct negative add
                for i in range(len(PosAddList)):
                    if len(PosAddList[i]) > 1:
                        RandIndex = random.randint(0,len(PosAddList[i])-1)
                        AddVector.append(PosAddList[i][RandIndex])
                    elif len(PosAddList[i]) == 1:
                        AddVector.append(PosAddList[i][0])
                    else:
                        error = True
            else:
                # construct positive add
                for i in range(len(NegAddList)):
                    if len(NegAddList[i]) > 0:
                        RandIndex = random.randint(0,len(NegAddList[i])-1)
                        AddVector.append(NegAddList[i][RandIndex])
                    elif len(NegAddList[i]) == 1:
                        AddVector.append(NegAddList[i][0])
                    else:
                        error = True
            if not error:
                for i in range(len(AddVector)):
                    add += AddVector[i] * 10 ** (self.results['dim']-1-i)
            else:
                add = -999999
        return add


    def MakePrimer(self):
        primer = []
        StartInt = random.randint(10**(self.results['dim']-1),10**self.results['dim']-1)
        while not self.allowedSum(StartInt):
            StartInt = random.randint(10**(self.results['dim']-1),10**self.results['dim']-1)
        runningSum = StartInt
        primer.append(runningSum)
        previosAdd = runningSum
        for i in range(self.results['length'] - 1):
            add = self.computeAdd(runningSum, previosAdd)
            runningSum += add
            previosAdd = add
            primer.append(add)
        primer.append(runningSum)
        return primer


    def GetName(self,Primer,Sign,i):
        if i == 0:
            name = 'voice/'+str(Primer[i])+'.mp3'
        elif i == 1:
            if Sign[i] == -1:
                name = 'voice/'+str(Primer[i])+'.mp3'
            else:
                name = 'voice/+'+str(Primer[i])+'.mp3'
        else:
            if Sign[i] == Sign[i-1]:
                if Sign[i] == -1:
                    name = 'voice/'+str(Primer[i])[1:]+'.mp3'
                else:
                    name = 'voice/'+str(Primer[i])+'.mp3'
            else:
                if Sign[i] == -1:
                    name = 'voice/'+str(Primer[i])+'.mp3'
                else:
                    name = 'voice/+'+str(Primer[i])+'.mp3'
        return name


    def GetNameNew(self,Primer,Sign,i):
        if i == 0:
            name = 'voice/'+str(Primer[i])+'.mp3'
        else:
            if Sign[i] == -1:
                name = 'voice/'+str(Primer[i])+'.mp3'
            else:
                name = 'voice/+'+str(Primer[i])+'.mp3'
        return name


    def MakeVoicePrimer(self,Primer,Interval):

        # TODO Add support for nubers larger then 999

        # Convert str to int
        primerInt = []
        for i in Primer:
            if int(i) != 0:
                primerInt.append(int(i))
        pause = AudioSegment.silent(Interval)
        Sign = np.sign(primerInt)
        recording = AudioSegment.silent(10)
        for i in range(len(primerInt)-1):
            if self.results['PPZ']:
                name = self.GetName(primerInt,Sign,i)
            else:
                name = self.GetNameNew(primerInt,Sign,i)
            recording += AudioSegment.from_mp3(name)
            recording += pause
        return recording


    def computeTable(self,FolderPath,tableID):
        tablePath = FolderPath + '/table_' + str(tableID+1) + '.txt'
        with open(tablePath, 'w') as file:
            writer = csv.writer(file)
            table = []
            for i in range(10):
                primer = self.MakePrimer()
                writer.writerow(primer)
                table.append(primer)
        if self.results['OZ']:
            keyFileExists = False
            keys = []
            for interval in self.PauseListC:
                PN = 1
                recording = self.Intro[0]
                for row in table:
                    key = u'Ответ на пример #{} : {}'.format(PN,row[-1])
                    keys.append(key)
                    Primer = self.MakeVoicePrimer(row,interval)
                    recording = recording + Primer + self.Intro[PN]
                    PN += 1
                SavePath = FolderPath + '/table_' + str(tableID+1) + '_' + str(interval) + '.mp3'
                recording.export(SavePath, format="mp3")
                if not keyFileExists:
                    keyFileName = FolderPath + '/table_' + str(tableID+1) + '_otveti.txt'
                    with codecs.open(keyFileName, 'w', 'utf-8') as file:
                        for i in keys:
                            file.write(i+u'\n')
                    keyFileExists = True
        return table


    def makeTable(self):
        FilePath = filedialog.asksaveasfilename(initialdir = 'DATA/',
                                                defaultextension="",
                                                filetypes=[('all files', '.*')],
                                                title="Choose location")
        if FilePath:
            self.btn_text.set("Я Работаю! Не нажимай!")
            keyPos = FilePath.find('DATA/')
            FolderPath = FilePath[keyPos:]
            os.mkdir(FolderPath)
            theme = self.Theme.get()
            themePar = u'Тема: ' + theme + u'. Подтемы: '
            themeNumber = self.ThemeID[theme]
            # Copy the inital allowed additives
            self.AllowedFD = deepcopy(self.Themes[themeNumber]['FD'])
            self.AllowedOD = deepcopy(self.Themes[themeNumber]['OD'])
            # Enrich additives according to the subtheme
            for subtheme in self.Themes[themeNumber]['Subthemes']:
                if self.SubThemeID[subtheme].get() == 1:
                    themePar += str(subtheme) + ','
                    for item in self.Themes[themeNumber]['SubthemesFD'].items():
                        if subtheme in item[1]:
                            self.AllowedFD[item[0]].append(subtheme)
                            self.AllowedFD[item[0]].append(subtheme)
                            self.AllowedFD[item[0]].append(subtheme)
                            self.AllowedFD[item[0]].append(subtheme)
                            self.AllowedFD[item[0]].append(subtheme)
                        if (themeNumber<3) and (subtheme*(-1) in item[1]):
                            self.AllowedFD[item[0]].append(subtheme*(-1))
                            self.AllowedFD[item[0]].append(subtheme*(-1))
                            self.AllowedFD[item[0]].append(subtheme*(-1))
                            self.AllowedFD[item[0]].append(subtheme*(-1))
                            self.AllowedFD[item[0]].append(subtheme*(-1))
                    for item in self.Themes[themeNumber]['SubthemesOD'].items():
                        if subtheme in item[1]:
                            self.AllowedOD[item[0]].append(subtheme)
                            self.AllowedOD[item[0]].append(subtheme)
                            self.AllowedOD[item[0]].append(subtheme)
                            self.AllowedOD[item[0]].append(subtheme)
                            self.AllowedOD[item[0]].append(subtheme)
                        if (themeNumber<3) and (subtheme*(-1) in item[1]):
                            self.AllowedOD[item[0]].append(subtheme*(-1))
                            self.AllowedOD[item[0]].append(subtheme*(-1))
                            self.AllowedOD[item[0]].append(subtheme*(-1))
                            self.AllowedOD[item[0]].append(subtheme*(-1))
                            self.AllowedOD[item[0]].append(subtheme*(-1))
            # Initialyze pdf overview file
            pdfPath = FolderPath + '/tables.pdf'
            doc = SimpleDocTemplate(pdfPath, pagesize=letter)
            # container for the 'Flowable' objects
            elements = []
            # Generate the table
            for tableID in range(self.results['ntables']):
                table = self.computeTable(FolderPath,tableID)
                trans_table = [[1,2,3,4,5,6,7,8,9,10]]
                for i in range(self.results['length']+1):
                    buffer = []
                    for j in range(10):
                        buffer.append(table[j][i])
                    trans_table.append(buffer)
                # add data to the table
                data = trans_table
                # style the pdf table
                t=Table(data)#,5*[0.4*inch], 4*[0.4*inch])
                t.setStyle(TableStyle([
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                ('BOX', (0,0), (0,-1), 0.25, colors.black),
                ('BOX', (1,0), (1,-1), 0.25, colors.black),
                ('BOX', (2,0), (2,-1), 0.25, colors.black),
                ('BOX', (3,0), (3,-1), 0.25, colors.black),
                ('BOX', (4,0), (4,-1), 0.25, colors.black),
                ('BOX', (5,0), (5,-1), 0.25, colors.black),
                ('BOX', (6,0), (6,-1), 0.25, colors.black),
                ('BOX', (7,0), (7,-1), 0.25, colors.black),
                ('BOX', (8,0), (8,-1), 0.25, colors.black),
                ('BOX', (9,0), (9,-1), 0.25, colors.black),
                ('BOX', (0,0), (-1,0), 0.25, colors.black),
                ('BOX', (0,-1), (-1,-1), 0.25, colors.black),
                ('TEXTCOLOR', (0,-1), (-1,-1), colors.red),
                ]))
                # write the pdf
                elements.append(deepcopy(t))
                elements.append(PageBreak())
            doc.build(elements)
            self.btn_text.set("Сгенерировать таблицу")

# Main program body
if __name__ == '__main__':

    root = tk.Tk()
    MainApplication(root).pack(side='top', fill='both', expand=True)
    root.mainloop()
