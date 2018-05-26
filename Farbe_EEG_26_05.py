# -*- coding: utf-8 -*-
from psychopy import visual, core, event, gui
import itertools, csv, time, openpyxl
from random import shuffle
import pandas as pd
import numpy as np


class RatingSkala:
    def __init__(self, win, wort):
        self.skala = visual.ImageStim(win, units="pix",image = 'rating.png', pos = (0,-200))
        self.pfeil = visual.ImageStim(win, units="pix",image = 'pfeil.png', pos = (0,-140))
        self.text = visual.TextStim(win, units="pix", text=wort, height= 60, color=(-1,-1,-1), pos=(0,100))
        self.schriftlinks = visual.TextStim(win, units="pix", text='mag ich garnicht', height= 20, color=(-1,-1,-1), pos=(-200,-250))
        self.schriftrechts = visual.TextStim(win, units="pix", text='mag ich sehr', height= 20, color=(-1,-1,-1), pos=(200,-250))
    def draw(self):
        enter = False
        pfeilpos = 0
        while enter == False:
            self.text.draw()
            self.skala.draw()
            self.schriftlinks.draw()
            self.schriftrechts.draw()
            self.pfeil.pos = (pfeilpos, -140)
            self.pfeil.draw()
            win.flip()
            if event.getKeys(['left']):
                if pfeilpos >= -100:
                    pfeilpos -= 100
                    self.text.draw()
                    self.schriftlinks.draw()
                    self.schriftrechts.draw()
                    self.skala.draw()
                    self.pfeil.pos = (pfeilpos, -140)
                    self.pfeil.draw()
                    win.flip()
            if event.getKeys(['right']):
                if pfeilpos <= 100:
                    pfeilpos += 100
                    self.text.draw()
                    self.schriftlinks.draw()
                    self.schriftrechts.draw()
                    self.skala.draw()
                    self.pfeil.pos = (pfeilpos, -140)
                    self.pfeil.draw()
                    win.flip()
            if event.getKeys(['escape']):
                core.quit()
            if event.getKeys(['return']):
                enter = True
                self.rating = pfeilpos/100
    def getRating(self):
        return self.rating

class ColorPreference:
    def __init__(self, win, color):
        self.skala = visual.ImageStim(win, units="pix",image = 'rating.png', pos = (0,-200))
        self.pfeil = visual.ImageStim(win, units="pix",image = 'pfeil.png', pos = (0,-140))
        self.color = visual.ImageStim(win, units="pix", image = color+'25.png', pos = (0,200))
        self.schriftlinks = visual.TextStim(win, units="pix", text='überhaupt nicht angenehm', height= 20, color=(-1,-1,-1), pos=(-200,-250))
        self.schriftrechts = visual.TextStim(win, units="pix", text='sehr angenehm', height= 20, color=(-1,-1,-1), pos=(200,-250))
    def draw(self):
        enter = False
        pfeilpos = 0
        while enter == False:
            self.color.draw()
            self.skala.draw()
            self.schriftlinks.draw()
            self.schriftrechts.draw()
            self.pfeil.pos = (pfeilpos, -140)
            self.pfeil.draw()
            win.flip()
            if event.getKeys(['left']):
                if pfeilpos >= -100:
                    pfeilpos -= 100
                    self.color.draw()
                    self.schriftlinks.draw()
                    self.schriftrechts.draw()
                    self.skala.draw()
                    self.pfeil.pos = (pfeilpos, -140)
                    self.pfeil.draw()
                    win.flip()
            if event.getKeys(['right']):
                if pfeilpos <= 100:
                    pfeilpos += 100
                    self.color.draw()
                    self.schriftlinks.draw()
                    self.schriftrechts.draw()
                    self.skala.draw()
                    self.pfeil.pos = (pfeilpos, -140)
                    self.pfeil.draw()
                    win.flip()
            if event.getKeys(['escape']):
                core.quit()
            if event.getKeys(['return']):
                enter = True
                self.rating = pfeilpos/100
    def getRating(self):
        return self.rating

class RatingSkala15:
    def __init__(self, win, wort, color):
        self.skala = visual.ImageStim(win, units="pix",image = 'rating15.png',pos = (0,-150))
        self.pfeil = visual.ImageStim(win, units="pix",image = 'pfeilk.png', pos = (0,-100))
        self.text = visual.TextStim(win, units="pix", text=wort, height= 30, color=(-1,-1,-1), pos=(0,0))
        self.schriftlinks = visual.TextStim(win, units="pix", text='passt sehr wenig', height= 10, color=(-1,-1,-1), pos=(-150,-200))
        self.schriftrechts = visual.TextStim(win, units="pix", text='passt sehr gut', height= 10, color=(-1,-1,-1), pos=(150,-200))
        self.color = visual.ImageStim(win, units="pix", image = color+'25.png', pos = (0,200))
    def draw(self):
        enter = False
        pfeilpos = -120
        while enter == False:
            self.color.draw()
            self.text.draw()
            self.skala.draw()
            self.schriftlinks.draw()
            self.schriftrechts.draw()
            self.pfeil.pos = (pfeilpos, -100)
            self.pfeil.draw()
            win.flip()
            if event.getKeys(['left']):
                if pfeilpos >= -60:
                    pfeilpos -= 60
                    self.text.draw()
                    self.schriftlinks.draw()
                    self.schriftrechts.draw()
                    self.color.draw()
                    self.skala.draw()
                    self.pfeil.pos = (pfeilpos, -100)
                    self.pfeil.draw()
                    win.flip()
            if event.getKeys(['right']):
                if pfeilpos <= 60:
                    pfeilpos += 60
                    self.text.draw()
                    self.schriftlinks.draw()
                    self.schriftrechts.draw()
                    self.color.draw()
                    self.skala.draw()
                    self.pfeil.pos = (pfeilpos, -100)
                    self.pfeil.draw()
                    win.flip()
            if event.getKeys(['escape']):
                core.quit()
            if event.getKeys(['return']):
                enter = True
                self.rating = pfeilpos/60
    def getRating(self):
        return self.rating

class ClickColor:
    def __init__(self, win, color):
        self.pic = visual.ImageStim(win, units="pix", image = color+'.png')
        self.frame = visual.Rect(win=win, units="pix", width=150, height=150, name = color)
        self.rahmen = visual.Rect(win=win, units="pix", width=150, height=150, name = color, lineWidth=5, lineColor='black')
    def draw(self, showcolor, pos):
        if showcolor == True:
            self.pic.pos = pos
            self.frame.pos = pos
            self.pic.draw()
            self.frame.draw()
            if mouse.isPressedIn(self.frame):
                return (color)
        elif showcolor == False:
            self.pic.pos = pos
            self.frame.pos = pos
            self.rahmen.pos = pos
            self.pic.draw()
            self.frame.draw()
            self.rahmen.draw()

def showText(messages): 
    
    for message in messages:
        c = None
        while c == None:
            text = visual.TextStim(win, text=message, wrapWidth = 30, color=(-1,-1,-1), pos = (0,0))
            text.draw()
            win.flip()
            c = event.waitKeys()

def showFixation(timetowait):
    fixation = visual.ShapeStim(win, vertices=((0, -0.4), (0, 0.4), (0,0), (-0.4,0), (0.4, 0)), lineWidth=2, closeShape=False,lineColor="black")
    fixation.draw()
    win.flip()
    core.wait(timetowait)

def showWhitescreen(timetowait):
    win.flip()
    core.wait(timetowait)

#Dialogfenster zur Eingabe des VP Codes und ggf. Metadaten
myDlg = gui.Dlg(title="Probanden-Info")
myDlg.addField('VP-Code:')
myDlg.addField('Alter:')
myDlg.addField('Studiengang:', choices =['Psychologie','Anderer','Kein Studierende*r'])
myDlg.addField('Geschlecht:', choices=['weiblich','männlich','non-binär'])
ok_data = myDlg.show()  # Dialog zeigen und auf OK oder Cancel warten
if myDlg.OK:
    gui_data = pd.DataFrame(ok_data).stack().unstack(0)
    print(ok_data)
else:
    print('user cancelled')
    
#####Excel-Datei erstellen
book = openpyxl.Workbook()
book.create_sheet('Data Sheet')
extraSheet = book['Sheet']
book.remove(extraSheet)
book.save(ok_data[0]+'_Data.xlsx')

#######speichere GUI-Daten
book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
sheet = book['Data Sheet']
c = sheet.max_column # Anzahl beschriebener Spalten
head1 = ['VP-Nr','Alter','Studiengang','Geschlecht']
for i in range(len(ok_data)):
    sheet.cell(row=1, column=c+i).value = head1[i]
    sheet.cell(row=2, column=c+i).value = ok_data[i]
book.save(ok_data[0]+'_Data.xlsx')

#Fenster erstellen
win = visual.Window(size=[1200,600], color= (1,1,1), monitor="testMonitor", units="deg") #spaeter evtl. fullscr=True // size=[1200,600]
anfangszeit = time.time()

position = [(-525,90),(-350,90),(-175,90),(0,90),(175,90),(350,90),(525,90),
            (-525,-90),(-350,-90),(-175,-90),(0,-90),(175,-90),(350,-90),(525,-90)]
#position = [(-400,0),(-200,0), (0,0), (200,0),(400,0), (-400,-200)] #für Vorzug/Ablehn
shuffle(position)
farbendict = {'green': (position[0], True), 'red': (position[1], True), 'blue': (position[2], True),'purple': (position[3], True),
                'orange': (position[4], True), 'yellow': (position[5], True), 'cyan': (position[6], True),'darkgreen': (position[7], True), 
                'darkred': (position[8], True), 'darkblue': (position[9], True),'darkpurple': (position[10], True), 
                'darkorange': (position[11], True), 'darkyellow': (position[12], True), 'darkcyan': (position[13], True)}
colorlist = ['green','blue','red','purple','orange','yellow', 'cyan', 
                'darkgreen','darkblue','darkred','darkpurple','darkorange','darkyellow', 'darkcyan']

shuffle(colorlist)

colors = colorlist *1
shuffle(colors)



#######speichere Farbreihenfolge
book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
sheet = book['Data Sheet']
c = sheet.max_column
head2 = []
for x in range(len(colors)):
    name = 'Farbe'+str(x+1)
    head2.append(name)
for i in range(len(colors)):
    sheet.cell(row=1, column=c+1+i).value = head2[i]
    sheet.cell(row=2, column=c+1+i).value = colors[i]
book.save(ok_data[0]+'_Data.xlsx')


#Fenster maximieren und Maus verstecken
win.winHandle.maximize()
win.winHandle.activate()
#win.fullscr = True
win.size = [1366,768]
mouse = event.Mouse(visible = False)
win.flip()

'''

#Startnachrichten
event.clearEvents()
messages = ['Willkommen zum Experiment "Farbe und EEG"!', 'Betrachte zunächst einfach nur die folgenden Farbtafeln auf dem Bildschirm.']
showText(messages)

#parallel.setData(0)
############################### Teil 1: EInzel-Farbtafeln

#Liste zum Speichern der Reihenfolge
#farbreihenfolge = []
for color in colors:
    showFixation(0.2)
    img = visual.ImageStim(win, units="pix",image = color+'5.png', pos = (0,0))
    img.draw()
    win.flip()
    #parallel.setData(Farbzahl)
    ######hier EEG Trigger
    core.wait(0.5) #Triallänge
    #parallel.setData(0)

showFixation(0.5)


####Zweiter Teil: subjektive Farbpräferenzratings

event.clearEvents()
messages = ['Gib nun an, ob du die jeweilige Farbe als eher angenehm oder unangehm empfindest!']
showText(messages)

farbratingdict = {}
event.clearEvents()
enter = False
for color in colorlist:
    showWhitescreen(0.5)
    colorpreference = ColorPreference(win, color)
    colorpreference.draw()
    farbratingdict[color] = colorpreference.getRating()
    win.flip()

######speichere Ratings
book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
sheet = book['Data Sheet']
c = sheet.max_column
ratingfarbe = list(farbratingdict.keys())
for i in range(len(list(farbratingdict.keys()))):
    sheet.cell(row=1, column=c+1+i).value = ratingfarbe[i]+'_rating'
    sheet.cell(row=2, column=c+1+i).value = farbratingdict[ratingfarbe[i]]
book.save(ok_data[0]+'_Data.xlsx')


########################################## Teil 3: Vorzugsfarben

event.clearEvents()
messages = ['Wähle nun drei Farben aus, die dir am angenehmsten sind!']
showText(messages)
mouse = event.Mouse(visible = True)
win.flip()
#Ergebnisliste
vorzug = []
ablehn = []
x = False
fertig = 0
event.clearEvents()
while x == False:
    for color, value in farbendict.items():
        c = ClickColor(win, color)
        clickedcolor = c.draw(value[1],value[0])
        if clickedcolor != None and len(vorzug) < 3:
            vorzug.append(clickedcolor)
            farbendict[color] = (value[0], False)
        if len(vorzug) == 3:
            fertig += 1
    win.flip()
    core.wait(0.01) ###### bei Fehler evtl Zeit hoch setzen zB 0.1
    if fertig > 300: ##### Endzeit ist ungefähr Wert/100 Sekunden
        x = True

#print(vorzug)
#print(res)
#print(farbendict)

win.flip()
core.wait(2)
######speichere Vorzugsfarben
book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
sheet = book['Data Sheet']
c = sheet.max_column
head3 = []
for x in range(len(vorzug)):
    name = 'Vorzugsfarbe'+str(x+1)
    head3.append(name)
for i in range(len(vorzug)):
    sheet.cell(row=1, column=c+1+i).value = head3[i]
    sheet.cell(row=2, column=c+1+i).value = vorzug[i]
book.save(ok_data[0]+'_Data.xlsx')


event.clearEvents()
messages = ['Wähle nun zwei Farben aus, die dir am Wenigsten gefallen!']
showText(messages)


# Teil 3: Ablehnfarbe


shuffle(position)
farbendict = {'green': (position[0], True), 'red': (position[1], True), 'blue': (position[2], True),'purple': (position[3], True),
                'orange': (position[4], True), 'yellow': (position[5], True), 'cyan': (position[6], True),'darkgreen': (position[7], True), 
                'darkred': (position[8], True), 'darkblue': (position[9], True),'darkpurple': (position[10], True), 
                'darkorange': (position[11], True), 'darkyellow': (position[12], True), 'darkcyan': (position[13], True)}


x = False
fertig = 0
event.clearEvents()
while x == False:
    for color, value in farbendict.items():
        c = ClickColor(win, color)
        clickedcolor = c.draw(value[1],value[0])
        if clickedcolor != None and len(ablehn) < 3:
            ablehn.append(clickedcolor)
            farbendict[color] = (value[0], False)
        if len(ablehn) == 3:
            fertig += 1
    win.flip()
    core.wait(0.01) ###### bei Fehler evtl Zeit hoch setzen zB 0.1
    if fertig > 300: ##### Endzeit ist ungefähr Wert/100 Sekunden
        x = True


#print(ablehn)
#print(res)
#print(farbendict)

######speichere Ablehnfarben
book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
sheet = book['Data Sheet']
c = sheet.max_column
head4 = []
for x in range(len(vorzug)):
    name = 'Ablehnfarbe'+str(x+1)
    head4.append(name)
for i in range(len(ablehn)):
    sheet.cell(row=1, column=c+1+i).value = head4[i]
    sheet.cell(row=2, column=c+1+i).value = ablehn[i]
book.save(ok_data[0]+'_Data.xlsx')

win.flip()
core.wait(2)

'''
############# Farbe einzeln mit Freitext-Wort-Assoziationen

event.clearEvents()
messages = ['Gib nun jeweils mindestens zwei Wörter an, die du spontan mit der präsentierten Farbe assoziierst (trenne die Wörter durch ein Komma). \n Beurteile anschließend, wie gut diese Assoziation zur Farbe passt!']
showText(messages)

assoziationsdict = {}
mouse = event.Mouse(visible = True)
übungslist = ['grey']

#Übungsdurchgang:
for color in übungslist:
    event.clearEvents()
    showWhitescreen(0.5)
    pic = visual.ImageStim(win, units="pix", image = color+'25.png')
    instruction = visual.TextStim(win ,color="black")
    quitKeys = ['escape', 'esc']
    ansKeys = [ 'delete']
    keyboardKeys = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','ä', 'ö', 'ü', 'ß']
    keydict = {'comma':','}
    answer = ''
    x = True
    while x == True:
        instruction.setText(u'{0}'.format(answer))
        instruction.draw()
        pic.pos = (0,200)
        pic.draw()
        #button = visual.Rect(win, units="pix", width=50, height=50, pos = (450,0), name = 'button')
        #button.draw()
        #buttonpic = visual.ImageStim(win, units="pix", image = 'button50.jpg', pos = (450,0))
        #buttonpic.draw()
        win.flip()
        # get some keys.
        for letter in (keyboardKeys):
            if event.getKeys([letter]):
                answer += letter

        if event.getKeys(['backspace']):
            answer = answer[:-1]

        if event.getKeys(['space']):
            answer += ' '

        for letter in (keydict.keys()):
                if event.getKeys([letter]):
                    answer += keydict[letter]

        if event.getKeys([quitKeys[0]]):
            core.quit()

        if event.getKeys([ansKeys[0]]): #####Eingabetaste wählen
            break
        #if mouse.isPressedIn(button):
        #    break
    assoziationsdict[color] = answer
    wörter = []
    wörter += answer.strip('\'').strip().split(',')
    matchratingdict = {}
    enter = False
    for wort in wörter:
        event.clearEvents()
        rate = RatingSkala15(win, wort, color)
        rate.draw()
        matchratingdict[wort] = rate.getRating()
    #print(matchratingdict)
    win.flip()
    ######speichere matching-ratings
    book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
    sheet = book['Data Sheet']
    c = sheet.max_column
    assiwörter = list(matchratingdict.keys())
    for i in range(len(list(matchratingdict.keys()))):
        sheet.cell(row=1, column=c+1+i).value = assiwörter[i]+'_'+color+'_matching'
        sheet.cell(row=2, column=c+1+i).value = matchratingdict[assiwörter[i]]
    book.save(ok_data[0]+'_Data.xlsx')


event.clearEvents()
messages = ['Dies war ein Testdurchgang, nun geht es los!']
showText(messages)


assoziationsdict = {}
for color in colorlist:
    event.clearEvents()
    showWhitescreen(0.5)
    pic = visual.ImageStim(win, units="pix", image = color+'25.png')
    instruction = visual.TextStim(win ,color="black")
    quitKeys = ['escape', 'esc']
    ansKeys = [ 'delete']
    keyboardKeys = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','ä', 'ö', 'ü', 'ß']
    keydict = {'comma':','}
    answer = ''
    x = True
    while x == True:
        instruction.setText(u'{0}'.format(answer))
        instruction.draw()
        pic.pos = (0,200)
        pic.draw()
        button = visual.Rect(win, units="pix", width=50, height=50, pos = (450,0), name = 'button')
        button.draw()
        buttonpic = visual.ImageStim(win, units="pix", image = 'button50.jpg', pos = (450,0))
        buttonpic.draw()
        win.flip()
        # get some keys.
        for letter in (keyboardKeys):
            if event.getKeys([letter]):
                answer += letter

        if event.getKeys(['backspace']):
            answer = answer[:-1]

        if event.getKeys(['space']):
            answer += ' '

        for letter in (keydict.keys()):
                if event.getKeys([letter]):
                    answer += keydict[letter]

        if event.getKeys([quitKeys[0]]):
            core.quit()

        if event.getKeys([ansKeys[0]]): #####Eingabetaste wählen
            break
        if mouse.isPressedIn(button):
            break
    assoziationsdict[color] = answer
    wörter = []
    wörter += answer.strip('\'').strip().split(',')
    matchratingdict = {}
    enter = False
    for wort in wörter:
        event.clearEvents()
        rate = RatingSkala15(win, wort, color)
        rate.draw()
        matchratingdict[wort] = rate.getRating()
    #print(matchratingdict)
    win.flip()
    ######speichere matching-ratings
    book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
    sheet = book['Data Sheet']
    c = sheet.max_column
    assiwörter = list(matchratingdict.keys())
    for i in range(len(list(matchratingdict.keys()))):
        sheet.cell(row=1, column=c+1+i).value = assiwörter[i]+'_'+color+'_matching'
        sheet.cell(row=2, column=c+1+i).value = matchratingdict[assiwörter[i]]
    book.save(ok_data[0]+'_Data.xlsx')


#Schleife zum Testen, wie die Tasten heissen
'''for i in range(4):
    keys = event.getKeys()
    print(keys)
    core.wait(1)'''

######speichere Assoziationen
book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
sheet = book['Data Sheet']
c = sheet.max_column
assifarben = list(assoziationsdict.keys())
for i in range(len(list(assoziationsdict.keys()))):
    sheet.cell(row=1, column=c+1+i).value = assifarben[i]+'_assoziation'
    sheet.cell(row=2, column=c+1+i).value = assoziationsdict[assifarben[i]]
book.save(ok_data[0]+'_Data.xlsx')

####Letzter Teil: Wortrating

event.clearEvents()
messages = ['Beurteile nun, wie sehr du die jeweilige Assoziation magst!']
showText(messages)

wortliste = []

for i in list(assoziationsdict.values()):
    wortliste += i.strip('\'').strip().split(',')

shuffle(wortliste)
ratingdict = {}

event.clearEvents()

enter = False
for wort in wortliste:
    showWhitescreen(0.5)
    rate = RatingSkala(win, wort)
    rate.draw()
    ratingdict[wort] = rate.getRating()
    win.flip()

######speichere Ratings
book = openpyxl.load_workbook(ok_data[0]+'_Data.xlsx')
sheet = book['Data Sheet']
c = sheet.max_column
ratingwort = list(ratingdict.keys())
for i in range(len(list(ratingdict.keys()))):
    sheet.cell(row=1, column=c+1+i).value = ratingwort[i]+'_rating'
    sheet.cell(row=2, column=c+1+i).value = ratingdict[ratingwort[i]]
book.save(ok_data[0]+'_Data.xlsx')

#Warteschleifenloop, bis Ende auf Tastendruck
#c =None
#while c==None:
showText(["Beenden mit beliebiger Taste"])
#c = event.waitKeys()
endzeit = time.time()
dauer = endzeit - anfangszeit
print('Versuchsdauer: {}'.format(dauer))

# Fenster und PsychoPy schliessen
win.close()
core.quit()