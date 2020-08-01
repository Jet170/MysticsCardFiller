from __future__ import print_function

from mailmerge import MailMerge
from docx2pdf import convert
from PIL import Image, ImageFont, ImageDraw
import csv
import textwrap
import datetime
import xml.etree.ElementTree as ET


#File names, set names, urls, all that stuff that could change
docname = "CardsForm.docx"
numCardsPerForm=8
sheetname="Cardlist.csv"
outputDocName='filledcards.docx'
outputPDFName="filledcards.pdf"
blankCardImage="Blank.png"
cardImageFolder='./CardImages/'
xmlFile = 'MysticsSet1.xml'
picURL = 'https://raw.githubusercontent.com/Jet170/MysticsCardFiller/master/CardImages'
setShortName = 'MS1'
setLongName = 'Mystics Set 1'

#Converts a csv file of cards to a word document with the cards filled out
def CsvToWord():
    print("Starting Card Filler...")

    #Read cards from csv
    with open(sheetname, newline='') as csvfile:
        spamreader = csv.reader(csvfile, dialect='excel')
        singleformdata = []
        carddata = []
        i = 1
        print("Reading CSV Data...")
        for row in spamreader:
            #For the first card, reset the form data storage object
            if i==1:
                singleformdata = {
                    'Card' + str(i) + 'Name':str(row[0]),
                    'Card' + str(i) + 'Mana':str(row[1]),
                    'Card' + str(i) + 'Type':str(row[2]),
                    'Card' + str(i) + 'MagicType':str(row[3]),
                    'Card' + str(i) + 'Text':str(row[4]),
                }
            #For each subsequent card, add the data to the form data
            else:
                singleformdata['Card' + str(i) + 'Name'] = str(row[0])
                singleformdata['Card' + str(i) + 'Mana'] = str(row[1])
                singleformdata['Card' + str(i) + 'Type'] = str(row[2])
                singleformdata['Card' + str(i) + 'MagicType'] = str(row[3])
                singleformdata['Card' + str(i) + 'Text'] = str(row[4])
            i+=1
            if i>numCardsPerForm:
                i=1
                carddata.append(singleformdata)

    #Fill the final sheet with blank cards if there the number of cards isn't divisible by four
    if i!=1:
        while i<=numCardsPerForm:
            singleformdata['Card' + str(i) + 'Name'] = ''
            singleformdata['Card' + str(i) + 'Mana'] = ''
            singleformdata['Card' + str(i) + 'Type'] = ''
            singleformdata['Card' + str(i) + 'MagicType'] = ''
            singleformdata['Card' + str(i) + 'Text'] = ''
            i+=1
        carddata.append(singleformdata)

    #Create new document
    document = MailMerge(docname)
    print("Creating Document...")
    document.merge_templates(carddata, 'page_break')
    document.write(outputDocName)
    print("Cards Sucessfully Exported to " + outputDocName)

#Converts the word doc to a pdf
def wordToPdf(requireDocPause):
    print("Converting from Word doc to pdf...")
    if(requireDocPause):
        input("Please open " + outputDocName + ", save over the original output file, then close it and press enter")
    convert(outputDocName, outputPDFName)

#Reads cards from the specified csv file
def readFromCSV():
    with open(sheetname, newline='') as csvfile:
        spamreader = csv.reader(csvfile, dialect='excel')
        carddata = []

        print("Reading CSV Data...")
        for row in spamreader:
            singleformdata = {
                'Name': str(row[0]),
                'Mana': str(row[1]),
                'Type': str(row[2]) + '-',
                'MagicType': str(row[3]),
                'Text': str(row[4])
            }
            carddata.append(singleformdata)
    return carddata


#Creates card images
def createCardImages():
    nameFont = ImageFont.truetype("constan.ttf", 20)
    nameCoords = [5, 5]
    manaFont = ImageFont.truetype("constani.ttf", 20)
    manaCoords = [226, 12]
    typeFont = ImageFont.truetype("constan.ttf", 15)
    typeCoords = [5, 30]
    magicFont = ImageFont.truetype("constani.ttf", 15)
    #Magic coords depend on card type
    textFont = ImageFont.truetype("constan.ttf", 16)
    textCoords = [5, 205]

    # Read cards from csv
    carddata = readFromCSV()

    print("Appending to cards...")
    for card in carddata:
       image = Image.open(blankCardImage)
       draw = ImageDraw.Draw(image)
       # Draw Name
       draw.text(nameCoords, card['Name'], (0, 0, 0), nameFont)
       twidth, theight = draw.textsize(card['Name'], nameFont)
       draw.line((nameCoords[0], nameCoords[1] + theight, nameCoords[0] + twidth, nameCoords[1] + theight), 'black')

       # Draw Mana
       draw.text(manaCoords, card['Mana'], (0, 0, 0), manaFont)

       # Draw Type
       draw.text(typeCoords, card['Type'], (0, 0, 0), typeFont)
       # Draw Magic Type
       draw.text((typeCoords[0] + draw.textsize(card['Type'], typeFont)[0], typeCoords[1]), card['MagicType'],
                 (0, 0, 0), magicFont)
       # Draw Card Text
       draw.multiline_text(textCoords, "\n".join(textwrap.wrap(card['Text'], width=33)), (0, 0, 0), textFont)

       image.save(cardImageFolder + card['Name'] + '.png')

    print('Images saved to ' + cardImageFolder)

#Create card XML File for cockatrice
def createCockatriceXML():
    print('Creating XML Objects...')
    currentDate = datetime.datetime.now()
    carddata = readFromCSV()
    topLevel = ET.Element('cockatrice_carddatabase')
    topLevel.set('version', '4')
    sets = ET.SubElement(topLevel, 'sets')
    set = ET.SubElement(sets, 'set')
    name = ET.SubElement(set, 'name')
    name.text = setShortName
    longname = ET.SubElement(set, 'longname')
    longname.text=setLongName
    settype = ET.SubElement(set, 'settype')
    settype.text = 'Custom'
    releasedate = ET.SubElement(set, 'releasedate')
    releasedate.text = currentDate.strftime('%Y-%m-%d')
    cards = ET.SubElement(topLevel, 'cards')
    print('Adding card data...')
    for dataCard in carddata:
        xmlCard = ET.SubElement(cards, 'card')
        name = ET.SubElement(xmlCard, 'name')
        name.text = dataCard['Name']
        text = ET.SubElement(xmlCard, 'text')
        text.text = 'Mana ' + dataCard['Mana'] + ' Type ' + dataCard['Type'] + ' Magic Type ' + dataCard['MagicType'] + ' Description ' + dataCard['Text']
        prop = ET.SubElement(xmlCard, 'prop')
        cardType = ET.SubElement(prop, 'type')
        cardType.text = dataCard['Type']
        manacost = ET.SubElement(prop, 'manacost')
        manacost.text = dataCard['Mana']
        cmc = ET.SubElement(prop, 'cmc')
        cmc.text = dataCard['Mana']
        colors = ET.SubElement(prop, 'colors')
        colors.text = dataCard['MagicType']
        set = ET.SubElement(xmlCard, 'set')
        set.set('picurl', picURL + '/' + dataCard['Name'].replace(' ', '%20') + '.png')
        set.text=setShortName
        tablerow = ET.SubElement(xmlCard, 'tablerow')
        tablerow.text='1'
        cipt=ET.SubElement(xmlCard, 'cipt')
        cipt.text='1'
    print('Saving file..')
    outputFile = open(xmlFile, 'w')
    outputFile.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    outputFile.write(ET.tostring(topLevel, encoding='unicode'))
    print('XML file created successfuly')

printable = input("Create printable? ")
if printable.upper() == 'YES' or printable.upper() == 'Y':
    CsvToWord()
pdf = input("Create pdf? ")
if(pdf.upper() == 'YES' or pdf.upper() == 'Y'):
    wordToPdf(True)
cockatrice = input("Update cockatrice? ")
if(cockatrice.upper() == 'YES' or cockatrice.upper() == 'Y'):
    createCardImages()
    createCockatriceXML()