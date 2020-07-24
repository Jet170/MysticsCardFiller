import csv
#Creates deck from CSV file
def deckFromCSV(csvFileName):
    with open(csvFileName, newline='') as csvfile:
        spamreader = csv.reader(csvfile, dialect='excel')
        carddata = []

        print("Reading CSV Data...")
        for row in spamreader:
            singleformdata = {
                'Copies': str(row[0]),
                'Name': str(row[1])
            }
            carddata.append(singleformdata)
    deckFile = open(csvFileName.replace('.csv', '') + '.txt', 'w')
    for card in carddata:
        deckFile.write(card['Copies'] + ' ' + card['Name'] + '\n')

csvname = input("Enter the csv name ")
deckFromCSV(csvname)
print("Deck saved as " + csvname.replace('.csv', ''))