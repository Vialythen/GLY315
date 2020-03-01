from xlrd import open_workbook
import matplotlib.pyplot as plt
import numpy


def simpson_index(species):
    nn = 0
    M = 0
    for numOfSpec in species:
        nn = nn + int(numOfSpec)*(int(numOfSpec)-1)
        M = M + int(numOfSpec)
    return 1/(nn/(M*(M-1)))



wb = open_workbook('raw data foram GLY 315 EXIII.xlsx')
for sheet in wb.sheets():

    zones = []

    for zone in range(6, 15):
        specimenInZone = []
        for species in range(4, 57):
            value = 0.0 if sheet.cell(species,zone).value == '' or sheet.cell(species,zone).value == ' ' else sheet.cell(species,zone).value
            value = str(round(value)) if isinstance(value, float) else value
            specimenInZone.append(value)
        zones.append(specimenInZone)

    diversities = []

    for zone in zones:
        print('{}: {}'.format(zone[0], str(simpson_index(zone[1:]))))
        diversities.append(simpson_index(zone[1:]))

    ages = [72.4, 74.5, 74.5, 74.5, 74.5, 74.5, 76.1, 76.1, 76.1]

    coef = numpy.polyfit(ages, diversities, 1)

    plt.plot(ages, diversities, 'ro', [60, 80], [coef[0]*60+coef[1], coef[0]*80+coef[1]])
    plt.axis([72,77,0,15])

    plt.title('Diversity Vs Age')
    plt.xlabel('Age (Ma)')
    plt.ylabel('Diversity\n(Recirocal Simpson\'s Index)')
    plt.text(72.5, 11, 'y = {}x + {}'.format(round(coef[0],4), round(coef[1],4)))



    plt.figure(2)

    feedingmode = []
    fmindex = []

    for species in range(5, 57):
        feedmode = sheet.cell(species, 4).value
        feedingmode.append(feedmode)
        if feedmode == 'herbivore':
            fmindex.append(1)
        elif feedmode == 'omnivore':
            fmindex.append(2)
        elif feedmode == 'detrivore' or feedmode == 'detritivore' or feedmode == 'detritovore':
            fmindex.append(3)
        elif feedmode == 'deposit feeder':
            fmindex.append(4)
        else:
            print('error reading {}'.format(feedmode))

    size = [155,190,155,155,205,210,175,133,155,165,205,220,300,213,155,275,325,348,310,305,250,285,213,215,160,160,250,280,175,190,190,190,180,190,260,180,235,210,325,300,175,165,190,305,190,165,165,155,203,198,190,203]
    
    xticks = ['Herbivore', 'Omnivore', 'Detritivore', 'Deposit Feeder']
    plt.xticks([1,2,3,4], xticks)
    plt.plot(fmindex, size, 'ro')        

    plt.title('Size Vs Feeding Mode')
    plt.xlabel('Feeding Mode')
    plt.ylabel('Size (microns)')

    plt.show()







