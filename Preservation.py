from xlrd import open_workbook
import matplotlib.pyplot as plt
import numpy


wb = open_workbook('FinalQ3Data.xlsx')
for sheet in wb.sheets():

    data = []

    for row in range(1, 21):
        sampledata = []
        for col in range(3, 9):
            value = float(sheet.cell(row,col).value)
            sampledata.append(value)
        data.append(sampledata)

    centroid = [x for x in data if x[5] < -20]

    plt.plot([x[5] for x in data if x[0] == 1], [x[1] for x in data if x[0] == 1], 'go', label='Integrity: 1')
    plt.plot([x[5] for x in data if x[0] == .6], [x[1] for x in data if x[0] == .6], 'yo', label='Integrity: .6')
    plt.plot([x[5] for x in data if x[0] == .1], [x[1] for x in data if x[0] == .1], 'ro', label='Integrity: .1')
    plt.plot([x[5] for x in data if x[0] == 0], [x[1] for x in data if x[0] == 0], 'ko', label='Integrity: 0')
    plt.plot(sum(x[5]*x[0] for x in centroid)/sum(x[0] for x in centroid), sum(x[1]*x[0] for x in centroid)/sum(x[0] for x in centroid), 'kx', label='Weighted Centroid')

    plt.title('Iron Vs Carbon-13')
    plt.xlabel('Carbon-13')
    plt.ylabel('Iron Count (%wt)')
    plt.legend()

    plt.figure(2)

    plt.plot([x[5] for x in data if x[0] == 1], [x[2] for x in data if x[0] == 1], 'go', label='Integrity: 1')
    plt.plot([x[5] for x in data if x[0] == .6], [x[2] for x in data if x[0] == .6], 'yo', label='Integrity: .6')
    plt.plot([x[5] for x in data if x[0] == .1], [x[2] for x in data if x[0] == .1], 'ro', label='Integrity: .1')
    plt.plot([x[5] for x in data if x[0] == 0], [x[2] for x in data if x[0] == 0], 'ko', label='Integrity: 0')
    plt.plot(sum(x[5]*x[0] for x in centroid)/sum(x[0] for x in centroid), sum(x[2]*x[0] for x in centroid)/sum(x[0] for x in centroid), 'kx', label='Weighted Centroid')

    plt.title('Manganese Vs Carbon-13')
    plt.xlabel('Carbon-13')
    plt.ylabel('Manganese Count (ppm)')
    plt.legend()

    plt.figure(3)

    plt.plot([x[5] for x in data if x[0] == 1], [x[3] for x in data if x[0] == 1], 'go', label='Integrity: 1')
    plt.plot([x[5] for x in data if x[0] == .6], [x[3] for x in data if x[0] == .6], 'yo', label='Integrity: .6')
    plt.plot([x[5] for x in data if x[0] == .1], [x[3] for x in data if x[0] == .1], 'ro', label='Integrity: .1')
    plt.plot([x[5] for x in data if x[0] == 0], [x[3] for x in data if x[0] == 0], 'ko', label='Integrity: 0')
    plt.plot(sum(x[5]*x[0] for x in centroid)/sum(x[0] for x in centroid), sum(x[3]*x[0] for x in centroid)/sum(x[0] for x in centroid), 'kx', label='Weighted Centroid')

    plt.title('Sulfur Vs Carbon-13')
    plt.xlabel('Carbon-13')
    plt.ylabel('Sulfur Count (ppm)')
    plt.legend()

    plt.figure(4)

    plt.plot([x[5] for x in data if x[0] == 1], [x[4] for x in data if x[0] == 1], 'go', label='Integrity: 1')
    plt.plot([x[5] for x in data if x[0] == .6], [x[4] for x in data if x[0] == .6], 'yo', label='Integrity: .6')
    plt.plot([x[5] for x in data if x[0] == .1], [x[4] for x in data if x[0] == .1], 'ro', label='Integrity: .1')
    plt.plot([x[5] for x in data if x[0] == 0], [x[4] for x in data if x[0] == 0], 'ko', label='Integrity: 0')
    plt.plot(sum(x[5]*x[0] for x in centroid)/sum(x[0] for x in centroid), sum(x[4]*x[0] for x in centroid)/sum(x[0] for x in centroid), 'kx', label='Weighted Centroid')

    plt.title('Oxygen Vs Carbon-13')
    plt.xlabel('Carbon-13')
    plt.ylabel('Oxygen Count (ppm)')
    plt.legend()

    plt.show()







