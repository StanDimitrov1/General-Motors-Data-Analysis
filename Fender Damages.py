import pandas as pd
import math

import pikepdf
import PyPDF2

from reportlab.lib.colors import HexColor

from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

wavy = [0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]     # arrays used to store wavy
dents = [0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]    # arrays used to store dents
bents = [0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]    # arrays used to store bents
corners = [0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]  # arrays used to store corners
edges = [0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]    # arrays used to store edges
damages = [0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]  # arrays used to store damages


def clear():  # function to clear value of arrays
    for x in range(0, 3):

        for y in range(0, 4):
            dents[x][y] = 0
            bents[x][y] = 0
            corners[x][y] = 0
            edges[x][y] = 0
            damages[x][y] = 0
            wavy[x][y] = 0


def zero(a):
    if a == "0":
        return ""
    else:
        return a


def write(part):
    global wavy
    global dents
    global bents
    global corners
    global edges
    global damages

    x1 = [70, 127, 200, 260]
    x2 = [343, 400, 473, 533]

    y1 = [620, 560, 500]
    y2 = [390, 330, 270]
    y3 = [170, 110, 50]

    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    c.setFillColor(HexColor('#FF0000'))
    c.setFontSize(20)

    for a in range(0, 3):

        for b in range(0, 4):
            d0 = dents[a][b]
            d = str(d0)
            d2 = zero(d)

            c.drawString(x1[b], y1[a], d2)

    for e in range(0, 3):

        for f in range(0, 4):
            b0 = bents[e][f]
            b = str(b0)
            b2 = zero(b)
            c.drawString(x1[f], y2[e], b2)

    for g in range(0, 3):

        for h in range(0, 4):
            c0 = corners[g][h]
            c1 = str(c0)
            c2 = zero(c1)
            c.drawString(x2[h], y2[g], c2)

    # COUNTING EDGES
    for i in range(0, 3):

        for j in range(0, 4):
            e0 = edges[i][j]
            e = str(e0)
            e2 = zero(e)
            c.drawString(x2[j], y1[i], e2)

    for k in range(0, 3):

        for l in range(0, 4):
            w0 = wavy[k][l]
            w = str(w0)
            w2 = zero(w)
            c.drawString(x2[l], y3[k], w2)

    for m in range(0, 3):

        for n in range(0, 4):
            d0 = wavy[m][n]
            d1 = str(d0)
            d2 = zero(d1)
            c.drawString(x1[n], y3[m], d2)

    # create a new PDF with Reportlab

    c.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)

    # read your existing PDF
    existing_pdf = PdfFileReader(open(part + ".pdf", "rb"))
    output = PdfFileWriter()

    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page2 = new_pdf.getPage(0)
    page.mergePage(page2)
    output.addPage(page)

    # finally, write "output" to a real file
    outputStream = open(part + " PLOT.pdf", "wb")
    output.write(outputStream)
    outputStream.close()


df = pd.read_excel('Fender Damages.xlsx')

r = 0
c = 1


def DENT(a, b, c):  # storing dent damages
    global dents

    int(a)
    int(b)
    int(c)

    dents[a][b] += c
    print(dents[a][b])


def BENT(a, b, c):  # storing bent damages
    global bents

    int(a)
    int(b)
    int(c)

    bents[a][b] += c
    print(bents[a][b])


def CORNER(a, b, c): # storing corner damages
    global corners

    int(a)
    int(b)
    int(c)

    corners[a][b] += c
    print(corners[a][b])


def EDGE(a, b, c): # stores edge damages
    global edges

    int(a)
    int(b)
    int(c)

    edges[a][b] += c
    print(edges[a][b])


def WAVY(a, b, c):  # storing wavy damages
    global wavy

    int(a)
    int(b)
    int(c)

    wavy[a][b] += c
    print(wavy[a][b])


def DAMAGED(a, b, c):  # storing extra damages
    global damages

    int(a)
    int(b)
    int(c)

    damages[a][b] += c
    print(damages[a][b])


# following if statements in function parse excel lines looking for the key words

def checker(damage):
    if noteU.find('1U') > 0:
        print(damage + " 1U")
        eval(damage + "(0,0,1)")

    if noteU.find('1M') > 0:
        print(damage + " 1M")
        eval(damage + "(1,0,1)")

    if noteU.find('1L') > 0:
        print(damage + " 1L")
        eval(damage + "(2,0,1)")

    if noteU.find('2U') > 0:
        print(damage + " 2U")
        eval(damage + "(0,1,1)")

    if noteU.find('2M') > 0:
        print(damage + " 2M")
        eval(damage + "(1,1,1)")

    if noteU.find('2L') > 0:
        print(damage + " 2L")
        eval(damage + "(2,1,1)")

    if noteU.find('3U') > 0:
        print(damage + " 3U")
        eval(damage + "(0,2,1)")

    if noteU.find('3M') > 0:
        print(damage + " 3M")
        eval(damage + "(1,2,1)")

    if noteU.find('3L') > 0:
        print(damage + " 3L")
        eval(damage + "(2,2,1)")

    if noteU.find('4U') > 0:
        print(damage + " 4U")
        eval(damage + "(0,3,1)")

    if noteU.find('4M') > 0:
        print(damage + " 4M")
        eval(damage + "(1,3,1)")

    if noteU.find('4L') > 0:
        print(damage + " 4L")
        eval(damage + "(2,3,1)")

# checking secondary condition (the key word showing the damage)
def checker2():
    if noteU.find('DENT' or 'DNT') > 0:
        checker('DENT')

    if noteU.find('BENT' or 'BNT') > 0:
        checker('BENT')

    # if noteU.find('DAMAGED' or 'DMG') > 0:
    # checker('DAMAGED')

    if noteU.find('CORNER') > 0:
        checker('CORNER')

    if noteU.find('ROLLED' and 'EDGE') > 0:
        checker('EDGE')

    if noteU.find('WAVY') > 0:
        checker('WAVY')

#matching part number to commodity
for c in range(2, 200):
    item = df.iloc[r, 0]
    r += 1
    # c+=1

    if item == 23303550:
        LGMC = c
        print(LGMC)

    if item == '23303551':
        LChevy = c
        print(LChevy)

    if item == '84214215':
        RChevy = c
        print(RChevy)

    if item == '84214216':
        RGMC = c
        print(RGMC)

count = 1
for b in range(1, 211):
    damage = df.iloc[b, 1]

    if damage == 'Transportation Damage':
        if count == 1:
            limit1 = b + 2
        if count == 2:
            limit2 = b + 2
        if count == 3:
            limit3 = b + 2
        count += 1
        print(count)

print("LH GMC!!!!")
for c in range(1, limit1 - 2):
    note = df.iloc[c, 2]
    noteU = note.upper()
    print(noteU)
    checker2()
    write('LH-GMC')

clear()

print('LH CHEVY!!!')
a = 0
for c2 in range(LChevy - 2, limit2 - 2):
    note = df.iloc[c2, 2]
    noteU = note.upper()
    a += 1
    print(str(a) + " " + noteU)

    checker2()
    write('LH-CHEVY')

clear()

print("RH-CHEVY!!!!")

b = 0
for c3 in range(RChevy - 2, limit3 - 2):
    note = df.iloc[c3, 2]
    noteU = note.upper()
    b += 1
    print(str(b) + " " + noteU)

    checker2()
    write('RH-CHEVY')

clear()

print("RH-GMC!!!!!!")

c = 0
for c3 in range(RGMC - 2, 211 - 2):
    note = df.iloc[c3, 2]
    noteU = note.upper()
    c += 1
    print(str(c) + " " + noteU)

    checker2()
    write('RH-GMC')

print(LGMC, LChevy, RChevy, RGMC)
print(limit1, limit2, limit3)

# List=[1,2,3],[4,5,6]
#   1,2,3
#   4,5,6

# printing results
print("DENTS")
print("1U:  " + str(dents[0][0]))
print("2U:  " + str(dents[0][1]))
print("3U:  " + str(dents[0][2]))
print("4U:  " + str(dents[0][3]))

print("1M:  " + str(dents[1][0]))
print("2M:  " + str(dents[1][1]))
print("3M:  " + str(dents[1][2]))
print("4M:  " + str(dents[1][3]))

print("1L:  " + str(dents[2][0]))
print("2L:  " + str(dents[2][1]))
print("3L:  " + str(dents[2][2]))
print("4L:  " + str(dents[2][3]))

print("BENTS")
print("1U:  " + str(bents[0][0]))
print("2U:  " + str(bents[0][1]))
print("3U:  " + str(bents[0][2]))
print("4U:  " + str(bents[0][3]))

print("1M:  " + str(bents[1][0]))
print("2M:  " + str(bents[1][1]))
print("3M:  " + str(bents[1][2]))
print("4M:  " + str(bents[1][3]))

print("1L:  " + str(bents[2][0]))
print("2L:  " + str(bents[2][1]))
print("3L:  " + str(bents[2][2]))
print("4L:  " + str(bents[2][3]))

print("CORNERS")
print("1U:  " + str(corners[0][0]))
print("2U:  " + str(corners[0][1]))
print("3U:  " + str(corners[0][2]))
print("4U:  " + str(corners[0][3]))

print("1M:  " + str(corners[1][0]))
print("2M:  " + str(corners[1][1]))
print("3M:  " + str(corners[1][2]))
print("4M:  " + str(corners[1][3]))

print("1L:  " + str(corners[2][0]))
print("2L:  " + str(corners[2][1]))
print("3L:  " + str(corners[2][2]))
print("4L:  " + str(corners[2][3]))

print("WAVY")
print("1U:  " + str(corners[0][0]))
print("2U:  " + str(corners[0][1]))
print("3U:  " + str(corners[0][2]))
print("4U:  " + str(corners[0][3]))

print("1M:  " + str(corners[1][0]))
print("2M:  " + str(corners[1][1]))
print("3M:  " + str(corners[1][2]))
print("4M:  " + str(corners[1][3]))

print("1L:  " + str(corners[2][0]))
print("2L:  " + str(corners[2][1]))
print("3L:  " + str(corners[2][2]))
print("4L:  " + str(corners[2][3]))
