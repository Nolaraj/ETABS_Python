# Co Ordinates Modelling
# Level1 = Floor wise Co ordinates, Level 2 = Along Horizontal bays co, ordinates  [[], [], []]
# Similarly same for y co_ordinates

import numpy as np
baysX = 4
baysY = 2
import math


Angle = 30

Angle_AddGrids = {}
X_BaysSpacing = [5,6,7,8,9]
Y_BaysSpacing = [3,4,5,6,7]
Origin_X = 0
Origin_Y = 0

Inclined_Bay = 3 #ws.cell(column=4, row=gdRow + 5).value
# for i in range(4, 13):
#     X_BaysSpacing.append(ws.cell(column=i, row=gdRow + 3).value)
#     Y_BaysSpacing.append(ws.cell(column=i, row=gdRow + 4).value)
# for row in range(gdRow + 1, gdRow + 7):
#     Angle_AddGrids[int(ws.cell(column=14, row=row).value)] = int(ws.cell(column=14 + 1, row=row).value)
X_Previous = Origin_X
Y_Previous = Origin_Y


blocks = []
block1 = []
block2 = []
block3 = []

for h in range(0, baysX + 1):
    Xs = []
    Ys = []

    # Phase 1 - Before bent
    if h < Inclined_Bay :
        for i in range(0, baysY + 1):
            if h == 0:
                Xs.append(X_Previous)
            else:
                Xs.append(X_Previous + X_BaysSpacing[h - 1])

            if i == 0:
                Ys.append(Y_Previous)
            else:
                Ys.append(Ys[i - 1] + Y_BaysSpacing[i - 1])
        line = []
        for i in range(0, len(Xs)):
            xy = [Xs[i], Ys[i]]
            line.append(xy)

        block1.append(line)
        X_Previous = Xs[0]

    if h == Inclined_Bay:
        blocks.append(block1)

    # Phase 2 - Bent
    if h == Inclined_Bay :
        pass



    # Phase 3 - After bent

    if h > Inclined_Bay:
        if h == (Inclined_Bay + 1):
            last_line = block1[-1]
            for k in last_line:
                x1 = k[0]
                y1 = k[1]

                xi = x1 + y1 * math.sin(math.radians(Angle))
                yi = y1 - y1 * math.cos(math.radians(Angle))

                xy = [xi, yi]
                line.append(xy)
        block3.append(line)

        last_line = block3[-1]
        for k in last_line:
            X_spacing = X_BaysSpacing[h]
            x1 = k[0]
            y1 = k[1]

            xi = x1 + X_spacing * math.cos(math.radians(Angle))
            yi = y1 - X_spacing * math.sin(math.radians(Angle))

            xy = [xi, yi]
            line.append(xy)
        block3.append(line)







