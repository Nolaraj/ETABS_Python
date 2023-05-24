# Co Ordinates Modelling
# Level1 = Floor wise Co ordinates, Level 2 = Along Horizontal bays co, ordinates  [[], [], []]
# Similarly same for y co_ordinates

baysX = 4
baysY = 2
import math
storey = 2
storey_height = 3

Angle = 30

Angle_AddGrids = {}
X_BaysSpacing = [5,6,2,8,9]
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


    # Phase 1 - Before bent
    if h < Inclined_Bay :
        Xs = []
        Ys = []
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
        line = []

    # Phase 2 - Bent
    if h == Inclined_Bay :
        last_line = block1[-1]
        for k in last_line:
            x1 = k[0]
            y1 = k[1]

            X_spacing = X_BaysSpacing[h-1]
            l = y1

            D = ((l*math.sin(math.radians(Angle)))**2 + ((l-l*math.cos(math.radians(Angle))))**2)**0.5
            s = math.ceil(D/X_spacing)
            print(D, s, X_spacing, h)
            if s==0:
                beta = 0

            else:
                beta = Angle / s

            xy = [x1, y1]
            line.append(xy)
            for k in range(1, s+1):
                xi = x1 + y1*math.sin(math.radians(beta))
                yi = y1 - y1*math.cos(math.radians(beta))

                xy = [xi, yi]
                line.append(xy)
            block2.append(line)             #Line in horizontal dir
            line = []
        blocks.append(block2)
        line = []
        print(block2)













    # Phase 3 - After bent

    if h > Inclined_Bay:
        if h == (Inclined_Bay + 1):
            last_line = block1[-1]
            for k in last_line:
                # print(h,1,k)
                x1 = k[0]
                y1 = k[1]

                xi = x1 + y1 * math.sin(math.radians(Angle))
                yi = y1 - y1 * math.cos(math.radians(Angle))

                xy = [xi, yi]
                line.append(xy)
        block3.append(line)
        line = []


        last_line = block3[-1]
        for k in last_line:
            X_spacing = X_BaysSpacing[h-1]
            x1 = k[0]
            y1 = k[1]

            xi = x1 + X_spacing * math.cos(math.radians(Angle))
            yi = y1 - X_spacing * math.sin(math.radians(Angle))

            xy = [xi, yi]
            line.append(xy)
        block3.append(line)

    if h == baysX:
        blocks.append(block3)
        line = []

for block in blocks:
    for h in range(1, storey + 1):
        if block.index() == 1:      #Bent type of block - horizontal line


        if block.index() == 0:
            # ALong Y Beams
            for line in block:
                for i in range(0, len(line)-1):
                    co_ordinates1 = line[i]
                    co_ordinates2 = line[i+1]


                    x1 = co_ordinates1[0]
                    y1 = co_ordinates1[1]
                    z1 = h * storey_height
                    x1 = co_ordinates2[0]
                    y1 = co_ordinates2[1]
                    z1 = h * storey_height

                    x1_key = line.index()
                    y1_key = co_ordinates1.index()
                    z1_key = h

                    x2_key = line.index()
                    y2_key = co_ordinates2.index()
                    z2_key = h

                    frame_key = f'{z1_key}{y1_key}{x1_key}to{z2_key}{y2_key}{x2_key}'
                    unique_name = f'{z1_key}{y1_key}{x1_key}to{z2_key}{y2_key}{x2_key}'
                    [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key,
                                                                    beam_section,
                                                                    unique_name)



# print(blocks)