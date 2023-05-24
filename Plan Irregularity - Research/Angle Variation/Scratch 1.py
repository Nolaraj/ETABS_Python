if CoOrdinates_Type == 3:
    triangles = []

    blocks = []
    bent_CoOrd = []


    def Computer(Angle):

        Angle_AddGrids = {}
        X_BaysSpacing = []
        Y_BaysSpacing = []
        Origin_X = 0
        Origin_Y = 0

        Inclined_Bay = ws.cell(column=4, row=gdRow + 5).value
        for i in range(4, 13):
            X_BaysSpacing.append(ws.cell(column=i, row=gdRow + 3).value)
            Y_BaysSpacing.append(ws.cell(column=i, row=gdRow + 4).value)
        for row in range(gdRow + 1, gdRow + 7):
            Angle_AddGrids[int(ws.cell(column=14, row=row).value)] = int(ws.cell(column=14 + 1, row=row).value)
        X_Previous = Origin_X
        Y_Previous = Origin_Y

        block1 = []
        block2 = []
        block3 = []
        print("A angle", Angle)

        for h in range(0, baysX + 1):

            # Phase 1 - Before bent
            if h < Inclined_Bay:
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
                print("Before bent angle", Angle)

            if h == Inclined_Bay:
                blocks.append(block1)
                line = []

            # Phase 2 - Bent
            if h == Inclined_Bay:
                # Along X beam in bent portion
                last_line = block1[-1]
                centre_Coordinates = last_line[0]
                for k in last_line:
                    if last_line.index(k) == 0:
                        x1 = k[0]
                        y1 = k[1]
                        xy = [x1, y1]
                        line.append(xy)
                        block2.append(line)  # Line in horizontal dir
                        line = []

                    if last_line.index(k) != 0:
                        x1 = k[0]
                        y1 = k[1]
                        xy = [x1, y1]
                        line.append(xy)

                        X_spacing = X_BaysSpacing[h - 1]
                        R = y1

                        circumference = 2*math.pi*R*Angle/360


                        Sides_No = ceil(circumference / X_spacing)
                        del_Angle = Angle / Sides_No
                        if R > 0:
                            # Result = Chord_Fit(X_spacing, Angle, R)
                            # print((X_spacing, Angle, R))
                            #
                            # Angle_Output = Result[0]
                            # Radius_Output = R
                            # Sides_No = Result[2]
                            #
                            # beta = 180 - (Angle / 2) - ((360 - Angle_Output) / 2)
                            # x0 = centre_Coordinates[0] + Radius_Output * sin(radians(beta))
                            # y0 = R - Radius_Output * cos(radians(beta))
                            #
                            # alpha = Angle_Output / Sides_No




                            for k in range(1, Sides_No + 1):
                                gamma = k * del_Angle
                                xi = x1 + R * sin(radians(gamma))
                                yi = R * cos(radians(gamma))

                                xy = [xi, yi]
                                line.append(xy)
                            block2.append(line)  # Line in horizontal dir
                            line = []

                blocks.append(block2)
                line = []

                # Along Y beam in bent portion
                Y_line = []

                last_line = blocks[-1][-1]
                S_last = blocks[-1][-2]

                a = floor((len(last_line) - 2) / 2)
                for index in range(1, a + 1):
                    start_y_line = []
                    end_y_line = []
                    for lines in block2:
                        if len(lines) > (index * 2):
                            start_y_line.append(lines[index])
                            end_y_line.append(lines[-index - 1])
                    Y_line.append(start_y_line)
                    Y_line.append(end_y_line)

                if len(last_line) % 2 == 1:
                    len_last = len(last_line)
                    len_Slast = len(S_last)

                    # Case 1
                    if len_last == (len_Slast + 1):
                        co_index11 = floor((len_Slast) / 2) - 1
                        co_index12 = floor((len_Slast) / 2)
                        co_index2 = floor(len_last / 2)

                        line.append(S_last[co_index11])
                        line.append(last_line[co_index2])
                        Y_line.append(line)
                        line = []

                        line.append(S_last[co_index12])
                        line.append(last_line[co_index2])
                        Y_line.append(line)
                        line = []

                    # Case 2
                    if Quadrilateral_Boundary is False:
                        if len_last == (len_Slast + 2):
                            co_index11 = floor((len_Slast) / 2)
                            co_index2 = floor(len_last / 2)

                            line.append(S_last[co_index11])
                            line.append(last_line[co_index2])
                            Y_line.append(line)
                            line = []

                blocks.append(Y_line)

            # Phase 3 - After bent

            if h > Inclined_Bay:
                if h == (Inclined_Bay + 1):
                    last_line = block1[-1]
                    for k in last_line:
                        x1 = k[0]
                        y1 = k[1]

                        xi = x1 + y1 * math.sin(math.radians(Angle))
                        yi = y1 * math.cos(math.radians(Angle))

                        # yi = y1 - y1 * math.cos(math.radians(Angle))

                        xy = [xi, yi]
                        line.append(xy)
                block3.append(line)
                line = []

                last_line = block3[-1]
                for k in last_line:
                    X_spacing = X_BaysSpacing[h - 1]
                    x1 = k[0]
                    y1 = k[1]

                    xi = x1 + X_spacing * math.cos(math.radians(Angle))
                    yi = y1 - X_spacing * math.sin(math.radians(Angle))

                    xy = [xi, yi]
                    line.append(xy)
                block3.append(line)

                print("After bent angle", Angle)

            if h == baysX:
                blocks.append(block3)
                line = []


    def Slab_for_bent():
        # Determination of all co_ordinates in bent
        block1 = blocks[1]
        for lines in block1:
            for Co_Ord in lines:
                bent_CoOrd.append(Co_Ord)

        shapes = Triangles_Finder([tuple(x) for x in bent_CoOrd])
        for x in shapes:
            triangles.append(x)


    def BeamColSlab_CodeModelling():

        for h in range(1, storey + 1):
            unique_name = ""
            frame_key = ""
            # Columns Modelling for bent
            for point in bent_CoOrd:
                x1 = point[0]
                y1 = point[1]
                z1 = (h - 1) * storey_height
                x2 = point[0]
                y2 = point[1]
                z2 = h * storey_height
                [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key,
                                                                col_section,
                                                                unique_name)

            # Slabs Modelling for bent portion
            for triangle in triangles:
                x1 = triangle[0][0]
                y1 = triangle[0][1]
                z1 = h * storey_height
                x2 = triangle[1][0]
                y2 = triangle[1][1]
                z2 = h * storey_height
                x3 = triangle[2][0]
                y3 = triangle[2][1]
                z3 = h * storey_height

                x = [x1, x2, x3, x1]
                y = [y1, y2, y3, y1]
                z = [z1, z2, z3, z1]
                slab_fname = " "

                ret = SapModel.AreaObj.AddByCoord(4, x, y, z, slab_fname, slab_section, unique_name)

        for block in blocks:
            for h in range(1, storey + 1):
                if (blocks.index(block) == 1 or blocks.index(block) == 2):  # Bent type of block - horizontal line
                    # ALong X (block1) and Y Beams (block2) in bent
                    for line in block:
                        for i in range(0, len(line) - 1):
                            co_ordinates1 = line[i]
                            co_ordinates2 = line[i + 1]

                            x1 = co_ordinates1[0]
                            y1 = co_ordinates1[1]
                            z1 = h * storey_height
                            x2 = co_ordinates2[0]
                            y2 = co_ordinates2[1]
                            z2 = h * storey_height

                            y1_key = alphabet[block.index(line) + 1]
                            y2_key = alphabet[block.index(line) + 1]

                            x1_key = alphabet[line.index(co_ordinates1) + 1]
                            x2_key = alphabet[line.index(co_ordinates2) + 1]

                            z1_key = alphabet[h]
                            z2_key = alphabet[h]

                            frame_key = f'{z1_key}{y1_key}{x1_key}to{z2_key}{y2_key}{x2_key}'
                            unique_name = f'{z1_key}{y1_key}{x1_key}{z2_key}{y2_key}{x2_key}'
                            [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key,
                                                                            beam_section,
                                                                            unique_name)

                if (blocks.index(block) == 0 or blocks.index(block) == 3):
                    # ALong Y Beams
                    for line in block:
                        for i in range(0, len(line) - 1):
                            co_ordinates1 = line[i]
                            co_ordinates2 = line[i + 1]

                            x1 = co_ordinates1[0]
                            y1 = co_ordinates1[1]
                            z1 = h * storey_height
                            x2 = co_ordinates2[0]
                            y2 = co_ordinates2[1]
                            z2 = h * storey_height

                            if blocks.index(block) == 0:
                                x1_key = block.index(line) + 1
                                x2_key = block.index(line) + 1

                            if blocks.index(block) == 2:
                                x1_key = len(blocks[0]) + block.index(line) + 1
                                x2_key = len(blocks[0]) + block.index(line) + 1

                            y1_key = line.index(co_ordinates1) + 1
                            y2_key = line.index(co_ordinates2) + 1

                            z1_key = h + 1
                            z2_key = h + 1

                            frame_key = f'{z1_key}{y1_key}{x1_key}to{z2_key}{y2_key}{x2_key}'
                            unique_name = f'{z1_key}{y1_key}{x1_key}{z2_key}{y2_key}{x2_key}'
                            [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key,
                                                                            beam_section,
                                                                            unique_name)

                    # ALong X Beams
                    for i in range(0, len(line)):
                        for j in range(0, len(block) - 1):
                            line1 = block[j]
                            line2 = block[j + 1]
                            co_ordinates1 = line1[i]
                            co_ordinates2 = line2[i]

                            x1 = co_ordinates1[0]
                            y1 = co_ordinates1[1]
                            z1 = h * storey_height
                            x2 = co_ordinates2[0]
                            y2 = co_ordinates2[1]
                            z2 = h * storey_height

                            if blocks.index(block) == 0:
                                x1_key = block.index(line) + 1
                                x2_key = block.index(line) + 1

                            if blocks.index(block) == 2:
                                x1_key = len(blocks[0]) + block.index(line) + 1
                                x2_key = len(blocks[0]) + block.index(line) + 1

                            y1_key = line1.index(co_ordinates1) + 1
                            y2_key = line2.index(co_ordinates2) + 1

                            z1_key = h + 1
                            z2_key = h + 1

                            frame_key = f'{z1_key}{y1_key}{x1_key}to{z2_key}{y2_key}{x2_key}'
                            unique_name = f'{z1_key}{y1_key}{x1_key}{z2_key}{y2_key}{x2_key}'
                            [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key,
                                                                            beam_section,
                                                                            unique_name)

                    # ALong Z Columns
                    for line in block:
                        for i in range(0, len(line)):
                            co_ordinates1 = line[i]

                            x1 = co_ordinates1[0]
                            y1 = co_ordinates1[1]
                            z1 = (h - 1) * storey_height
                            x2 = co_ordinates1[0]
                            y2 = co_ordinates1[1]
                            z2 = h * storey_height

                            if blocks.index(block) == 0:
                                x1_key = block.index(line) + 1
                                x2_key = block.index(line) + 1

                            if blocks.index(block) == 2:
                                x1_key = len(blocks[0]) + block.index(line) + 1
                                x2_key = len(blocks[0]) + block.index(line) + 1

                            y1_key = line.index(co_ordinates1) + 1
                            y2_key = line.index(co_ordinates1) + 1

                            z1_key = h
                            z2_key = h + 1

                            frame_key = f'{z1_key}{y1_key}{x1_key}to{z2_key}{y2_key}{x2_key}'
                            unique_name = f'{z1_key}{y1_key}{x1_key}{z2_key}{y2_key}{x2_key}'
                            [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key,
                                                                            col_section,
                                                                            unique_name)

                    # Slabs Modelling
                    for i in range(0, len(block) - 1):
                        for j in range(0, len(line) - 1):
                            line1 = block[i]
                            line2 = block[i + 1]

                            co_ordinates1 = line1[j]
                            co_ordinates2 = line1[j + 1]
                            co_ordinates3 = line2[j + 1]
                            co_ordinates4 = line2[j]

                            x1 = co_ordinates1[0]
                            y1 = co_ordinates1[1]
                            z1 = h * storey_height
                            x2 = co_ordinates2[0]
                            y2 = co_ordinates2[1]
                            z2 = h * storey_height
                            x3 = co_ordinates3[0]
                            y3 = co_ordinates3[1]
                            z3 = h * storey_height
                            x4 = co_ordinates4[0]
                            y4 = co_ordinates4[1]
                            z4 = h * storey_height

                            x = [x1, x2, x3, x4]
                            y = [y1, y2, y3, y4]
                            z = [z1, z2, z3, z4]

                            unique_name = f'{h}{alphabet[i]}{number[j]}'
                            slab_fname = f'{h}{i + 1}{j + 1}{h}{i + 1}{j + 2}{h}{i + 2}{j + 2}{h}{i + 2}{j + 1}'
                            ret = SapModel.AreaObj.AddByCoord(4, x, y, z, slab_fname, slab_section, unique_name)


    Computer(Angle)
    Slab_for_bent()
    BeamColSlab_CodeModelling()