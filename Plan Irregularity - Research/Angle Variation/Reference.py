
X_CoOrdiatesFloor = []
Y_CoOrdiatesFloor = []
storey_height = 3
storey = 7



















def beamcolumnslab_modelling():
    for h in range(1, storey + 1):

        # Beams along X Modelling
        for i in range(0, len(X_CoOrdiatesFloor[0])):
            for j in range(0, len(X_CoOrdiatesFloor)):
                try:
                    x1 = X_CoOrdiatesFloor[j][i]
                    y1 = Y_CoOrdiatesFloor[j][i]
                    z1 = h * storey_height
                    x2 = X_CoOrdiatesFloor[j + 1][i]
                    y2 = Y_CoOrdiatesFloor[j + 1][i]
                    z2 = h * storey_height

                    frame_key = f'{h + 1}{i + 1}{j + 1}to{h}{i + 2}{j + 2}'
                    unique_name = f'{h + 1}{i + 1}{j + 1}{h}{i + 2}{j + 2}'
                    [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key, beam_section,
                                                                    unique_name)
                except:
                    pass

        # Beams along Y Modelling
        for i in range(0, len(X_CoOrdiatesFloor)):
            for j in range(0, len(X_CoOrdiatesFloor[i])):
                try:
                    x1 = X_CoOrdiatesFloor[i][j]
                    y1 = Y_CoOrdiatesFloor[i][j]
                    z1 = h * storey_height
                    x2 = X_CoOrdiatesFloor[i][j + 1]
                    y2 = Y_CoOrdiatesFloor[i][j + 1]
                    z2 = h * storey_height

                    frame_key = f'{h +1}{j+1}{i+1}to{h + 1}{j+2}{i+2}'
                    unique_name = f'{h + 1}{j+1}{i+1}{h + 1}{j+2}{i+2}'
                    [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key, beam_section,
                                                                    unique_name)
                except:
                    pass


        # Columns along Z Modelling
        for i in range(0, len(X_CoOrdiatesFloor[0])):
            for j in range(0, len(X_CoOrdiatesFloor)):
                try:
                    x1 = X_CoOrdiatesFloor[j][i]
                    y1 = Y_CoOrdiatesFloor[j][i]
                    z1 = (h - 1) * storey_height
                    x2 = X_CoOrdiatesFloor[j][i]
                    y2 = Y_CoOrdiatesFloor[j][i]
                    z2 = h * storey_height

                    frame_key = f'{h}{i+1}{j+1}to{h + 1}{i+1}{j+1}'
                    unique_name = f'{h}{i+1}{j+1}{h + 1}{i+1}{j+1}'
                    [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key, col_section,
                                                                    unique_name)
                except:
                    pass



        # Slabs Modelling
        for i in range(0, len(X_CoOrdiatesFloor[0])):
            for j in range(0, len(X_CoOrdiatesFloor)):
                try:
                    x1 = X_CoOrdiatesFloor[j][i]
                    y1 = Y_CoOrdiatesFloor[j][i]
                    z1 = h * storey_height
                    x2 = X_CoOrdiatesFloor[j + 1][i]
                    y2 = Y_CoOrdiatesFloor[j + 1][i]
                    z2 = h * storey_height
                    x3 = X_CoOrdiatesFloor[j + 1][i+1]
                    y3 = Y_CoOrdiatesFloor[j + 1][i+1]
                    z3 = h * storey_height
                    x4 = X_CoOrdiatesFloor[j ][i + 1]
                    y4 = Y_CoOrdiatesFloor[j ][i + 1]
                    z4 = h * storey_height

                    x = [x1, x2, x3, x4]
                    y = [y1, y2, y3, y4]
                    z = [z1, z2, z3, z4]

                    unique_name = f'{h}{alphabet[j]}{number[i]})'
                    slab_fname =  f'{h}{i+1}{j+1}{h}{i + 1}{j + 2}{h}{i+2}{j+2}{h}{i+2}{j+1}'
                    ret = SapModel.AreaObj.AddByCoord(4, x, y, z, slab_fname, slab_section, unique_name)




                    frame_key = f'{h + 1}{i + 1}{j + 1}to{h}{i + 2}{j + 2}'
                    unique_name = f'{h + 1}{i + 1}{j + 1}{h}{i + 2}{j + 2}'
                    [frame_key, ret] = SapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frame_key, beam_section,
                                                                    unique_name)
                except:
                    pass





