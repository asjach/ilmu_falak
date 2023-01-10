# TEST ELP1
from terms_vsop87_elp2000 import *
from math import radians, cos, sin, atan, sqrt
from khafid_fungsi import julian_ephemeris_day, julian_datum


def ELP2000(tjj):
    cpi = 4*atan(1)
    cpi2 = 2*cpi
    pis2 = cpi/2
        
    rad = 648000/cpi
    deg = cpi/180
    c1 = 60
    c2 = 3600
    ath = 384747.980674317
    a0 = 384747.980644895
    am = 0.074801329518
    alfa = 0.002571881335
    dtasm = 2*alfa/(3*am)

    #  Lunar arguments.
    w = [
            [(218 + 18/c1 + 59.95571/c2)*deg, 1732559343.73604/rad, -5.8883/rad, 0.006604/rad, -0.00003169/rad],
            [(83 + 21/c1 + 11.67475/c2)*deg, 14643420.2632/rad, -38.2776/rad, -0.045047/rad, 0.00021301/rad], 
            [(125 + 2/c1 + 40.39816/c2)*deg, -6967919.3622/rad, 6.3622/rad, 0.007625/rad, -0.00003586/rad]
        ]

    eart = [(100 + 27/c1 + 59.22059/c2)*deg, 129597742.2758/rad, -0.0202/rad, 0.000009/rad, 0.00000015/rad]
    peri = [(102 + 56/c1 + 14.42753/c2)*deg, 1161.2283/rad, 0.5327/rad, -0.000138/rad, 0]


    # Planetary arguments.
    preces = 5029.0966/rad
    p11 = (252 + 15/c1 + 3.25986/c2)*deg
    p21 = (181 + 58/c1 + 47.28305/c2)*deg
    p31 = eart[0]
    p41 = (355 + 25/c1 + 59.78866/c2)*deg
    p51 = (34 + 21/c1 + 5.34212/c2)*deg
    p61 = (50 + 4/c1 + 38.89694/c2)*deg
    p71 = (314 + 3/c1 + 18.01841/c2)*deg
    p81 = (304 + 20/c1 + 55.19575/c2)*deg
    p12 = 538101628.68898/rad
    p22 = 210664136.43355/rad
    p32 = eart[1]
    p42 = 68905077.59284/rad
    p52 = 10925660.42861/rad
    p62 = 4399609.65932/rad
    p72 = 1542481.19393/rad
    p82 = 786550.32074/rad

    #Corrections of the constants (fit to DE200/LE200).
    delnu = 0.55604/rad/w[0][1]
    dele = 0.01789/rad
    delg = -0.08066/rad
    delnp = -0.06424/rad/w[0][1]
    delep = -0.12879/rad

    #Delaunays arguments.

    d11 = w[0][0] - eart[0]
    d41 = w[0][0] - w[2][0]
    d31 = w[0][0] - w[1][0]
    d21 = eart[0] - peri[0]
    d12 = w[0][1] - eart[1]
    d42 = w[0][1] - w[2][1]
    d32 = w[0][1] - w[1][1]
    d22 = eart[1] - peri[1]
    d13 = w[0][2] - eart[2]
    d43 = w[0][2] - w[2][2]
    d33 = w[0][2] - w[1][2]
    d23 = eart[2] - peri[2]
    d14 = w[0][3] - eart[3]
    d44 = w[0][3] - w[2][3]
    d34 = w[0][3] - w[1][3]
    d24 = eart[3] - peri[3]
    d15 = w[0][4] - eart[4]
    d45 = w[0][4] - w[2][4]
    d35 = w[0][4] - w[1][4]
    d25 = eart[4] - peri[4]
    

    d11 = d11 + cpi
    z1 = w[0][0]
    z2 = w[0][1] + preces


    #  matrix nilai delaunays argument
    # matrix_delaunays = [
    #     [d11, d12, d13, d14, d15],
    #     [d21, d22, d23, d24, d25],
    #     [d31, d32, d33, d34, d35],
    #     [d41, d42, d43, d44, d45],
        
    # ]

    matrix_delaunays = [
        [d11, d21, d31, d41],
        [d12, d22, d32, d42],
        [d13, d23, d33, d43],
        [d14, d24, d34, d44],
        [d15, d25, d35, d45]
    ]

    #  Precession matrix.

    p1 = 0.000010180391
    p2 = 0.00000047020439
    p3 = -5.417367E-10
    p4 = -2.507948E-12
    p5 = 4.63486E-15
    q1 = -0.000113469002
    q2 = 0.00000012372674
    q3 = 0.000000001265417
    q4 = -1.371808E-12
    q5 = -3.20334E-15

    T1 = 1

    T2 = (tjj - 2451545)/36525
    T3 = T2*T2
    T4 = T3*T2
    T5 = T4*T2

    Xe = Ye = Ze = 0

    elp1 = ELP200082B_ELP1
    for i in range(len(elp1)):
        Y = 0
        Y += elp1[i][0]*d11*T1 + elp1[i][1]*d21*T1 + elp1[i][2]*d31*T1 +elp1[i][3]*d41*T1
        Y += elp1[i][0]*d12*T2 + elp1[i][1]*d22*T2 + elp1[i][2]*d32*T2 +elp1[i][3]*d42*T2
        Y += elp1[i][0]*d13*T3 + elp1[i][1]*d23*T3 + elp1[i][2]*d33*T3 +elp1[i][3]*d43*T3
        Y += elp1[i][0]*d14*T4 + elp1[i][1]*d24*T4 + elp1[i][2]*d34*T4 +elp1[i][3]*d44*T4
        Y += elp1[i][0]*d15*T5 + elp1[i][1]*d25*T5 + elp1[i][2]*d35*T5 +elp1[i][3]*d45*T5
        Xe += (elp1[i][4] + (elp1[i][5] + dtasm*elp1[i][9])*(delnp - am*delnu) + elp1[i][6]*delg + elp1[i][7]*dele + elp1[i][8]*delep)*sin(Y)

    elp2 = ELP200082B_ELP2
    for i in range(len(elp2)):
        Y = 0
        Y += elp2[i][0]*d11*T1 + elp2[i][1]*d21*T1 + elp2[i][2]*d31*T1 +elp2[i][3]*d41*T1
        Y += elp2[i][0]*d12*T2 + elp2[i][1]*d22*T2 + elp2[i][2]*d32*T2 +elp2[i][3]*d42*T2
        Y += elp2[i][0]*d13*T3 + elp2[i][1]*d23*T3 + elp2[i][2]*d33*T3 +elp2[i][3]*d43*T3
        Y += elp2[i][0]*d14*T4 + elp2[i][1]*d24*T4 + elp2[i][2]*d34*T4 +elp2[i][3]*d44*T4
        Y += elp2[i][0]*d15*T5 + elp2[i][1]*d25*T5 + elp2[i][2]*d35*T5 +elp2[i][3]*d45*T5
        Ye += (elp2[i][4] + (elp2[i][5] + dtasm*elp2[i][9])*(delnp - am*delnu) + elp2[i][6]*delg + elp2[i][7]*dele + elp2[i][8]*delep)*sin(Y)

    elp3 = ELP200082B_ELP3
    for i in range(len(elp3)):
        Y = 0
        Y += elp3[i][0]*d11*T1 + elp3[i][1]*d21*T1 + elp3[i][2]*d31*T1 + elp3[i][3]*d41*T1
        Y += elp3[i][0]*d12*T2 + elp3[i][1]*d22*T2 + elp3[i][2]*d32*T2 + elp3[i][3]*d42*T2
        Y += elp3[i][0]*d13*T3 + elp3[i][1]*d23*T3 + elp3[i][2]*d33*T3 + elp3[i][3]*d43*T3
        Y += elp3[i][0]*d14*T4 + elp3[i][1]*d24*T4 + elp3[i][2]*d34*T4 + elp3[i][3]*d44*T4
        Y += elp3[i][0]*d15*T5 + elp3[i][1]*d25*T5 + elp3[i][2]*d35*T5 + elp3[i][3]*d45*T5
        Ze += (elp3[i][4] + (elp3[i][5] + dtasm*elp3[i][9])*(delnp - am*delnu) + elp3[i][6]*delg + elp3[i][7]*dele + elp3[i][8]*delep)*cos(Y)

    elp4 = ELP200082B_ELP4
    for i in range(len(elp4)):
        Y = 0
        Y += radians(elp4[i][5]) + elp4[i][0]*z1*T1 + elp4[i][1]*d11*T1 + elp4[i][2]*d21*T1 + elp4[i][3]*d31*T1 + elp4[i][4]*d41*T1
        Y += elp4[i][0]*z2*T2 + elp4[i][1]*d12*T2 + elp4[i][2]*d22*T2 + elp4[i][3]*d32*T2 + elp4[i][4]*d42*T2
        Xe += elp4[i][6]*sin(Y)

    elp5 = ELP200082B_ELP5
    for i in range(len(elp5)):
        Y = 0
        Y += radians(elp5[i][5]) + elp5[i][0]*z1*T1 + elp5[i][1]*d11*T1 + elp5[i][2]*d21*T1 + elp5[i][3]*d31*T1 + elp5[i][4]*d41*T1
        Y += elp5[i][0]*z2*T2 + elp5[i][1]*d12*T2 + elp5[i][2]*d22*T2 + elp5[i][3]*d32*T2 + elp5[i][4]*d42*T2
        Ye += elp5[i][6]*sin(Y)

    elp6 = ELP200082B_ELP6
    for i in range(len(elp6)):
        Y = 0
        Y += radians(elp6[i][5]) + elp6[i][0]*z1*T1 + elp6[i][1]*d11*T1 + elp6[i][2]*d21*T1 + elp6[i][3]*d31*T1 + elp6[i][4]*d41*T1
        Y += elp6[i][0]*z2*T2 + elp6[i][1]*d12*T2 + elp6[i][2]*d22*T2 + elp6[i][3]*d32*T2 + elp6[i][4]*d42*T2
        Ze += elp6[i][6]*sin(Y)

    elp7 = ELP200082B_ELP7
    for i in range(len(elp7)):
        Y = 0
        Y += radians(elp7[i][5]) + elp7[i][0]*z1*T1 + elp7[i][1]*d11*T1 + elp7[i][2]*d21*T1 + elp7[i][3]*d31*T1 + elp7[i][4]*d41*T1
        Y += elp7[i][0]*z2*T2 + elp7[i][1]*d12*T2 + elp7[i][2]*d22*T2 + elp7[i][3]*d32*T2 + elp7[i][4]*d42*T2
        Xe += elp7[i][6]*T2*sin(Y)

    elp8 = ELP200082B_ELP8
    for i in range(len(elp8)):
        Y = 0
        Y += radians(elp8[i][5]) + elp8[i][0]*z1*T1 + elp8[i][1]*d11*T1 + elp8[i][2]*d21*T1 + elp8[i][3]*d31*T1 + elp8[i][4]*d41*T1
        Y += elp8[i][0]*z2*T2 + elp8[i][1]*d12*T2 + elp8[i][2]*d22*T2 + elp8[i][3]*d32*T2 + elp8[i][4]*d42*T2
        Ye += elp8[i][6]*T2*sin(Y)

    elp9 = ELP200082B_ELP9
    for i in range(len(elp9)):
        Y = 0
        Y += radians(elp9[i][5]) + elp9[i][0]*z1*T1 + elp9[i][1]*d11*T1 + elp9[i][2]*d21*T1 + elp9[i][3]*d31*T1 + elp9[i][4]*d41*T1
        Y += elp9[i][0]*z2*T2 + elp9[i][1]*d12*T2 + elp9[i][2]*d22*T2 + elp9[i][3]*d32*T2 + elp9[i][4]*d42*T2
        Ze += elp9[i][6]*T2*sin(Y)
 
    elp10 = ELP200082B_ELP10
    for i in range(len(elp10)):
        Y = 0
        Y += elp10[i][11]*deg
        Y += elp10[i][0]*p11*T1 + elp10[i][1]*p21*T1 + elp10[i][2]*p31*T1 + elp10[i][3]*p41*T1 + elp10[i][4]*p51*T1 + elp10[i][5]*p61*T1 + elp10[i][6]*p71*T1 + elp10[i][7]*p81*T1 + elp10[i][8]*d11*T1 + elp10[i][9]*d31*T1 + elp10[i][10]*d41*T1
        Y += elp10[i][0]*p12*T2 + elp10[i][1]*p22*T2 + elp10[i][2]*p32*T2 + elp10[i][3]*p42*T2 + elp10[i][4]*p52*T2 + elp10[i][5]*p62*T2 + elp10[i][6]*p72*T2 + elp10[i][7]*p82*T2 + elp10[i][8]*d12*T2 + elp10[i][9]*d32*T2 + elp10[i][10]*d42*T2
        Xe += elp10[i][12]*sin(Y)

    elp11 = ELP200082B_ELP11
    for i in range(len(elp11)):
        Y = 0
        Y += elp11[i][11]*deg
        Y += elp11[i][0]*p11*T1 + elp11[i][1]*p21*T1 + elp11[i][2]*p31*T1 + elp11[i][3]*p41*T1 + elp11[i][4]*p51*T1 + elp11[i][5]*p61*T1 + elp11[i][6]*p71*T1 + elp11[i][7]*p81*T1 + elp11[i][8]*d11*T1 + elp11[i][9]*d31*T1 + elp11[i][10]*d41*T1
        Y += elp11[i][0]*p12*T2 + elp11[i][1]*p22*T2 + elp11[i][2]*p32*T2 + elp11[i][3]*p42*T2 + elp11[i][4]*p52*T2 + elp11[i][5]*p62*T2 + elp11[i][6]*p72*T2 + elp11[i][7]*p82*T2 + elp11[i][8]*d12*T2 + elp11[i][9]*d32*T2 + elp11[i][10]*d42*T2
        Ye += elp11[i][12]*sin(Y)

    elp12 = ELP200082B_ELP12
    for i in range(len(elp12)):
        Y = 0
        Y += elp12[i][11]*deg
        Y += elp12[i][0]*p11*T1 + elp12[i][1]*p21*T1 + elp12[i][2]*p31*T1 + elp12[i][3]*p41*T1 + elp12[i][4]*p51*T1 + elp12[i][5]*p61*T1 + elp12[i][6]*p71*T1 + elp12[i][7]*p81*T1 + elp12[i][8]*d11*T1 + elp12[i][9]*d31*T1 + elp12[i][10]*d41*T1
        Y += elp12[i][0]*p12*T2 + elp12[i][1]*p22*T2 + elp12[i][2]*p32*T2 + elp12[i][3]*p42*T2 + elp12[i][4]*p52*T2 + elp12[i][5]*p62*T2 + elp12[i][6]*p72*T2 + elp12[i][7]*p82*T2 + elp12[i][8]*d12*T2 + elp12[i][9]*d32*T2 + elp12[i][10]*d42*T2
        Ze += elp12[i][12]*sin(Y)

    elp13 = ELP200082B_ELP13
    for i in range(len(elp13)):
        Y = 0
        Y += elp13[i][11]*deg
        Y += elp13[i][0]*p11*T1 + elp13[i][1]*p21*T1 + elp13[i][2]*p31*T1 + elp13[i][3]*p41*T1 + elp13[i][4]*p51*T1 + elp13[i][5]*p61*T1 + elp13[i][6]*p71*T1 + elp13[i][7]*p81*T1 + elp13[i][8]*d11*T1 + elp13[i][9]*d31*T1 + elp13[i][10]*d41*T1
        Y += elp13[i][0]*p12*T2 + elp13[i][1]*p22*T2 + elp13[i][2]*p32*T2 + elp13[i][3]*p42*T2 + elp13[i][4]*p52*T2 + elp13[i][5]*p62*T2 + elp13[i][6]*p72*T2 + elp13[i][7]*p82*T2 + elp13[i][8]*d12*T2 + elp13[i][9]*d32*T2 + elp13[i][10]*d42*T2
        Xe += elp13[i][12]*T2*sin(Y)

    elp14 = ELP200082B_ELP14
    for i in range(len(elp14)):
        Y = 0
        Y += elp14[i][11]*deg
        Y += elp14[i][0]*p11*T1 + elp14[i][1]*p21*T1 + elp14[i][2]*p31*T1 + elp14[i][3]*p41*T1 + elp14[i][4]*p51*T1 + elp14[i][5]*p61*T1 + elp14[i][6]*p71*T1 + elp14[i][7]*p81*T1 + elp14[i][8]*d11*T1 + elp14[i][9]*d31*T1 + elp14[i][10]*d41*T1
        Y += elp14[i][0]*p12*T2 + elp14[i][1]*p22*T2 + elp14[i][2]*p32*T2 + elp14[i][3]*p42*T2 + elp14[i][4]*p52*T2 + elp14[i][5]*p62*T2 + elp14[i][6]*p72*T2 + elp14[i][7]*p82*T2 + elp14[i][8]*d12*T2 + elp14[i][9]*d32*T2 + elp14[i][10]*d42*T2
        Ye += elp14[i][12]*T2*sin(Y)

    elp15 = ELP200082B_ELP15
    for i in range(len(elp15)):
        Y = 0
        Y += elp15[i][11]*deg
        Y += elp15[i][0]*p11*T1 + elp15[i][1]*p21*T1 + elp15[i][2]*p31*T1 + elp15[i][3]*p41*T1 + elp15[i][4]*p51*T1 + elp15[i][5]*p61*T1 + elp15[i][6]*p71*T1 + elp15[i][7]*p81*T1 + elp15[i][8]*d11*T1 + elp15[i][9]*d31*T1 + elp15[i][10]*d41*T1
        Y += elp15[i][0]*p12*T2 + elp15[i][1]*p22*T2 + elp15[i][2]*p32*T2 + elp15[i][3]*p42*T2 + elp15[i][4]*p52*T2 + elp15[i][5]*p62*T2 + elp15[i][6]*p72*T2 + elp15[i][7]*p82*T2 + elp15[i][8]*d12*T2 + elp15[i][9]*d32*T2 + elp15[i][10]*d42*T2
        Ze += (elp15[i][12])*T2*sin(Y)

    elp16 = ELP200082B_ELP16
    for i in range(len(elp16)):
        Y = 0
        Y += elp16[i][11]*deg
        Y += elp16[i][0]*p11*T1 + elp16[i][1]*p21*T1 + elp16[i][2]*p31*T1 + elp16[i][3]*p41*T1 + elp16[i][4]*p51*T1 + elp16[i][5]*p61*T1 + elp16[i][6]*p71*T1 + elp16[i][7]*d11*T1 + elp16[i][8]*d21*T1 + elp16[i][9]*d31*T1 + elp16[i][10]*d41*T1
        Y += elp16[i][0]*p12*T2 + elp16[i][1]*p22*T2 + elp16[i][2]*p32*T2 + elp16[i][3]*p42*T2 + elp16[i][4]*p52*T2 + elp16[i][5]*p62*T2 + elp16[i][6]*p72*T2 + elp16[i][7]*d12*T2 + elp16[i][8]*d22*T2 + elp16[i][9]*d32*T2 + elp16[i][10]*d42*T2
        Xe += elp16[i][12]*sin(Y)

    elp17 = ELP200082B_ELP17
    for i in range(len(elp17)):
        Y = 0
        Y += elp17[i][11]*deg
        Y += elp17[i][0]*p11*T1 + elp17[i][1]*p21*T1 + elp17[i][2]*p31*T1 + elp17[i][3]*p41*T1 + elp17[i][4]*p51*T1 + elp17[i][5]*p61*T1 + elp17[i][6]*p71*T1 + elp17[i][7]*d11*T1 + elp17[i][8]*d21*T1 + elp17[i][9]*d31*T1 + elp17[i][10]*d41*T1
        Y += elp17[i][0]*p12*T2 + elp17[i][1]*p22*T2 + elp17[i][2]*p32*T2 + elp17[i][3]*p42*T2 + elp17[i][4]*p52*T2 + elp17[i][5]*p62*T2 + elp17[i][6]*p72*T2 + elp17[i][7]*d12*T2 + elp17[i][8]*d22*T2 + elp17[i][9]*d32*T2 + elp17[i][10]*d42*T2
        Ye += elp17[i][12]*sin(Y)

    elp18 = ELP200082B_ELP18
    for i in range(len(elp18)):
        Y = 0
        Y += elp18[i][11]*deg
        Y += elp18[i][0]*p11*T1 + elp18[i][1]*p21*T1 + elp18[i][2]*p31*T1 + elp18[i][3]*p41*T1 + elp18[i][4]*p51*T1 + elp18[i][5]*p61*T1 + elp18[i][6]*p71*T1 + elp18[i][7]*d11*T1 + elp18[i][8]*d21*T1 + elp18[i][9]*d31*T1 + elp18[i][10]*d41*T1
        Y += elp18[i][0]*p12*T2 + elp18[i][1]*p22*T2 + elp18[i][2]*p32*T2 + elp18[i][3]*p42*T2 + elp18[i][4]*p52*T2 + elp18[i][5]*p62*T2 + elp18[i][6]*p72*T2 + elp18[i][7]*d12*T2 + elp18[i][8]*d22*T2 + elp18[i][9]*d32*T2 + elp18[i][10]*d42*T2
        Ze += elp18[i][12]*sin(Y)

    elp19 = ELP200082B_ELP19
    for i in range(len(elp19)):
        Y = 0
        Y += elp19[i][11]*deg
        Y += elp19[i][0]*p11*T1 + elp19[i][1]*p21*T1 + elp19[i][2]*p31*T1 + elp19[i][3]*p41*T1 + elp19[i][4]*p51*T1 + elp19[i][5]*p61*T1 + elp19[i][6]*p71*T1 + elp19[i][7]*d11*T1 + elp19[i][8]*d21*T1 + elp19[i][9]*d31*T1 + elp19[i][10]*d41*T1
        Y += elp19[i][0]*p12*T2 + elp19[i][1]*p22*T2 + elp19[i][2]*p32*T2 + elp19[i][3]*p42*T2 + elp19[i][4]*p52*T2 + elp19[i][5]*p62*T2 + elp19[i][6]*p72*T2 + elp19[i][7]*d12*T2 + elp19[i][8]*d22*T2 + elp19[i][9]*d32*T2 + elp19[i][10]*d42*T2
        Xe += elp19[i][12]*T2*sin(Y)

    elp20 = ELP200082B_ELP20
    for i in range(len(elp20)):
        Y = 0
        Y += elp20[i][11]*deg
        Y += elp20[i][0]*p11*T1 + elp20[i][1]*p21*T1 + elp20[i][2]*p31*T1 + elp20[i][3]*p41*T1 + elp20[i][4]*p51*T1 + elp20[i][5]*p61*T1 + elp20[i][6]*p71*T1 + elp20[i][7]*d11*T1 + elp20[i][8]*d21*T1 + elp20[i][9]*d31*T1 + elp20[i][10]*d41*T1
        Y += elp20[i][0]*p12*T2 + elp20[i][1]*p22*T2 + elp20[i][2]*p32*T2 + elp20[i][3]*p42*T2 + elp20[i][4]*p52*T2 + elp20[i][5]*p62*T2 + elp20[i][6]*p72*T2 + elp20[i][7]*d12*T2 + elp20[i][8]*d22*T2 + elp20[i][9]*d32*T2 + elp20[i][10]*d42*T2
        Ye += elp20[i][12]*T2*sin(Y)

    elp21 = ELP200082B_ELP21
    for i in range(len(elp21)):
        Y = 0
        Y += elp21[i][11]*deg
        Y += elp21[i][0]*p11*T1 + elp21[i][1]*p21*T1 + elp21[i][2]*p31*T1 + elp21[i][3]*p41*T1 + elp21[i][4]*p51*T1 + elp21[i][5]*p61*T1 + elp21[i][6]*p71*T1 + elp21[i][7]*d11*T1 + elp21[i][8]*d21*T1 + elp21[i][9]*d31*T1 + elp21[i][10]*d41*T1
        Y += elp21[i][0]*p12*T2 + elp21[i][1]*p22*T2 + elp21[i][2]*p32*T2 + elp21[i][3]*p42*T2 + elp21[i][4]*p52*T2 + elp21[i][5]*p62*T2 + elp21[i][6]*p72*T2 + elp21[i][7]*d12*T2 + elp21[i][8]*d22*T2 + elp21[i][9]*d32*T2 + elp21[i][10]*d42*T2
        Ze += (elp21[i][12])*T2*sin(Y)

    elp22 = ELP200082B_ELP22
    for i in range(len(elp22)):
        Y = 0
        Y += radians(elp22[i][5]) + elp22[i][0]*z1*T1 + elp22[i][1]*d11*T1 + elp22[i][2]*d21*T1 + elp22[i][3]*d31*T1 + elp22[i][4]*d41*T1
        Y += elp22[i][0]*z2*T2 + elp22[i][1]*d12*T2 + elp22[i][2]*d22*T2 + elp22[i][3]*d32*T2 + elp22[i][4]*d42*T2
        Xe += elp22[i][6]*sin(Y)

    elp23 = ELP200082B_ELP23
    for i in range(len(elp23)):
        Y = 0
        Y += radians(elp23[i][5]) + elp23[i][0]*z1*T1 + elp23[i][1]*d11*T1 + elp23[i][2]*d21*T1 + elp23[i][3]*d31*T1 + elp23[i][4]*d41*T1
        Y += elp23[i][0]*z2*T2 + elp23[i][1]*d12*T2 + elp23[i][2]*d22*T2 + elp23[i][3]*d32*T2 + elp23[i][4]*d42*T2
        Ye += elp23[i][6]*sin(Y)

    elp24 = ELP200082B_ELP24
    for i in range(len(elp24)):
        Y = 0
        Y += radians(elp24[i][5]) + elp24[i][0]*z1*T1 + elp24[i][1]*d11*T1 + elp24[i][2]*d21*T1 + elp24[i][3]*d31*T1 + elp24[i][4]*d41*T1
        Y += elp24[i][0]*z2*T2 + elp24[i][1]*d12*T2 + elp24[i][2]*d22*T2 + elp24[i][3]*d32*T2 + elp24[i][4]*d42*T2
        Ze += elp24[i][6]*sin(Y)

    elp25 = ELP200082B_ELP25
    for i in range(len(elp25)):
        Y = 0
        Y += radians(elp25[i][5]) + elp25[i][0]*z1*T1 + elp25[i][1]*d11*T1 + elp25[i][2]*d21*T1 + elp25[i][3]*d31*T1 + elp25[i][4]*d41*T1
        Y += elp25[i][0]*z2*T2 + elp25[i][1]*d12*T2 + elp25[i][2]*d22*T2 + elp25[i][3]*d32*T2 + elp25[i][4]*d42*T2
        Xe += elp25[i][6]*T2*sin(Y)

    elp26 = ELP200082B_ELP26
    for i in range(len(elp26)):
        Y = 0
        Y += radians(elp26[i][5]) + elp26[i][0]*z1*T1 + elp26[i][1]*d11*T1 + elp26[i][2]*d21*T1 + elp26[i][3]*d31*T1 + elp26[i][4]*d41*T1
        Y += elp26[i][0]*z2*T2 + elp26[i][1]*d12*T2 + elp26[i][2]*d22*T2 + elp26[i][3]*d32*T2 + elp26[i][4]*d42*T2
        Ye += elp26[i][6]*T2*sin(Y)

    elp27 = ELP200082B_ELP27
    for i in range(len(elp27)):
        Y = 0
        Y += radians(elp27[i][5]) + elp27[i][0]*z1*T1 + elp27[i][1]*d11*T1 + elp27[i][2]*d21*T1 + elp27[i][3]*d31*T1 + elp27[i][4]*d41*T1
        Y += elp27[i][0]*z2*T2 + elp27[i][1]*d12*T2 + elp27[i][2]*d22*T2 + elp27[i][3]*d32*T2 + elp27[i][4]*d42*T2
        Ze += elp27[i][6]*T2*sin(Y)

    elp28 = ELP200082B_ELP28
    for i in range(len(elp28)):
        Y = 0
        Y += radians(elp28[i][5]) + elp28[i][0]*z1*T1 + elp28[i][1]*d11*T1 + elp28[i][2]*d21*T1 + elp28[i][3]*d31*T1 + elp28[i][4]*d41*T1
        Y += elp28[i][0]*z2*T2 + elp28[i][1]*d12*T2 + elp28[i][2]*d22*T2 + elp28[i][3]*d32*T2 + elp28[i][4]*d42*T2
        Xe += elp28[i][6]*sin(Y)

    elp29 = ELP200082B_ELP29
    for i in range(len(elp29)):
        Y = 0
        Y += radians(elp29[i][5]) + elp29[i][0]*z1*T1 + elp29[i][1]*d11*T1 + elp29[i][2]*d21*T1 + elp29[i][3]*d31*T1 + elp29[i][4]*d41*T1
        Y += elp29[i][0]*z2*T2 + elp29[i][1]*d12*T2 + elp29[i][2]*d22*T2 + elp29[i][3]*d32*T2 + elp29[i][4]*d42*T2
        Ye += elp29[i][6]*sin(Y)

    elp30 = ELP200082B_ELP30
    for i in range(len(elp30)):
        Y = 0
        Y += radians(elp30[i][5]) + elp30[i][0]*z1*T1 + elp30[i][1]*d11*T1 + elp30[i][2]*d21*T1 + elp30[i][3]*d31*T1 + elp30[i][4]*d41*T1
        Y += elp30[i][0]*z2*T2 + elp30[i][1]*d12*T2 + elp30[i][2]*d22*T2 + elp30[i][3]*d32*T2 + elp30[i][4]*d42*T2
        Ze += elp30[i][6]*sin(Y)

    elp31 = ELP200082B_ELP31
    for i in range(len(elp31)):
        Y = 0
        Y += radians(elp31[i][5]) + elp31[i][0]*z1*T1 + elp31[i][1]*d11*T1 + elp31[i][2]*d21*T1 + elp31[i][3]*d31*T1 + elp31[i][4]*d41*T1
        Y += elp31[i][0]*z2*T2 + elp31[i][1]*d12*T2 + elp31[i][2]*d22*T2 + elp31[i][3]*d32*T2 + elp31[i][4]*d42*T2
        Xe += elp31[i][6]*sin(Y)

    elp32 = ELP200082B_ELP32
    for i in range(len(elp32)):
        Y = 0
        Y += radians(elp32[i][5]) + elp32[i][0]*z1*T1 + elp32[i][1]*d11*T1 + elp32[i][2]*d21*T1 + elp32[i][3]*d31*T1 + elp32[i][4]*d41*T1
        Y += elp32[i][0]*z2*T2 + elp32[i][1]*d12*T2 + elp32[i][2]*d22*T2 + elp32[i][3]*d32*T2 + elp32[i][4]*d42*T2
        Ye += elp32[i][6]*sin(Y)
    
    elp33 = ELP200082B_ELP33
    for i in range(len(elp33)):
        Y = 0
        Y += radians(elp33[i][5]) + elp33[i][0]*z1*T1 + elp33[i][1]*d11*T1 + elp33[i][2]*d21*T1 + elp33[i][3]*d31*T1 + elp33[i][4]*d41*T1
        Y += elp33[i][0]*z2*T2 + elp33[i][1]*d12*T2 + elp33[i][2]*d22*T2 + elp33[i][3]*d32*T2 + elp33[i][4]*d42*T2
        Ze += elp33[i][6]*sin(Y)

    elp34 = ELP200082B_ELP34
    for i in range(len(elp34)):
        Y = 0
        Y += radians(elp34[i][5]) + elp34[i][0]*z1*T1 + elp34[i][1]*d11*T1 + elp34[i][2]*d21*T1 + elp34[i][3]*d31*T1 + elp34[i][4]*d41*T1
        Y += elp34[i][0]*z2*T2 + elp34[i][1]*d12*T2 + elp34[i][2]*d22*T2 + elp34[i][3]*d32*T2 + elp34[i][4]*d42*T2
        Xe += elp34[i][6]*T3*sin(Y)

    elp35 = ELP200082B_ELP35
    for i in range(len(elp35)):
        Y = 0
        Y += radians(elp35[i][5]) + elp35[i][0]*z1*T1 + elp35[i][1]*d11*T1 + elp35[i][2]*d21*T1 + elp35[i][3]*d31*T1 + elp35[i][4]*d41*T1
        Y += elp35[i][0]*z2*T2 + elp35[i][1]*d12*T2 + elp35[i][2]*d22*T2 + elp35[i][3]*d32*T2 + elp35[i][4]*d42*T2
        Ye += elp35[i][6]*T3*sin(Y)

    elp36 = ELP200082B_ELP36
    for i in range(len(elp36)):
        Y = 0
        Y += radians(elp36[i][5]) + elp36[i][0]*z1*T1 + elp36[i][1]*d11*T1 + elp36[i][2]*d21*T1 + elp36[i][3]*d31*T1 + elp36[i][4]*d41*T1
        Y += elp36[i][0]*z2*T2 + elp36[i][1]*d12*T2 + elp36[i][2]*d22*T2 + elp36[i][3]*d32*T2 + elp36[i][4]*d42*T2
        Ze += elp36[i][6]*sin(Y)


    Xe = Xe/rad + w[0][0] + w[0][1]*T2 + w[0][2]*T3 + w[0][3]*T4 + w[0][4]*T5
    Ye = Ye/rad
    Ze = Ze*a0/ath
    X1 = Ze*cos(Ye)
    X2 = X1*sin(Xe)
    X1 = X1*cos(Xe)
    X3 = Ze*sin(Ye)
    pw = (p1 + p2*T2 + p3*T3 + p4*T4 + p5*T5)*T2
    qw = (q1 + q2*T2 + q3*T3 + q4*T4 + q5*T5)*T2
    RA = 2*sqrt(1 - pw*pw - qw*qw)
    pwqw = 2*pw*qw
    pw2 = 1 - 2*pw*pw
    qw2 = 1 - 2*qw*qw
    pw = pw*RA
    qw = qw*RA
    Xe = pw2*X1 + pwqw*X2 + pw*X3
    Ye = pwqw*X1 + qw2*X2 - qw*X3
    Ze = -pw*X1 + qw*X2 + (pw2 + qw2 - 1)*X3

    return Xe, Ye, Ze





jam = 0
menit = 0
detik = 0
timezone = 0
waktu = jam/24 + menit/1440 + detik/86400


tahun = 1993
bulan = 1
tanggal = 13 + waktu
lintang = (6, 57, 0, "LS")
bujur = (107, 37, 0, "BT")
jde = julian_datum(tahun, bulan, tanggal)
x = ELP2000(jde)
print(x)




