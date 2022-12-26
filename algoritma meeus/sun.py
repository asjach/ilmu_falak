from math import radians, sin, cos, tan, asin, atan, atan2, degrees
from nutation import *
from moon import *

class Sun:
    def __init__(self, t_td, lst_nampak, latitude):
        self.t              = t_td
        self.tau            = t_td/10
        self.nutasi         = Nutation(t_td)
        self.delta_psi      = self.nutasi.delta_psi
        self.epsilon        = self.nutasi.epsilon
        data_bulan          = Moon(self.t,0,0,0,0)
        self.T              = data_bulan.koreksi_bujur_bulan
        self.lst_nampak     = lst_nampak
        self.latitude       = radians(latitude)


    @property
    def theta_terkoreksi(self):
        return (self.theta() + self.delta_theta()) % 360

    
    @property
    def koreksi_aberasi(self):
        return -20.4898/(3600*self.jarak_bumi_matahari)


    @property
    def bujur_matahari_nampak(self):
        return radians((self.theta_terkoreksi + self.delta_psi + self.koreksi_aberasi) % 360)


    @property
    def lintang_matahari_tampak(self):
        tau = self.tau
        B0 = list_B0
        B1 = list_B1
        totalB0 = 0
        totalB1 = 0

        for i in range(len(B0)):
            totalB0 += B0[i][0] * cos(B0[i][1] + B0[i][2] * tau)

        for i in range(len(B1)):
            totalB1 += B1[i][0] * cos(B1[i][1] + B1[i][2] * tau)
        totalB1 *= tau
        B = degrees((totalB0 + totalB1)/100000000)*3600
        beta = -B
        delta_beta = 0.03916 * (cos(self.lambda_aksen())-sin(self.lambda_aksen()))
        beta_terkoreksi = beta + delta_beta
        return radians(beta_terkoreksi/3600)


    @property
    def jarak_bumi_matahari(self):
        tau = self.tau
        R0 = list_R0
        R1 = list_R1
        R2 = list_R2
        R3 = list_R3
        R4 = list_R4

        totalR0 = 0
        totalR1 = 0
        totalR2 = 0
        totalR3 = 0
        totalR4 = 0
        for i in range(len(R0)):
            totalR0 += R0[i][0] * cos(R0[i][1] + R0[i][2] * tau)

        for i in range(len(R1)):
            totalR1 += R1[i][0] * cos(R1[i][1] + R1[i][2] * tau)

        for i in range(len(R2)):
            totalR2 += R2[i][0] * cos(R2[i][1] + R2[i][2] * tau)

        for i in range(len(R3)):
            totalR3 += R3[i][0] * cos(R3[i][1] + R3[i][2] * tau)

        for i in range(len(R4)):
            totalR4 += R4[i][0] * cos(R4[i][1] + R4[i][2] * tau)

        return (totalR0 + totalR1 * tau + totalR2*tau**2 + totalR3*tau**3 + totalR4*tau**4)/100000000

    
    @property
    def sudut_jari_jari_matahari(self):
        return (959.63/3600)/self.jarak_bumi_matahari


    @property
    def alpha_matahari(self):
        y = sin((self.bujur_matahari_nampak))*cos(self.epsilon) -tan(self.lintang_matahari_tampak)*sin(self.epsilon)
        x = cos((self.bujur_matahari_nampak))
        return (degrees(atan2(y,x)) % 360)


    @property
    def delta_matahari(self):
        return asin(sin(self.lintang_matahari_tampak)*cos(self.epsilon) + cos(self.lintang_matahari_tampak)*sin(self.epsilon)*sin(self.bujur_matahari_nampak))

    @property
    def hour_angle_matahari(self):
        return radians(self.lst_nampak*15-self.alpha_matahari)


    @property
    def azimut_selatan(self):
        y = sin(self.hour_angle_matahari)
        x = cos(self.hour_angle_matahari)*sin(self.latitude) - tan(self.delta_matahari)*cos(self.latitude)
        return degrees((atan2(y,x)))

    @property
    def azimut_matahari(self):
        return (self.azimut_selatan + 180) % 360


    @property
    def altitude(self):
        return degrees(asin(sin(self.latitude)*sin(self.delta_matahari)+cos(self.latitude)*cos(self.delta_matahari)*cos(self.hour_angle_matahari)))


    @property
    def sudut_parallaks_matahari(self):
        return degrees(atan(6378.14/(self.jarak_bumi_matahari*149598000)))





    def theta(self):
        return (degrees(self.L()) + 180) % 360


    def delta_theta(self):
        return -0.09033/3600


    def L(self):
        return (self.L0() + self.L1() + self.L2() + self.L3() + self.L4() + self.L5())/100000000

    #@property
    def lambda_aksen(self):
        return radians((self.theta() - 1.397*self.T - 0.00031*self.T*self.T) % 360)



    def L0(self):
        tau = self.tau
        l0 = list_L0
        total = 0

        for i in range(len(l0)):
            total += l0[i][0] * cos(l0[i][1] + l0[i][2] * tau)
        return total


    def L1(self):
        tau = self.tau
        l1 = list_L1
        total = 0

        for i in range(len(l1)):
            total += l1[i][0] * cos(l1[i][1] + l1[i][2] * tau)
        return total * tau


    def L2(self):
        tau = self.tau
        l2 = list_L2
        total = 0

        for i in range(len(l2)):
            total += l2[i][0] * cos(l2[i][1] + l2[i][2] * tau)
        return total * tau** 2


    def L3(self):
        tau = self.tau
        l3 = list_L3
        total = 0

        for i in range(len(l3)):
            total += l3[i][0] * cos(l3[i][1] + l3[i][2] * tau)
        return total * tau**3


    def L4(self):
        tau = self.tau
        l4 = list_L4
        total = 0

        for i in range(len(l4)):
            total += l4[i][0] * cos(l4[i][1] + l4[i][2] * tau)
        return total * tau**4


    def L5(self):
        tau = self.tau
        l5 = list_L5
        total = 0

        for i in range(len(l5)):
            total += l5[i][0] * cos(l5[i][1] + l5[i][2] * tau)
        return total*tau**5



list_L0 = [
[175347046, 0, 0],
[3341656, 4.6692568, 6283.07585],
[34894, 4.6261, 12566.1517],
[3497, 2.7441, 5753.3849],
[3418, 2.8289, 3.5231],
[3136, 3.6277, 77713.7715],
[2676, 4.4181, 7860.4194],
[2343, 6.1352, 3930.2097],
[1324, 0.7425, 11506.7698],
[1273, 2.0371, 529.691],
[1199, 1.1096, 1577.3435],
[990, 5.233, 5884.927],
[902, 2.045, 26.298],
[857, 3.508, 398.149],
[780, 1.179, 5223.694],
[753, 2.533, 5507.553],
[505, 4.583, 18849.228],
[492, 4.205, 775.523],
[357, 2.92, 0.067],
[317, 5.849, 11790.629],
[284, 1.899, 796.298],
[271, 0.315, 10977.079],
[243, 0.345, 5486.778],
[206, 4.806, 2544.314],
[205, 1.869, 5573.143],
[202, 2.458, 6069.777],
[156, 0.833, 213.299],
[132, 3.411, 2942.463],
[126, 1.083, 20.775],
[115, 0.645, 0.98],
[103, 0.636, 4694.003],
[102, 0.976, 15720.839],
[102, 4.267, 7.114],
[99, 6.21, 2146.17],
[98, 0.68, 155.42],
[86, 5.98, 161000.69],
[85, 1.3, 6275.96],
[85, 3.67, 71430.7],
[80, 1.81, 17260.15],
[79, 3.04, 12036.46],
[75, 1.76, 5088.63],
[74, 3.5, 3154.69],
[74, 4.68, 801.82],
[70, 0.83, 9437.76],
[62, 3.98, 8827.39],
[61, 1.82, 7084.9],
[57, 2.78, 6286.6],
[56, 4.39, 14143.5],
[56, 3.47, 6279.55],
[52, 0.19, 12139.55],
[52, 1.33, 1748.02],
[51, 0.28, 5856.48],
[49, 0.49, 1194.45],
[41, 5.37, 8429.24],
[41, 2.4, 19651.05],
[39, 6.17, 10447.39],
[37, 6.04, 10213.29],
[37, 2.57, 1059.38],
[36, 1.71, 2352.87],
[36, 1.78, 6812.77],
[33, 0.59, 17789.85],
[30, 0.44, 83996.85],
[30, 2.74, 1349.87],
[25, 3.16, 4690.48]
]

list_L1 = [
[628331966747, 0, 0],
[206059, 2.678235, 6283.07585],
[4303, 2.6351, 12566.1517],
[425, 1.59, 3.523],
[119, 5.796, 26.298],
[109, 2.966, 1577.344],
[93, 2.59, 18849.23],
[72, 1.14, 529.69],
[68, 1.87, 398.15],
[67, 4.41, 5507.55],
[59, 2.89, 5223.69],
[56, 2.17, 155.42],
[45, 0.4, 796.3],
[36, 0.47, 775.52],
[29, 2.65, 7.11],
[21, 5.34, 0.98],
[19, 1.85, 5486.78],
[19, 4.97, 213.3],
[17, 2.99, 6275.96],
[16, 0.03, 2544.31],
[16, 1.43, 2146.17],
[15, 1.21, 10977.08],
[12, 2.83, 1748.02],
[12, 3.26, 5088.63],
[12, 5.27, 1194.45],
[12, 2.08, 4694],
[11, 0.77, 553.57],
[10, 1.3, 6286.6],
[10, 4.24, 1349.87],
[9, 2.7, 242.73],
[9, 5.64, 951.72],
[8, 5.3, 2352.87],
[6, 2.65, 9437.76],
[6, 4.67, 4690.48]
]


list_L2 = [
[52919, 0, 0],
[8720, 1.0721, 6283.0758],
[309, 0.867, 12566.152],
[27, 0.05, 3.52],
[16, 5.19, 26.3],
[16, 3.68, 155.42],
[10, 0.76, 18849.23],
[9, 2.06, 77713.77],
[7, 0.83, 775.52],
[5, 4.66, 1577.34],
[4, 1.03, 7.11],
[4, 3.44, 5573.14],
[3, 5.14, 796.3],
[3, 6.05, 5507.55],
[3, 1.19, 242.73],
[3, 6.12, 529.69],
[3, 0.31, 398.15],
[3, 2.28, 553.57],
[2, 4.38, 5223.69],
[2, 3.75, 0.98]
]

list_L3 = [
[289, 5.844, 6283.076],
[35, 0, 0],
[17, 5.49, 12566.15],
[3, 5.2, 155.42],
[1, 4.72, 3.52],
[1, 5.3, 18849.23],
[1, 5.97, 242.73]
]

list_L4 = [
[114, 3.142, 0],
[8, 4.13, 6283.08],
[1, 3.84, 12566.15]
]

list_L5 = [
[1, 3.14, 0]
]



list_R0 = [
[100013989, 0, 0],
[1670700, 3.0984635, 6283.07585],
[13956, 3.05525, 12566.1517],
[3084, 5.1985, 77713.7715],
[1628, 1.1739, 5753.3849],
[1576, 2.8469, 7860.4194],
[925, 5.453, 11506.77],
[542, 4.564, 3930.21],
[472, 3.661, 5884.927],
[346, 0.964, 5507.553],
[329, 5.9, 5223.694],
[307, 0.299, 5573.143],
[243, 4.273, 11790.629],
[212, 5.847, 1577.344],
[186, 5.022, 10977.079],
[175, 3.012, 18849.228],
[110, 5.055, 5486.778],
[98, 0.89, 6069.78],
[86, 5.69, 15720.84],
[86, 1.27, 161000.69],
[65, 0.27, 17260.15],
[63, 0.92, 529.69],
[57, 2.01, 83996.85],
[56, 5.24, 71430.7],
[49, 3.25, 2544.31],
[47, 2.58, 775.52],
[45, 5.54, 9437.76],
[43, 6.01, 6275.96],
[39, 5.36, 4694],
[38, 2.39, 8827.39],
[37, 0.83, 19651.05],
[37, 4.9, 12139.55],
[36, 1.67, 12036.46],
[35, 1.84, 2942.46],
[33, 0.24, 7084.9],
[32, 0.18, 5088.63],
[32, 1.78, 398.15],
[28, 1.21, 6286.6],
[28, 1.9, 6279.55],
[26, 4.59, 10447.39]
]


list_R1 = [
[103019, 1.10749, 6283.07585],
[1721, 1.0644, 12566.1517],
[702, 3.142, 0],
[32, 1.02, 18849.23],
[31, 2.84, 5507.55],
[25, 1.32, 5223.69],
[18, 1.42, 1577.34],
[10, 5.91, 10977.08],
[9, 1.42, 6275.96],
[9, 0.27, 5486.78]
]


list_R2 = [
[4359, 5.7846, 6283.0758],
[124, 5.579, 12566.152],
[12, 3.14, 0],
[9, 3.63, 77713.77],
[6, 1.87, 5573.14],
[3, 5.47, 18849.23]
]


list_R3 = [
[145, 4.273, 6283.076],
[7, 3.92, 12566.15]
]


list_R4 = [
[4, 2.56, 6283.08]
]


list_B0 = [
[280, 3.199, 84334.662],
[102, 5.422, 5507.553],
[80, 3.88, 5223.69],
[44, 3.7, 2352.87],
[32, 4, 1577.34]
]

list_B1 = [
[9, 3.9, 5507.55],
[6, 1.73, 5223.69]

]