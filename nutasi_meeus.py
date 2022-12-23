def u(t_td):
    return t_td/100


def epsilon_zero(u):
    epsilon_0 = 23 + 26/60 + 21.448/3600 
    epsilon_0 += (-4680.93*u - 1.55*u**2 + 1999.25*u**3 - 51.38*u**4 - 249.67*u**5 - 39.05*u**6+ 7.12*u**7 + 27.87*u**8 + 5.79*u**9 + 2.45*u**10)/3600
    return epsilon_0

def delta_epsilon():
    pass


def epsilon(epsilon_zero, delta_epsilon):
    return epsilon_zero + delta_epsilon