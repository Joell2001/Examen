def caudal(Qmax, Pwf, Pr):

    Qo = Qmax * (1 - 0.2 * (Pwf / Pr) - 0.8 * (Pwf / Pr) ** 2)
    return Qo

