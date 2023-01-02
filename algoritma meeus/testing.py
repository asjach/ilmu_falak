


def cek_argumen(x):
    if isinstance(x, str):
        print('str')
    elif isinstance(x, int):
        print('int')
    elif isinstance(x, tuple):
        print('tuple')


x = (10,10,11, "LS")
cek_argumen(x)