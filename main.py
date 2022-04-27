import openpyxl
import matplotlib.pyplot as plt

# open and read exel file
book = openpyxl.open('testfile.xlsx', read_only=True)
sheet = book.active
cells1 = sheet['A3':'B9']
cells2 = sheet['A12': 'B18']
cells3 = sheet['A21': 'B27']


def show_diagram(alpha, fx):
    plt.plot(alpha, fx, 'o-r')
    plt.title('F(x)')
    plt.grid(True)
    plt.show()


# I use the linear interpolation formula
# to find the value of a function from a given alpha
def interpolate(cells):
    list_of_alpha = []
    list_of_Fx = []
    for alpha, Fx in cells:
        list_of_alpha.append(alpha.value)
        list_of_Fx.append(Fx.value)
    list_of_Fx = list_of_Fx[::-1]
    list_of_alpha = list_of_alpha[::-1]
    alpha_to_interpolate = float(input('Enter alpha: '))
    if alpha_to_interpolate not in list_of_alpha:
        list_of_alpha.append(alpha_to_interpolate)
        list_of_alpha = sorted(list_of_alpha)
        alpha_index = list_of_alpha.index(alpha_to_interpolate)

        A = (list_of_Fx[alpha_index] - list_of_Fx[alpha_index - 1]) / \
            (list_of_alpha[alpha_index + 1] - list_of_alpha[alpha_index - 1])
        B = list_of_Fx[alpha_index - 1] - A * list_of_alpha[alpha_index - 1]
        resultFx = round(A * alpha_to_interpolate + B, 2)
        list_of_Fx.insert(alpha_index, resultFx)
    show_diagram(alpha=list_of_alpha, fx=list_of_Fx)


interpolate(cells=cells1)
interpolate(cells=cells2)
interpolate(cells=cells3)
