import openpyxl as opx
import datetime

def num_years(begin_date, end_date):

    years = end_date.year - begin_date.year
    if end_date.month < begin_date.month or (end_date.month == begin_date.month and end_date.day < begin_date.day):
        years -= 1
    return years


def num_tagen(begin_date, end_date):

    difference = end_date - begin_date.replace(end_date.year)
    return difference


def berechnunf(start_datum, end_datum,kapital,zins):
    num_year = num_years(start_datum,end_datum)
    num_tage = num_tagen(start_datum,end_datum).days
    return kapital*(1+zins)**num_year + kapital*((zins/365.25))*num_tage


#get current time
end_date =datetime.datetime(2020,1,10)         # Input Date


# read source data
workbook = opx.load_workbook("TestData.xlsx")   #Input source data
sheet = workbook.worksheets[0]

# worksheet => store data in 2D list
einzahlungen = [[] for i in range(sheet.max_row)]
i = 0
for row in sheet:
    einzahlungen[i].append(row[0].value)
    einzahlungen[i].append(row[1].value)
    i = i +1

#initial
min_zin = -1.0
max_zin = 1.0
zins_a = 0.0
zins_b = 0.0
sum_a = 0
sum_b = 0
end_sum = 100000    #Input total amount

#Main
z= 0
while abs(end_sum-sum_a) >=0.1:
    sum_a = 0
    for row in range(len(einzahlungen)):
        sum_a = sum_a +berechnunf(einzahlungen[row][0],end_date,einzahlungen[row][1],zins_a)
    zins_b = zins_a
    if sum_a < end_sum:
            zins_a =((max_zin-zins_a)/2) + zins_a
            min_zin = zins_b
    if sum_a > end_sum:
            zins_a = ((zins_a-min_zin)/2) + min_zin
            max_zin = zins_b
    z = z +1
print(z)
print("Bis " + str(end_date))
print("Zins ist: " + str(zins_a))
