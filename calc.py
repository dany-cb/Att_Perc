import pandas as pd
import csv
from os.path import exists


def xl_csv(xl_file):

    read_file = pd.read_excel(xl_file)
    read_file.to_csv("as.csv", index=None, header=False)


def perc(roll_no):
    file = open("as.csv")
    reader = csv.reader(file)
    total_hrs = present = 0

    for row in reader:
        if row[3] == roll_no:
            for i in row[5:-3]:
                if i and int(float(i)) in [0, 1]:
                    total_hrs += 1
                    present += float(i)
    file.close()
    print(f"Present: {present}\nTotal: {total_hrs}")
    return present / total_hrs


if __name__ == "__main__":
    if not exists("as.csv"):
        xl_csv("Attendance.xlsx")
    percent = perc("311120205038") * 100
    print("Total Percentage = %.2f" % percent, "%", sep="")
