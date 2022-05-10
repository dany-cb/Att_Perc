from email.policy import default
from numpy import roll
import pandas as pd
import csv
from os.path import exists


def xl_csv(xl_file):

    read_file = pd.read_excel(xl_file)
    read_file.to_csv("as.csv", index=None, header=False)


def perc(roll_no=""):
    file = open("as.csv")
    reader = csv.reader(file)
    total_hrs = present = 0
    start = False

    if roll_no:
        for row in reader:
            if row[3] == roll_no:
                for i in row[5:-3]:
                    if i and int(float(i)) in [0, 1]:
                        total_hrs += 1
                        present += float(i)
        file.close()
        print("Present: {}\nTotal hours: {}".format(present, total_hrs))
        return [present / total_hrs * 100, present, total_hrs]

    else:
        attendance = dict()
        for row in reader:
            if row[3] == "311120205001":
                start = True
            elif row[3] == "":
                start = False

            if start:
                hrs = on = 0
                for i in row[5:-3]:
                    if i and int(float(i)) in [0, 1]:
                        hrs += 1
                        on += int(float(i))
                attendance[row[4]] = add(attendance.get(row[4], [0, 0]), [on, hrs])
        file.close()
        return attendance


def add(l1, l2):
    temp = []
    for i in range(len(l1)):
        temp.append(l1[i] + l2[i])
    return temp


def correction(percent):
    p = int(input("Additional Present: ")) + percent[1]
    h = int(input("Additional Hrs: ")) + percent[2]
    return [p / h * 100, p, h]


if __name__ == "__main__":
    if not exists("as.csv"):
        xl_csv("Attendance.xlsx")

    percent = perc("311120205020")
    print("Total Percentage = %.2f" % percent[0], "%", sep="")
    yes = input("Any changes?")
    if yes and (yes == "Y" or yes == "y"):
        percent = correction(percent)
    need = 0
    while percent[0] < 75:
        need += 1
        percent[0] = (percent[1] + need) / (percent[2] + need) * 100
    print("Need to attend {} classes for {:.2f}%.".format(need, percent[0]))

    # atdnce_dict = perc()
    # for i in atdnce_dict.keys():
    #     print("{:<30}{:6.2f}%".format(i, atdnce_dict[i][0] / atdnce_dict[i][1] * 100))
