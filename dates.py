import csv
from datetime import datetime, date, timedelta

f = open('dates.csv', 'w')
writer = csv.writer(f)

startDate = datetime(2022, 1, 3)
endDate = datetime(2022, 9, 5)

currentDay = startDate


while currentDay < endDate:
    weekDates = []
    for x in range(7):
        weekDates.append(currentDay.strftime("%d/%m/%Y").lstrip("0"))
        currentDay += timedelta(days=1)
    writer.writerow(weekDates)
f.close()
