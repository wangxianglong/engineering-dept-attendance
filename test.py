
def cal_remaining_hours(index,last_remaining_hours) -> list:
    if index <= 3:
        if last_remaining_hours[index] != "":
            if float(last_remaining_hours[index]) >= last_remaining_hours[4]:
                last_remaining_hours[index] = float(last_remaining_hours[index]) - last_remaining_hours[4]
                last_remaining_hours[4] = 0
            else:
                last_remaining_hours[4] = last_remaining_hours[4] - float(last_remaining_hours[index])
                last_remaining_hours[index] = 0
                cal_remaining_hours(index + 1,last_remaining_hours)
             
        else:
            cal_remaining_hours(index + 1,last_remaining_hours)

    return last_remaining_hours

start_index = 3
for i in range(1,9):
 
    result = list()
    for j in range(0,4):
        result.append(j + start_index)
    print(result,sep="\n")
    start_index = start_index + 4

# rest_hours = 300
# result = cal_remaining_hours(0,['26.5', '', '', '5',300])
# print(result)

a = float("26.5")
b = float("8")
print(a + b)
