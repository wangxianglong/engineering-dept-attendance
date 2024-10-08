
def cal_remaining_hours(index,last_remaining_hours,rest_hours) -> tuple:
    if index <= 3:
        if last_remaining_hours[index] != "":
            if float(last_remaining_hours[index]) >= rest_hours:
                last_remaining_hours[index] = float(last_remaining_hours[index]) - rest_hours
            else:
                rest_hours = rest_hours - float(last_remaining_hours[index])
                last_remaining_hours[index] = 0
                cal_remaining_hours(index + 1,last_remaining_hours,rest_hours)
        else:
            cal_remaining_hours(index + 1,last_remaining_hours,rest_hours)

    return (last_remaining_hours,rest_hours)

rest_hours = 30
result = cal_remaining_hours(0,['26.5', '5', '14', ''],rest_hours)
print(result)
print(rest_hours)
