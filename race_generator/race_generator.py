import openpyxl, os, datetime, math
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook
from win32com.client import Dispatch

def count_time(start):
    time_sec = start
    time_hour, time_sec =  time_sec // 3600, time_sec % 3600
    time_min, time_sec = time_sec // 60, time_sec % 60
    time_hour = str(int(time_hour))
    time_min = str(int(time_min))
    time_sec = str(round(time_sec, 2))
    time_sec = time_sec.split('.')
    time_mil = time_sec[1]
    time_sec = time_sec[0]
    count_up_str = "{}:{}:{}.{}".format(time_hour.zfill(2),
                                                    time_min.zfill(2),
                                                    time_sec.zfill(2),
                                                    time_mil.zfill(2))
    return count_up_str

def secs2time(s):
    ms = int((s - int(s)) * 1000000)
    s = int(s)
    # Get rid of this line if s will never exceed 86400
    #while s >= 24*60*60:
        #s -= 24*60*60
    h = s / (60*60)
    s -= h*60*60
    m = s / 60
    s -= m*60
    return datetime.time(h, m, s, ms)

# function to get the required time for each distance
def time_calc(dist, tspeed):
    hour_in_seconds = 3600.0
    return round((dist / tspeed) * hour_in_seconds, 1)

def get_time_list(total, speed):
    current_dist = 0.0
    time_list = []

    #while (current_dist < total - .1):
    for i in range(0, total):
        current_dist += .1
        #time_list.append(str(secs2time(time_calc(current_dist, speed))))
        time_list.append(count_time(time_calc(current_dist, speed)))
    
    return time_list

def create_folder(path):
    try:
        os.mkdir(path)
    except OSError:
        pass
    else:
        pass 

def save_list_in_excel(pulse_list, target_list, mile_list, tire_size, target_miles, target_speed, number_of_legs, end_1st_leg=0, end_2nd_leg=0):
    # Create the spreadsheet 
    # filename = r"C:\Users\AIOTIK-005\Desktop\Route_created.xlsx"
    book = Workbook()
    ws = book.active

    # number of legs
    ws['D1'] = "Number of Legs"
    ws['D2'] = int(number_of_legs)
    ws['E1'] = "End of 1st Leg (feet)"
    ws['E2'] = int(end_1st_leg)
    ws['F1'] = "End of 2st Leg (feet)"
    ws['F2'] = int(end_2nd_leg)


    # Reference
    ws['H1'] = "REF"

    # write the target data here
    ws['I1'] = target_speed
    ws['J1'] = "MPH"

    ws['K1'] = tire_size
    ws['L1'] = "Inch"
    
    ws['M1'] = target_miles
    ws['N1'] = "Miles"

    # setup colum A and B width
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 22
    ws.column_dimensions['F'].width = 22
    # set sheet name 
    ws.append(["PULSES", "TIME", "MILES"])

    # append the values to the column 
    # append the mile_list to the spreadsheet 
    r = 3
    for statN in mile_list:
        ws.cell(row = r, column = 3).value = statN
        r += 1
    
    # append the pulse list to the spreadsheet
    r = 3
    for statN in pulse_list:
        # ws.cell(row = r, column = 1).value = '=(63360/K1/2) * {}'.format(str(r - 2))
        ws.cell(row = r, column = 1).value = '{}'.format(statN)
        ws.cell(row = r, column = 1).number_format = '0'
        #ws.cell(row = r, column = 1).value = statN
        r += 1
    
    # append the target list to spreadsheet
    r = 3
    for statN in target_list:
        ws.cell(row = r, column = 2).value = '=((C{} / I1) * 3600.0) / 86400'.format(r)
        ws.cell(row = r, column = 2).number_format = 'hh:mm:ss.00'
        r += 1

    # create the spreadsheet 
    user = os.getcwd().split('\\')
    user = user[0] + '\\' + user[1] + '\\' + user[2]
    filename = str(input("Enter the name for the spreadsheet: "))
    filepath = user + "\Desktop\\" + filename + r'.xlsx'
    # filepath = r"C:\Users\\" + getpass.getuser() +  r"\Desktop\\" + filename + r".xlsx"
    folder_name = 'race_data'
    create_folder(folder_name)
    path = os.getcwd() + '\\' + folder_name + '\\' + filename + r'.xlsx'
    print(f"Your file location is here: {path}")
    book.save(path) # save the spreadsheet generated
    just_open(path) # open the spreadsheet, save and generate the formula data for race

def just_open(path):
    """
                Simulate opening the excel sheet manually
                :param strFileName: opened excel file name (suffix.xlsx format)
    """
    try:
        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = False
        xlBook = xlApp.Workbooks.Open(path)
        xlBook.Save()
        xlBook.Close()
    except:
        print("Please open %s manually, modify a blank value and save it at will" % path)

def calculate_pulses(tire_size):
    return (63360 / tire_size) / 2         # rotations per mile = (mile in inches / tire circumference) 
                                           # 6336 = 0.1 miles
                                           # 528 feet = 6336 inches

def get_pulse_by_distance(first_pulse, distance_end_1st_leg):
    distance_list = []
    mile_counter = 0
    while (mile_counter < distance_end_1st_leg):
        mile_counter += .1
        mile_counter = round(mile_counter, 1)
        distance_list.append(mile_counter)

    dist_result = round(first_pulse * len(distance_list), 0)
    return dist_result

def get_pulses_list(first_pulse, total):
    multiplier = 0
    pulse_list = []

    for i in range(0, total):
        multiplier += 1
        result = round(first_pulse * multiplier, 0)
        pulse_list.append(result)
    
    return pulse_list

def get_miles_list(total):
    mile_list = []
    mile_counter = 0

    while (mile_counter < total):
        mile_counter += .1
        mile_counter = round(mile_counter, 1)
        mile_list.append(mile_counter)
    
    return mile_list

def convert_miles_to_feet(miles):
    return (miles - .1) * 5280

def convert_feet_to_inch(feet):
    return feet * 12

def calculate_tire_rotation_from_feet(inches, tire_size):
    return round(inches / tire_size, 2)

def calculate_total_pulses(tire_rotation):
    return tire_rotation * 5

def calculate_minimum_pulse(total_pulses, distance_size):
    return (total_pulses / distance_size)

def calculate_list_pulses(total_pulses, distance_size, more_pulses=0):
    pulse_list = []
    for index in range(distance_size + 1):
        conversion = int((total_pulses * index) + more_pulses)
        pulse_list.append(conversion)
    
    return pulse_list

def get_total_pulses(distance, tire_size, total_miles, add_more_pulses=0):
    inches_distance = convert_feet_to_inch(distance)
    tire_rotation = calculate_tire_rotation_from_feet(inches_distance, tire_size)
    total_pulses = int(calculate_total_pulses(tire_rotation))
    return total_pulses 

def get_mile_from_feet_times_ten(feet):
    value =  feet / 5280
    val_rounded = round(value)
    result = abs(value - val_rounded)
    if result >= 0.1:
        return round( (feet / 5280 ) * 10)
    else:
        return val_rounded * 10



def get_pulses_in_leg(distance, tire_size, leg_length, add_more_pulses=0, more_pulses=0):
    # 1. get the total pulses for the end of leg 
    total_pulses = round(((distance * 12) / tire_size) * 5)
    min_pulse = ((total_pulses) / leg_length) 
    pulses = []
    if add_more_pulses == 0:
        for item in range(1, leg_length):
            pulse = ((min_pulse) * (item))
            pulses.append(pulse)
        # append the last 1st leg 
        pulses.append(total_pulses)
    else:
        n = add_more_pulses
        limit = n + leg_length
        while n < limit:
            n += 1
            pulse = (min_pulse * (n - add_more_pulses)  + more_pulses)
            pulses.append(pulse)

    return pulses

# main application
def run(): 
    try:
        # a warning message 
        print("////////////////////////////////////////////////////")
        print("PLEASE, CLOSE ANY EXCEL FILE RELATED TO THIS PROGRAM")
        print("////////////////////////////////////////////////////\n")
        
        # in miles get the user data input 
        tire_size = float(input("Input the tire size (inches): "))
        total_length = float(input("Input total race length (miles): "))
        my_speed = float(input("Input target speed (mph): "))
        number_of_legs = int(input("Input the number of legs (1 or 2) depending of the type of race: "))
        
        
        # get mile and time list generated 
        mile_list = get_miles_list(total_length)
        timer_list = get_time_list(len(mile_list), my_speed)

        
        if number_of_legs == 1:
            distance_1st_leg = float(input("Input the end of 1st leg (feet): "))
            total_miles_1st_leg = get_mile_from_feet_times_ten(distance_1st_leg) # convert to miles and times 10 for the spreadsheet
            pulse_list = get_pulses_in_leg(distance_1st_leg, tire_size, total_miles_1st_leg)

            # save the data into an spreadsheet
            save_list_in_excel(pulse_list, timer_list, mile_list, tire_size, total_length, my_speed, number_of_legs)
            print("\n////////////////////////////////////////////////////")
            print("SPREADSHEET CREATED, NOW THIS WINDOW IS ABLE TO BE CLOSED")
            print("////////////////////////////////////////////////////\n")
            input("Enter any key to exit")
        elif number_of_legs == 2:
            
            distance_1st_leg = float(input("Input the end of 1st leg (feet): "))
            distance_2nd_leg = float(input("Input the end of 2st leg (feet): "))

            total_miles_1st_leg = get_mile_from_feet_times_ten(distance_1st_leg) # convert to miles and times 10 for the spreadsheet
            total_miles_2nd_leg = get_mile_from_feet_times_ten(distance_2nd_leg) # convert to miles and times 10 for the spreadsheet
            print(f"total miles 1 {total_miles_1st_leg} total miles 2 {total_miles_2nd_leg}")
            pulse_list = get_pulses_in_leg(distance_1st_leg, tire_size, total_miles_1st_leg)
            total_pulses = get_total_pulses(distance_1st_leg, tire_size, total_miles_1st_leg)

            pulses_2nd_leg = get_pulses_in_leg(distance_2nd_leg, tire_size, total_miles_2nd_leg, total_miles_1st_leg, total_pulses)
            pulse_list.extend(pulses_2nd_leg)
            # remove the float point 
            pulse_list = ([int(p) for p in pulse_list])
            # save the data into an spreadsheet
            save_list_in_excel(pulse_list, timer_list, mile_list, tire_size, total_length, my_speed, number_of_legs, distance_1st_leg, distance_2nd_leg)
            print("\n////////////////////////////////////////////////////")
            print("SPREADSHEET CREATED, NOW THIS WINDOW IS ABLE TO BE CLOSED")
            print("////////////////////////////////////////////////////\n")
            input("Enter any key to exit")

        else:
            print("please Input a number of legs valid: (1 o 2)")
            exit()
    except PermissionError:
        print("There's a spreadsheet with the same name in the same location")
        input("choose another name and try again. \nPress any key to terminate this program.")


