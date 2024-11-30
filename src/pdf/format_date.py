def format_date(day,month,year):

    # takes the inputs dd, mm or mmm, yyyy and returns the date in 3 formats:
    # [ddth mmm yyyy, ddmmyy, mm/dd/yyyy]

    nn=["01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31"]            
    dayth = ["st","nd","rd","th","th","th","th","th","th","th","th","th","th","th","th","th","th","th","th","th","st","nd","rd","th","th","th","th","th","th","th","th","st"]
    months_short=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    months_long=["January","February","March","April","May","June","July","August","September","October","November","December"]

    try:
        if isinstance(day,str):
            try:
                day = int(day)
            except:
                print("problem converting day string to integer")

        if isinstance(day,int):
            if day>=1 and day<=31:
                day_str = str(day)+dayth[day-1]
                dd_str=nn[day-1]
            else:
                print ("error with day range")
        
        if isinstance(month,str):
            try:
                month = int(month)
            except:
                #note that month is not a string of an integer
                pass

        if isinstance(month,int):
            if month>=1 and month<=12:
                month_str = months_long[month-1]
                mm_str=nn[month-1]
            else:
                print ("error month is out of range")

        if isinstance(month,str):
            if month.capitalize() in months_short:
                month_str = months_long[months_short.index(month.capitalize())]
                mm_str=nn[months_short.index(month.capitalize())]
            elif month.capitalize() in months_long:
                month_str = month.capitalize()
                mm_str=nn[months_long.index(month.capitalize())]
                pass
            else:
                print ("error: month string not recognised")

        if isinstance(year,int):
            year_str = str(year)
        else:
            year_str = year
    
    except Exception as e:
        print("Problem generating some part of date string:",e)

    try:
        long_date_str = day_str+" "+month_str+" "+year_str
        short_date_str= dd_str+mm_str+year_str[-2:]
        excel_date_str= mm_str+"/"+dd_str+"/"+year_str
    except Exception as e:
        print ("problem forming date string:",e)
        long_date_str="<< DATE ERROR >>"
        short_date_str="<< DATE ERROR >>"
        excel_date_str = "<< DATE ERROR >>"

    return ([long_date_str,short_date_str,excel_date_str])
