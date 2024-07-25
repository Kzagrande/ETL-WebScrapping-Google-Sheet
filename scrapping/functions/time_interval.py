from datetime import datetime, timedelta



def get_current_and_last_hour():
    
    #hora atual tirando um milisegundo
    now = datetime.now()
    now = now.replace(minute=0, second=0, microsecond=0)
    one_second = timedelta(seconds=1)
    last_second = now - one_second
    format_last_second = last_second.strftime("%Y-%m-%d %H:%M:%S")

    #hora passada 
    one_hour = timedelta(hours=1)
    last_hour = now - one_hour
    format_last_hour = last_hour.strftime("%Y-%m-%d %H:%M:%S")


    result = {
        'last_second': format_last_second,
        'last_hour': format_last_hour
    }

    print('last_second', format_last_second)
    print( 'last_hour', format_last_hour)
    return result

get_current_and_last_hour()

