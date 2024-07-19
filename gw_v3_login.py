from gw_v3_set_up import driver,data,Keys,EC,By,WebDriverWait,json
from gw_v3_functions import commons, time

def Login(domain,user,password):
    commons.Title("LOGIN")
    driver.get(data["domain"] % domain)

    # Input Id 
    try :
        commons.FindElementById(":r5:").send_keys(user)
    except :
        commons.FindElementById(":r3:").send_keys(user)
    commons.Content("Input ID")
    
    # Input Pass 
    commons.FindElementById("gw_pass").send_keys(password)
    commons.Content("Input Pass")
    
    # Click Btn
    try :
        commons.FindElementById(":r6:").send_keys(Keys.RETURN)
    except :
        commons.FindElementById(":r4:").send_keys(Keys.RETURN)
        
    commons.Content("Click on button login")
    
    
    
    return True
    
    
