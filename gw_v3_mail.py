from gw_v3_set_up import driver,data,Keys,EC,By,WebDriverWait,json
from gw_v3_functions import commons, time



def AccessMail():
    commons.Title("MENU MAIL")
    driver.find_element_by_xpath(data["mail"]["menu"]).click()
    commons.Content("Click on menu Mail")
    if commons.IsDisplayedByXpath(data["mail"]["menu"]) == True :
        Access = True
        commons.CasePass("[MAIL]Access menu mail")
    else :
        Access = False
        commons.CaseFail("[MAIL]Access menu mail")
    
    return Access

def WritingSetting():
    commons.Title("Writing Setting")
    Reply_To = driver.find_element_by_xpath(data["mail"]["reply_to"])
    Reply_To.clear()
    Reply_To.send_keys(data["mail"]["mail_to"])
    commons.Content("Input Reply-To")
    
    driver.find_element_by_xpath(data["mail"]["setting_save"]).click()
    commons.Content("Click on button save")
    
    if commons.IsDisplayedByXpath(data["pop_save"]) == True :
        m = driver.find_element_by_xpath(data["pop_save"]).text
        print(m)
def Settings():
    # Frame
    driver.switch_to.frame(commons.FindElementById("newMailIframe"))
    
    driver.find_element_by_xpath(data["mail"]["setting"]).click()
    commons.Content("Click on submenu Settings")
    if commons.IsDisplayedByXpath(data["mail"]["settings"]) == True :
        Settings = True
        commons.CasePass("[MAIL]Access left menu settings")
    else :
        Settings = False
        commons.CaseFail("[MAIL]Access left menu settings")
    
    
    if  Settings == True :
        WritingSetting()
    
    
    
    # Close Frame
    commons.SwitchToDefaultContent()
        
    
    
def MenuMail():
    Access = AccessMail() 
    if  Access == True :
        Settings()
    