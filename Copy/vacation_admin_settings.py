from lib2to3.pgen2.token import RPAR
from msilib.schema import ComboBox
from vacation_login import driver
from vacation_functions import commons,data,datetime,time,Keys
from vacation_param import Pass,Fail,xpath,Des,pr_ad,pr_rq,pr_ap

def CreateVc():
    i             = 1
    Result_Name   = False 
    Result_Number = False
    Result_Find   = False  
    Vacation_Name = commons.Time()
    
    driver.find_element_by_xpath(pr_ad.dt_ad["bt_create"]).click() 
    if commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_next"]) == True :
        commons.CasePass(pr_ad.ClickNextButton[Pass][Des])
        Name_Vacation = commons.SendKey(Vacation_Name,pr_ad.dt_ad["ip_name"])
        if  Name_Vacation == True :
            Result_Name   = True
            commons.CasePass(pr_ad.InputVacationName[Pass][Des])
        else :
            commons.WriteOnExcel(pr_ad.InputVacationName[Fail])

        driver.find_element_by_xpath(pr_ad.dt_ad["bt_next"]).click()
        Number_Vacation = commons.SendKey(12,pr_ad.dt_ad["number_day_off"])
        if  Number_Vacation == True :
            Result_Number   = True
            commons.CasePass(pr_ad.InputVacationNumber[Pass][Des])
        else :
            commons.WriteOnExcel(pr_ad.InputVacationNumber[Fail])
        if  Result_Name == True and Result_Number == True :
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_save"]).click()
            driver.find_element_by_xpath(pr_ad.dt_ad["vacation_list"]).click()
            time.sleep(1)
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_reresh"]).click()
            
            time.sleep(3)
            List_Vacation  = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
            Total_Vacation = commons.TotalData(List_Vacation)
            if  Total_Vacation == 0:
                commons.WriteOnExcel(pr_ad.DisplayVacation[Fail])
            else:
                while i <= Total_Vacation:
                    Name_Vacation = commons.GetTextWithI(pr_ad.dt_ad["vc_name"] , str(i))
                    if  Name_Vacation == Vacation_Name :
                        Result_Find   = True
                        break
                    i = i + 1
                if  Result_Find == True:
                    commons.WriteOnExcel(pr_ad.DisplayVacation[Pass])
                else :
                    commons.WriteOnExcel(pr_ad.DisplayVacation[Fail])
            
    else :
        commons.WriteOnExcel(pr_ad.ClickNextButton[Fail])
def DeleteVacation():
    i               = 1
    Result_Delete   = False
    Result_Find     = True
    time.sleep(3)
    List_Vacation   = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
    Total_Vacation  = commons.TotalData(List_Vacation)
    if  Total_Vacation == 0 :
        commons.WriteOnExcel(pr_ad.NoVacationToDelete[Pass])
    else :
        # Choose vacation to delete 
        Vacation_Name = commons.GetText(pr_ad.dt_ad["vc_name"] % str(i))
        driver.find_element_by_xpath(pr_ad.dt_ad["bt_delete_vc"]).click()
        if  commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_delete_vc2"]) == True :
            commons.CasePass(pr_ad.ClickOnIconDelete[Pass][Des])
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_delete_vc2"]).click()
            if  commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_close"]) == False :
                commons.CasePass(pr_ad.ClickOnButtonDelete[Pass][Des])
                Result_Delete = True
            else:
                commons.WriteOnExcel(pr_ad.ClickOnButtonDelete[Fail])
        else :
            commons.WriteOnExcel(pr_ad.ClickOnIconDelete[Fail])
        
        List_Vacation     = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
        Total_Vacation    = commons.TotalData(List_Vacation)
        if  Result_Delete == True :
            while i <= Total_Vacation :
                Name_Vacation = commons.GetText(pr_ad.dt_ad["vc_name"] % str(i))
                if  Name_Vacation == Vacation_Name :
                    Result_Find   = False
                    break
                i = i + 1
            if  Result_Find == True :
                commons.WriteOnExcel(pr_ad.DeleteVacation[Pass])
            else :
                commons.WriteOnExcel(pr_ad.DeleteVacation[Fail])
def SelectUserFromDepart():
    
    List_Department  = driver.find_elements_by_xpath(pr_rq.rq_vc["list_depart_cc"])    
    Total_Department = commons.TotalData(List_Department) + 1
    for i in range(1 ,Total_Department):
        time.sleep(1)
        Depart_Has_User = commons.IsDisplayedByXpath(pr_rq.rq_vc["single_depart"] % str(i)) 
        if  Depart_Has_User == True :
            driver.find_element_by_xpath(pr_rq.rq_vc["single_depart"] % str(i)).click()
            List_User  = driver.find_elements_by_xpath(pr_rq.rq_vc["list_user"])
            Total_User = commons.TotalData(List_User) + 1
            for j in range(1 , Total_User):
                Is_User = commons.IsDisplayedByXpath(pr_rq.rq_vc["is_user"] % str(j)) 
                if  Is_User == True:
                    User_Name = driver.find_element_by_xpath(pr_rq.rq_vc["depart"] % (str(i),str(j))).text
                    driver.find_element_by_xpath(pr_rq.rq_vc["click_user"] % (str(i),str(j))).click()
                    return User_Name      
    return False 

def CountAllManager():
    # Count all vacation from all page #
    time.sleep(3)
    i             = 1
    Total_Request = 0
    if commons.IsDisplayedByXpath(pr_ap.vc_ap["check_list"]) == False :
       
        driver.find_element_by_xpath(pr_ap.vc_ap["end_page"]).click()
        time.sleep(2)
        End_Page_Text = commons.GetText(pr_ap.vc_ap["page_current"])
        End_Page      = int(End_Page_Text)
        while  i <= End_Page:
            if i == End_Page :
                driver.find_element_by_xpath(pr_ap.vc_ap["end_page"]).click()
                time.sleep(3)
                List_Manager  = driver.find_elements_by_xpath(pr_ad.dt_ad["list_manager"])
                Total         = commons.TotalData(List_Manager)
                Total_Request = Total_Request + Total
            else:
                Total_Request = Total_Request + 10
            i = i + 1
        driver.find_element_by_xpath(pr_ap.vc_ap["to_first_page"]).click()
    return Total_Request

def AddManager():
    Result_Add   = False
    Result_Find  = False
    i          = 1
    driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_manager"]).click()
    if  commons.IsDisplayedByXpath(pr_ad.dt_ad["ip_search_user"]) == True :
        commons.CasePass(pr_ad.ClickOnManagerButton[Pass][Des])
        Manager_Name = SelectUserFromDepart()
        if Manager_Name != True :
            commons.CasePass(pr_ad.ClickSelectUser[Pass][Des])
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_user"]).click()
            List_Added  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
            Total_Added = commons.TotalData(List_Added)
            if  Total_Added != 0 :
                commons.CasePass(pr_ad.ClickToAddUser[Pass][Des])
                driver.find_element_by_xpath(pr_ad.dt_ad["bt_save_user"]).click()
                if  commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_add_manager"]) ==True :
                    commons.WriteOnExcel(pr_ad.AddManager[Pass])
                    Result_Add = True
                else :
                    commons.WriteOnExcel(pr_ad.AddManager[Fail])
            else :
                commons.WriteOnExcel(pr_ad.ClickToAddUser[Fail])
        else :
            commons.CaseFail(pr_ad.ClickSelectUser[Fail])
        if Result_Add == True :
            Total_Manager = CountAllManager()
            if  Total_Manager == 0 :
                commons.WriteOnExcel(pr_ad.DisplayManager[Fail])
            else :
                Manager_Name  = commons.ReplaceSpace(Manager_Name)
                while i <= Total_Manager :
                    Name_Manager  = commons.GetTextWithI(pr_ad.dt_ad["manager"] , str(i))
                    Name_Manager  = commons.ReplaceSpace(Name_Manager)
                    if  Name_Manager == Manager_Name :
                        Result_Find  = True
                        break
                    i = i + 1
                if Result_Find == True :
                    commons.WriteOnExcel(pr_ad.DisplayManager[Pass])
                else :
                    commons.WriteOnExcel(pr_ad.DisplayManager[Fail])
    else :
        commons.WriteOnExcel(pr_ad.ClickOnManagerButton[Fail])
        
def DeleteManager():
    Result_Find   = False
    Result_Delete = False
    i             = 1
    commons.ClickLinkText("Manager Settings")
    if  commons.IsDisplayedByXpath(pr_ap.vc_ap["check_list"]) == True :
        commons.WriteOnExcel(pr_ad.NoManagerToDelete[Pass])
    else :
        Name_Manager  = commons.GetTextWithI(pr_ad.dt_ad["manager"] , str(1))
        List_Manager  = driver.find_elements_by_xpath(pr_ad.dt_ad["list_manager"])
        Total_Manager = commons.TotalData(List_Manager)
        driver.find_element_by_css_selector(pr_ad.dt_ad["ic_delete"]).click()
        if commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_delete"]) == True :
            commons.CasePass(pr_ad.ClickOnIconDelete[Pass][Des])
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_delete"]).click()
            if  commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_refresh"]) == True :
                commons.WriteOnExcel(pr_ad.DeleteManager[Pass])
                Result_Delete = True
            else :
                commons.WriteOnExcel(pr_ad.DeleteManager[Fail])
        else :
            commons.WriteOnExcel(pr_ad.ClickOnIconDelete[Fail])

        if  Result_Delete == True :
            if  Total_Manager == 1 :
                if  commons.IsDisplayedByXpath(pr_ap.vc_ap["check_list"]) == False :
                    commons.WriteOnExcel(pr_ad.RemovedManager[Pass])
                else :
                    commons.WriteOnExcel(pr_ad.RemovedManager[Fail])
            else :
                while i <= Total_Manager :
                    Manager_Name  = commons.GetTextWithI(pr_ad.dt_ad["manager"] , str(i))
                    Manager_Name  = commons.ReplaceSpace(Name_Manager)
                    if  Name_Manager == Manager_Name :
                        Result_Find  = True
                        break
                    i = i + 1
                if  Result_Find == True :
                    commons.WriteOnExcel(pr_ad.RemovedManager[Fail])
                else :
                    commons.WriteOnExcel(pr_ad.RemovedManager[Pass])

def SelectUser(Button,Msg):
    i     = 1
    Add   = False
    if  commons.IsDisplayedByXpath(pr_ad.dt_ad["ip_search_user"]) == True :
        commons.CasePass(Msg[Pass][Des])
        Manager_Ae  = SelectUserFromDepart()
        Manager_Ae  = commons.ReplaceSpace(Manager_Ae)
        List_Apporver  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
        Total_Apporver = commons.TotalData(List_Apporver)
        while  i <= Total_Apporver :
            Ad_Approver = commons.GetTextWithI(pr_ad.dt_ad["mn_name"],str(i))
            Ad_Approver = commons.ReplaceSpace(Ad_Approver)
            if  Ad_Approver == Manager_Ae :
                Add         = True
                break
            i = i + 1

        driver.find_element_by_xpath(pr_ad.dt_ad[Button]).click()
        time.sleep(3)
        Apporver_List  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
        Apporver_Total = commons.TotalData(Apporver_List)

    else :
        commons.WriteOnExcel(Msg[Fail])
    return  Add , Apporver_Total


def AddArbitraryDecisionSetting():

    
    i      = 1
    j      = 1
    Add    = False
    Save   = True
    Find   = True
    Saved  = False
    driver.find_element_by_xpath(pr_ad.dt_ad["bt_select_approver"]).click()
    if  commons.IsDisplayedByXpath(pr_ad.dt_ad["ip_search_user"]) == True :
        commons.CasePass(pr_ad.ClickOnButtonApprover[Pass][Des])
        Manager_Ad  = SelectUserFromDepart()
        Manager_Ad  = commons.ReplaceSpace(Manager_Ad)
        if  Manager_Ad != True :
            commons.CasePass(pr_ad.ClickSelectUserAd[Pass][Des])
            List_Apporver  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
            Total_Apporver = commons.TotalData(List_Apporver)
            while  i <= Total_Apporver :
                Ad_Approver = commons.GetTextWithI(pr_ad.dt_ad["mn_name"],str(i))
                Ad_Approver = commons.ReplaceSpace(Ad_Approver)
                if  Ad_Approver == Manager_Ad :
                    Add         = True
                    break
                i = i + 1

            driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_arbitrary"]).click()
            time.sleep(3)
            Apporver_List  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
            Apporver_Total = commons.TotalData(Apporver_List)
            if  Add == False :
                if  Apporver_Total == Total_Apporver + 1 :
                    commons.CasePass(pr_ad.ClickToAddUserAd[Pass][Des])
                else :
                    Save  = False
                    commons.WriteOnExcel(pr_ad.ClickToAddUserAd[Fail])
            else:
                if  Apporver_Total == Total_Apporver :
                    commons.CasePass(pr_ad.ClickToAddUserAd[Pass][Des])
                else :
                    Save  = False
                    commons.WriteOnExcel(pr_ad.ClickToAddUserAd[Fail])
            
            if  Save == True :
                driver.find_element_by_xpath(pr_ad.dt_ad["bt_save_user"]).click()
                if commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_select_approver"]) == True :
                    commons.WriteOnExcel(pr_ad.ArbitraryDecision[Pass])
                else :
                    Find = False
                    commons.WriteOnExcel(pr_ad.ArbitraryDecision[Fail])
            
            if  Find == True :
                j    = 1
                driver.find_element_by_xpath(pr_ad.dt_ad["bt_select_approver"]).click()
                Apporver_List  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
                Apporver_Total = commons.TotalData(Apporver_List)
                while  i <= Apporver_Total :
                    Ad_Approver = commons.GetTextWithI(pr_ad.dt_ad["mn_name"],str(j))
                    Ad_Approver = commons.ReplaceSpace(Ad_Approver)
                    if  Ad_Approver == Manager_Ad :
                        Saved       = True
                        break
                    j = j + 1

            if  Saved == True :
                commons.WriteOnExcel(pr_ad.DisplayArbitraryDecision[Pass])
            else :
                commons.WriteOnExcel(pr_ad.DisplayArbitraryDecision[Fail])
        else :
            commons.CaseFail(pr_ad.ClickSelectUserAd[Fail])
    else :
        commons.WriteOnExcel(pr_ad.ClickOnButtonApprover[Fail])
   

def AddApprovalException():
    i     = 1
    Add   = False
    Save  = True
    Find  = True
    driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_approval_exception"]).click()
    if  commons.IsDisplayedByXpath(pr_ad.dt_ad["ip_search_user"]) == True :
        commons.CasePass(pr_ad.ClickOnButtonAdd[Pass][Des])
        Manager_Ae  = SelectUserFromDepart()
        Manager_Ae  = commons.ReplaceSpace(Manager_Ae)
        List_Apporver  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
        Total_Apporver = commons.TotalData(List_Apporver)
        while  i <= Total_Apporver :
            Ad_Approver = commons.GetTextWithI(pr_ad.dt_ad["mn_name"],str(i))
            Ad_Approver = commons.ReplaceSpace(Ad_Approver)
            if  Ad_Approver == Manager_Ae :
                Add         = True
                break
            i = i + 1

        driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_user"]).click()
        time.sleep(3)
        Apporver_List  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
        Apporver_Total = commons.TotalData(Apporver_List)
        if  Add == False :
            if  Apporver_Total == Total_Apporver + 1 :
                commons.CasePass(pr_ad.ClickToAddUserAe[Pass][Des])
            else :
                Save  = False
                commons.WriteOnExcel(pr_ad.ClickToAddUserAe[Fail])
        else:
            if  Apporver_Total == Total_Apporver :
                commons.CasePass(pr_ad.ClickToAddUserAe[Pass][Des])
            else :
                Save  = False
                commons.WriteOnExcel(pr_ad.ClickToAddUserAe[Fail])
        
        if  Save == True :
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_save_user"]).click()
            if commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_add_approval_exception"]) == True :
                commons.WriteOnExcel(pr_ad.ApprovalException[Pass])
            else :
                Find = False
                commons.WriteOnExcel(pr_ad.ApprovalException[Fail])

        if  Find == True :
            j    = 1
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_approval_exception"]).click()
            Apporver_List  = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
            Apporver_Total = commons.TotalData(Apporver_List)
            while  i <= Apporver_Total :
                Ad_Approver = commons.GetTextWithI(pr_ad.dt_ad["mn_name"],str(j))
                Ad_Approver = commons.ReplaceSpace(Ad_Approver)
                if  Ad_Approver == Manager_Ae :
                    Saved       = True
                    break
                j = j + 1

        if  Saved == True :
            commons.WriteOnExcel(pr_ad.DisplayApprovalException[Pass])
        else :
            commons.WriteOnExcel(pr_ad.DisplayApprovalException[Fail]) 

        
    else :
        commons.WriteOnExcel(pr_ad.ClickOnButtonAdd[Fail])



def BasicSettings():
    commons.ClickLinkText("Basic Settings")
    if  commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_add_approval_exception"]) == True :
        commons.WriteOnExcel(pr_ad.AccessSubMenuBs[Pass])
        AddApprovalException()
    else :
        commons.WriteOnExcel(pr_ad.AccessSubMenuBs[Fail])


def ManagerSettings():
    commons.ClickLinkText("Manager Settings")
    if  commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_add_manager"]) == True :
        commons.WriteOnExcel(pr_ad.AccessSubMenuMn[Pass])
        AddManager()
        DeleteManager()

        commons.WaitToClick(pr_ad.dt_ad["mn_approval_settings"])
        if  commons.IsDisplayedByXpath(pr_ad.dt_ad["bt_select_approver"]) == True :
            commons.WriteOnExcel(pr_ad.AccessTabAP[Pass])
            AddArbitraryDecisionSetting()
             
        else :
            commons.WriteOnExcel(pr_ad.AccessTabAP[Fail])
    else :
        commons.WriteOnExcel(pr_ad.AccessSubMenuMn[Fail])
    
def CreateVacation():
    commons.ClickLinkText("Create Vacation")
    if  commons.IsDisplayedByXpath(pr_ad.dt_ad["tab_list_vc"]) == True :
        commons.WriteOnExcel(pr_ad.AccessSubMenuVc[Pass])
        CreateVc()
        DeleteVacation()
    else :
        commons.WriteOnExcel(pr_ad.AccessSubMenuVc[Fail])

def AdminSettings():
    if commons.IsDisplayedByTextLink("Create Vacation") == True :
        #CreateVacation()
        #ManagerSettings()
        BasicSettings()
    else :
        commons.WriteOnExcel(pr_ad.UserNoAdmin[Pass])
