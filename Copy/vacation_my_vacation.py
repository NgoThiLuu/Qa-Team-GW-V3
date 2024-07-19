import re
from ast import Str
from xmlrpc.client import TRANSPORT_ERROR
from vacation_login import driver
from vacation_param import pr_rq
from vacation_functions import commons,relativedelta,data,datetime,time,timedelta,Pass,Fail
from vacation_param import xpath,type_vc


def info_cc(selected_cc , saved_cc):
    commons.Content(data["cc"]["selected"] % selected_cc)
    commons.Content(data["cc"]["saved"]    % saved_cc)
   
def vacation_request(vacation_request,i):
    Possition_Icon = commons.IsDisplayedByXpath(data["ic_before"])
    vacation_request["vc_name"]      = xpath.LiVacation(i,Possition_Icon)
    vacation_request["vc_date"]      = xpath.LiRequest(i,"da")
    vacation_request["request_date"] = xpath.LiRequest(i,"rd")
    vacation_request["status"]       = xpath.LiRequest(i,"st")
    vacation_request["vc_date"]      = xpath.VcReplace(vacation_request,"da")
    vacation_request["vc_name"]      = xpath.VcReplace(vacation_request,"na")
    return vacation_request

def infor(vacation,title,type_request): 
    hour_use  = type_vc.hour_use(type_request)
    name      = data["infor"]["name"]   % vacation["vacation_name"]
    total     = data["infor"]["total"]  % vacation["total"]
    used      = data["infor"]["used"]   % vacation["used"]
    remain    = data["infor"]["remain"] % vacation["remain"]
    hour      = data["infor"]["hour"]   % str(hour_use)
    data1      = data["infor"]["data"]  % (title,name,total)
    data2     = data["infor"]["data1"]  % (used,remain,hour)
    return data1 + data2

def approver_name(total_app,list_approver):
    try :
        i = 1
        while i <= total_app :
            app_name = commons.GetTextWithI(pr_rq.rq_vc["ap_name"],str(i))
            app_name = app_name.replace(" ", "")
            list_approver.append(app_name)
            i = i + 1  
    except :
        return list_approver

    return list_approver

def check_reason(info_before,info_after,result,type_request):
    info_after["reason"]  = commons.GetText(pr_rq.rq_vc["content_reason"])
    info_before["reason"] = pr_rq.rq_vc["reason_text"]
    msg                   = type_vc.MsgDetailReason(type_request)
    if info_after["reason"] == info_before["reason"] :
        result["rs_resaon"] = True 
        commons.CasePass(msg[Pass]["description"])
    else:
        commons.WriteOnExcel(msg[Fail])
        
def result_number(number_before,number_after,msg):
    if number_before != number_after:
        commons.WriteOnExcel(msg[Fail])
        return False
    else:
        commons.CasePass(msg[Pass]["description"])
        return True
    
def result_before(number_before,number_after,msg):
    if number_before == number_after:
        commons.CasePass(msg[Pass]["description"])
    else:
        commons.WriteOnExcel(msg[Fail])
       
def result_after(a,b,c,msg):
    if  a ==  True and \
        b ==  True and \
        c ==  True :
        commons.CasePass(msg[Pass]["description"])
    else:
        commons.WriteOnExcel(msg[Fail])


def infor_detail(title,info_vc):
    result_approver = isinstance(info_vc["approver"], str)
    if result_approver == False :
        approver = ""
        for name in info_vc["approver"] :
            approver = approver  + name + ","
    else :
        approver =  info_vc["approver"]

    date = xpath.InDetail(info_vc,"if")
    
    if int(date) > 0 :
        vacation_date = xpath.InDetail(info_vc,"re")
    else :
        vacation_date = xpath.InDetail(info_vc,"no")
       
    request_date = xpath.InDetail(info_vc,"dt")
    used         = xpath.InDetail(info_vc,"us")
    reason       = xpath.InDetail(info_vc,"ld")
    approver     = data["detail"]["approver"] % approver
    data1         = data["detail"]["data"]  % (title,vacation_date,request_date)            
    data2        = data["detail"]["data1"] % (used,approver,reason)    
    commons.Content( data1 + data2 )

def SatAndSun(today) :
    WeekDayList = []
    sat = sun = today
    sun += timedelta(days = 6 - sun.weekday())
    sat += timedelta(days = 5 - sat.weekday())
   
    WeekDayList.append(sat)
    WeekDayList.append(sun)
    while sun.month == today.month :
        sat += timedelta(days = 7)
       
        WeekDayList.append(sat)
        sun += timedelta(days = 7)
    
        WeekDayList.append(sun)
    return WeekDayList


def NextMonth(Vacation_Date):
    if  Vacation_Date.day == int(data["month"][str(Vacation_Date.month)]) :
        Vacation_Date = Vacation_Date  + timedelta(days = 1)
        return Vacation_Date
    else :
        return True

def DateToStr(Vacation_Date):
    Vacation_Date = Vacation_Date.strftime("%Y-%m-%d")
    return Vacation_Date


def RequestDate(Used_List):
    Vacation_Date = datetime.date.today() 
    WeekDayList    = SatAndSun(Vacation_Date)
    while DateToStr(Vacation_Date) in Used_List or Vacation_Date in WeekDayList :
        Vacation_Date = Vacation_Date + timedelta(days = 1)
        next          = NextMonth(Vacation_Date)
        if  next != True :
            Vacation_Date = next 
            WeekDayList   = SatAndSun(next)
    return Vacation_Date
   
    
def next_date(Request_Date):
    
    if int(data["month"][str(Request_Date.month)]) == Request_Date.day:
        Request_Date = Request_Date + timedelta(days = 1)

    elif Request_Date.weekday() == 5 :
        Request_Date = Request_Date +  timedelta(days = 2)
        
    elif Request_Date.weekday() == 6 :
        Request_Date = Request_Date + timedelta(days = 1)
        if int(data["month"][str(Request_Date.month)]) == Request_Date.day:
            Request_Date = Request_Date  + timedelta(days = 1)
    else :
        Request_Date = Request_Date + timedelta(days = 1)

    return Request_Date
    

def split_date_from_continuous_date(continuous_date,date_used) :
    if continuous_date.rfind("~") > 0 :
        start_date  = continuous_date[None: int(continuous_date.rfind("~"))]
        start_date  = datetime.datetime.strptime(start_date , '%Y-%m-%d').date()
        end_date    = continuous_date[int(continuous_date.rfind("~")) + 1: None]
        end_date    = datetime.datetime.strptime(end_date , '%Y-%m-%d').date()
        next_date_1 = start_date
        while next_date_1 !=  end_date :
            date_used.append(str(next_date_1))
            if start_date == end_date :
                break
            next_date_1 = next_date(next_date_1)
        date_used.append(str(end_date))
    else :
        date_used.append(continuous_date)
    
def get_vacation_date(continuous_date):
    date_used = []
    if continuous_date.rfind("~") > 0 :
        start_date  = continuous_date[None: int(continuous_date.rfind("~"))]
        start_date  = datetime.datetime.strptime(start_date , '%Y-%m-%d').date()
        end_date    = continuous_date[int(continuous_date.rfind("~")) + 1: None]
        end_date    = datetime.datetime.strptime(end_date , '%Y-%m-%d').date()
        next_date_1 = start_date
        while next_date_1 !=  end_date :
            date_used.append(str(next_date_1))
            if start_date == end_date :
                break
            next_date_1 = next_date(next_date_1)
        date_used.append(str(end_date))
    else :
        date_used.append(continuous_date)        
    return date_used


def choose_end_date(request_date,date_used):
    start_date   = request_date
    request_date = next_date(request_date)
    
    if  start_date != request_date :
        if request_date.weekday() == 5 :
            request_date = request_date + timedelta(days = 2)
            if int(data["month"][str(request_date.month)]) == request_date.day:
                request_date = request_date + timedelta(days = 1)

        if  request_date.weekday() == 6 :
            request_date = request_date + timedelta(days = 1)
            if int(data["month"][str(request_date.month)]) == request_date.day:
                request_date = request_date + timedelta(days = 1)

    end_date = request_date
    if str(end_date) not in date_used:
        return end_date
    else :
        return False

def choose_start_date(Date_Used,Request_Date):
    # Find unused date , not saturday , not sunday , not holiday to use for request vacation # 
    if Request_Date.weekday() == 5 :
        Request_Date = Request_Date + timedelta(days = 2)
    if  Request_Date.weekday() == 6 :
        Request_Date = Request_Date +  timedelta(days = 1)
    if str(Request_Date) in Date_Used  :
        Request_Date = Request_Date + timedelta(days = 1)
        if Request_Date.weekday() == 5 :
            Request_Date = Request_Date + timedelta(days = 2)
        if  Request_Date.weekday() == 6 :
            Request_Date = Request_Date + timedelta(days = 1)
        
        while str(Request_Date) in Date_Used  :
            Request_Date = next_date(Request_Date)
            if Request_Date.weekday() == 5 :
                Request_Date = Request_Date + timedelta(days = 2)
            if  Request_Date.weekday() == 6 :
                Request_Date = Request_Date + timedelta(days = 1)
    return Request_Date


def click_date(request_date):
    if request_date.day < 25:
        for i in range(2,8) :
            for j in range(2,8):
                date_at_calendar = driver.find_element_by_xpath(data["click_date"] % (str(i),str(j)))
                if str(date_at_calendar.text) == str(request_date.day):
                    date_at_calendar.click()
                    return True
        return False
    else:  
        for i in range(2,8) :
            for j in range(2,8):
                date_at_calendar = driver.find_element_by_xpath(data["click_date"] % (str(i),str(j)))
                if str(date_at_calendar.text) == str(request_date.day):
                    date_at_calendar.click()
                    selected_date = commons.GetText(pr_rq.rq_vc["selected_date"])
                    if selected_date.rfind("[") > 0:
                        selected_date = xpath.SelectDate(selected_date,"y")
                    else:
                        selected_date = xpath.SelectDate(selected_date,"n")
                        #selected_date = selected_date[12: None].replace(" ", "")
                    if str(request_date) == selected_date:
                        return True
                    
                    else:
                        date_at_calendar.click()   
                        
        return False

def get_days_and_hour(data_column):
    # Get days , hour of column data 4.5D , 4D4H , - #  
    number_day = {"day":"","hour":""}

    if data_column.replace(" ", "") == "-":
        number_day["day"] = float(0)
    elif data_column.rfind("D") < 0 :
        number_day["day"] = float(0)
    else :
        number_day["day"] = float(xpath.DataColumn(data_column,"d"))
        

    if data_column.rfind("H") < 0 :
        number_day["hour"] = float(0)
    else :
        number_day["hour"] = float(xpath.DataColumn(data_column,"h"))

    return number_day

def change_hour_to_day(tp1,tp2,oneday,plus,hour_use,use_hour_unit,type_request):

    # USE HOUR UNIT FOR VACATION # 
    # The unit for calculation is hour ,convert to hour before calculation #
   
    number = ""
    if use_hour_unit == True :
        
        # Hour_use is int  ,ex hour_use = 4 #
        # Plus or minus data 2 column #
        if tp2 != None:
            number = TrueCalculateDataForTwoColumns(plus,oneday,tp1,tp2)
            
        else:
        # Plus or minus data of 1 column with number #
            number = TrueCalculateDataForColumnAndNumber(plus,oneday,type_request,tp1,hour_use)
    else:
        # NOT USE HOUR UNIT FOR REQUEST #
        # The unit for calculation is days ,convert to days before calculation #
        # Hour used have to convert to day #
        
        if tp2 != None:
            number = FaseCalculateDataForTwoColumns(plus,tp1,tp2)
        else:
            number = FaseCalculateDataForColumnAndNumber(plus,type_request,tp1,hour_use)
    return number

def TrueCalculateDataForTwoColumns(plus,oneday,tp1,tp2) :
   
    tp1 = get_days_and_hour(tp1)
    tp2 = get_days_and_hour(tp2)

    if  plus =="plus":
        total_hour = xpath.TotalHour(tp1,tp2,oneday)
        day = total_hour // oneday
        hour = total_hour % oneday
    
    if  plus == "minus":
        l1 = int(tp1["day"])*oneday + int(tp1["hour"])
        l2 = int(tp2["day"])*oneday + int(tp2["hour"])
        total_hour_remain = l1-l2
        if total_hour_remain < 0:
            total_hour_remain = total_hour_remain*(-1)
        day  = total_hour_remain // oneday
        hour = total_hour_remain % oneday
    
    if str(day) == "0" and str(hour) == "0":
        return "0"
    else:
        if str(day) == "0" :
            return str(hour) + "H"
        elif str(hour) == "0" :
            return str(day) + "D"
        else :
            return str(day) + "D" + " " + str(hour) + "H"

def TrueCalculateDataForColumnAndNumber(plus,oneday,type_request,tp1,hour_use) :
    hour_use = int(hour_use)
    tp1 = get_days_and_hour(tp1)

    if plus =="plus":
        if  type_request == "half_day"  :
            hour_use   = 4
            total_hour = int(tp1["hour"])  + int(tp1["day"]*oneday) + hour_use

        elif type_request  == "hour":
            hour_use   = 2
            total_hour = int(tp1["hour"])  + int(tp1["day"]*oneday) + hour_use

        else:
            total_hour = int(tp1["hour"])  + int(tp1["day"]*oneday) + hour_use*oneday
        day  = total_hour // oneday
        hour = total_hour % oneday

    
    if  plus =="minus":
        l1 = int(tp1["day"])*oneday + int(tp1["hour"])
        if  type_request == "half_day"  :
            hour_use          = 4
            total_hour_remain = l1 - hour_use

        elif type_request == "hour":
            hour_use          = 2
            total_hour_remain = l1 - hour_use

        else:
            total_hour_remain = l1 - int(hour_use)*oneday
        day  = total_hour_remain // oneday
        hour = total_hour_remain % oneday

    if str(day) == "0" and str(hour) == "0":
        return "0"
    else:
        
        if str(day) == "0" :
            return str(hour) + "H"
        elif str(hour) == "0" :
            return str(day) + "D"
        else:
            return str(day) + "D " + str(hour) + "H"

def FaseCalculateDataForTwoColumns(plus,tp1,tp2) :
    tp1 = get_days_and_hour(tp1)
    tp2 = get_days_and_hour(tp2)

    # Day after plus #
    if  plus == "plus":
        day = float(tp1["day"]) + float(tp2["day"])
    
    # Day after minus #
    if  plus == "minus":
        day = float(tp1["day"]) - float(tp2["day"])
        if day < 0:
            day = day*(-1)
    
    if str(day) != "0" :
        if str(day)[int(str(day).rfind(".")) + 1: None] == "0":
            return str(day)[None: int(str(day).rfind("."))] + "D"
        else:
            return str(day) + "D"
    else:
        return "0"

def FaseCalculateDataForColumnAndNumber(plus,type_request,tp1,hour_use) :
    if type_request == "half_day" :
        hour_use = 0.5

    # Hour_use is fload ,ex hour_use  =  0.5 #
    # Plus or minus data of 1 column with number #

    tp1 = get_days_and_hour(tp1)
    if plus == "plus":
        day = float(tp1["day"]) + float(hour_use)
    
    if  plus == "minus":
        day = float(tp1["day"]) - float(hour_use)
        if day < 0 :
            day = day *(-1)

    if str(day) != "0.0" :      
        if str(day)[int(str(day).rfind("."))+1: None] == "0":
            return str(day)[None: int(str(day).rfind("."))] + "D"
        else:
            return str(day) + "D"
    else :
        return "0"

def select_user_from_depart():
    list_department  = driver.find_elements_by_xpath(pr_rq.rq_vc["list_depart_cc"])    
    total_department = commons.TotalData(list_department)
    for i in range(1,total_department+1):
        time.sleep(1)
        depart_has_user = commons.IsDisplayedByXpath(pr_rq.rq_vc["single_depart"] % str(i)) 
        if depart_has_user == True :
            driver.find_element_by_xpath(pr_rq.rq_vc["single_depart"] % str(i)).click()
            total_user = driver.find_elements_by_xpath(pr_rq.rq_vc["list_user"])
            for j in range(1 , len(total_user) + 1):
                is_user = commons.IsDisplayedByXpath(pr_rq.rq_vc["is_user"] % str(j)) 
                if is_user == True:
                    user_name = driver.find_element_by_xpath(pr_rq.rq_vc["depart"] % (str(i),str(j))).text
                    driver.find_element_by_xpath(pr_rq.rq_vc["click_user"] % (str(i),str(j))).click()
                    return user_name      
    return False 

def hours_set_from_time_card(type_request):

    # Specific working hours from time card  #  
    hour_use = commons.GetText(pr_rq.rq_vc["hour_use"])
    hour_use = xpath.TcHourUse(hour_use,"d")

    if type_request == "all" :
        if len(hour_use) == 0 :
            return 8
        else :
            hour_use = xpath.TcHour(hour_use)
            if hour_use == 1 :
                return 8
            else :
                return hour_use

    elif type_request == "hour":
        hour_use = commons.GetText(pr_rq.rq_vc["hour_use_h"])
        hour_use = xpath.TcHourUse(hour_use,"h")
        return xpath.TcHour(hour_use)

    else:
        return 4   
    
def vacation_use_for_request():
    # Information about number of selected vacation for request vacation #  
    vacation =  {"vacation_name"  : "",
                "number_of_days" : "",
                "number_of_hours": ""
                }
    vacation_name = commons.GetText(pr_rq.rq_vc["vacation_name"])
    if vacation_name.rfind("(") > 0 :
        vacation["vacation_name"] = xpath.VacationName(vacation_name)
    else :
        vacation["vacation_name"] = vacation_name

    days = xpath.DaysUse(vacation_name)
    if int(days.rfind("D")) > 0 :
        vacation["number_of_days"] = xpath.VacationNumberDay(vacation_name,"d")
        if int(days.rfind("H")) < 0:
            vacation["number_of_hours"] = "0"
        else:
            vacation["number_of_hours"] = xpath.VacationNumberHours(vacation_name)
    else:
        vacation["number_of_days"] = "0"
        if int(days.rfind("H")) > 0 :
            vacation["number_of_hours"] = xpath.VacationNumberDay(vacation_name,"h")
        else:
            vacation["number_of_hours"] = "0"

    if int(days.rfind("D")) < 0 and int(days.rfind("H")) < 0 :
        vacation["number_of_days"]  = "0"
        vacation["number_of_hours"] = "0"

    return vacation

def hour_used(use_hour_unit,type_use):
    if bool(use_hour_unit) == True :
        if type_use =="all":
            hour_use = 1
        elif type_use == "hour":
            hour_use = hours_set_from_time_card(type_use)
        elif type_use == "vc_con":
            hour_use = 2
        else:
            hour_use = 0.5
        return float(hour_use)
    else:
        if type_use == "all":
            hour_use = 1 
        else:
            hour_use = 4/10
        return float(hour_use)

def click_date_time_card(request_date):
    day = request_date.day
    for i in range(1,6) :
        for j in range(1,6):
            if i == 1 :
                date_to_click = driver.find_element_by_xpath(data["get_date"] % (str(i) , str(i)))
                date_at_calendar = date_to_click.text
                if day > 25 and date_at_calendar == str(day) :
                    pass
                else:
                    if date_at_calendar == str(day):
                        date_to_click.click()
                        return True
            else:
                date_to_click = driver.find_element_by_xpath(data["get_date"] % (str(i) , str(i)))
                date_at_calendar = date_to_click.text
                if date_at_calendar == str(day):
                    date_to_click.click()
                    return True

def view_detail_used(type_request,use_hour_unit):
    
    if type_request == "all":
        return "1D"
    elif type_request == "vc_con":
        return "2D"
    elif type_request == "hour":
        return "2H"
    else:
        if use_hour_unit == True:
            return "4H"
        else:
            return "0.5D"

def approver_list(check_no_approver):
    
    if len(check_no_approver) != 0 :
        return True
    else :
        commons.WriteOnExcel(pr_rq.NoApprover)
        return False

def click_approver():
    result ={
        "user_name"   : "",
        "is_selected" : ""
    }
    user                = commons.ReturnElement(pr_rq.rq_vc["sl_ap_firt"])
    result["user_name"] = user.text
    user.click()
    button              = commons.ReturnElement(pr_rq.rq_vc["bt_ap_firt"])
    is_selected         = button.is_selected()
    msg                 = pr_rq.SelectApprover
    if is_selected == True :
        result["is_selected"] = True 
        commons.CasePass(msg[Pass]["description"])
        commons.ClickElementWithXpath(pr_rq.rq_vc["bt_add"])
    else :
        result["is_selected"] = False
        commons.WriteOnExcel(msg[Fail])
    return result

def add_approver():
    msg = pr_rq.AddApprover
    if  commons.IsDisplayedByXpath(pr_rq.rq_vc["check_bt_save"]) == False :
        commons.CasePass(msg[Pass]["description"])
        driver.find_element_by_xpath(pr_rq.rq_vc["bt_save"]).click()
        return True
    else :
        commons.WriteOnExcel(msg[Fail])
        return False

def save_approver(user_name,select_approver):
    if commons.IsDisplayedByXpath( pr_rq.rq_vc["bt_select_cc"]) == True:
        msg = pr_rq.SaveApprover
        commons.CasePass(msg[Pass]["description"])
        commons.Scroll()

        to_element = driver.find_element_by_xpath(pr_rq.rq_vc["bt_select_cc"])
        commons.ScrollingToTarget(to_element)
    
        list_app  = driver.find_elements_by_xpath(pr_rq.rq_vc["list_approver1"])
        total_app = commons.TotalData(list_app)
        msg0 = pr_rq.Approver
        if total_app == 0:
            commons.WriteOnExcel(msg0[Fail])
        else:
            for i in range(1 , total_app + 1):
                app_name = commons.GetTextWithI(data["approver_name"] , str(i))
                if app_name.strip() == user_name.strip():
                    commons.CasePass(msg0[Pass]["description"])
                    select_approver["result_approver"] = True
                    select_approver["approver_name"]   = user_name.strip()
                    break
            if select_approver["result_approver"] == False :
                commons.WriteOnExcel(msg0[Fail])
    else:
        commons.WriteOnExcel(msg[Fail])

def ApproverIsApprovalException(type_request,info_before):
    info_before["approver"] = pr_rq.msg["approval_exception"]
    msg = type_vc.MsgDetailApproverException(type_request)
    if commons.IsDisplayedByXpath(pr_rq.rq_vc["approval_exception"]) == False:
        commons.CasePass(msg[Pass]["description"])
    else:
        commons.WriteOnExcel(msg[Fail])

def ApproverIsApproverLine(total_approver,approver,result,info_after,content_approver,msg):
    i = 1
    while i <= total_approver :
        try :
            approver_name = commons.GetTextWithI(pr_rq.rq_vc["approver"] , str(i))
            approver_name = commons.ReplaceSpace(approver_name)
            if approver_name.strip() in approver["approver_name"]:
                content_approver       = True
                result["rs_approver"]  = True
                info_after["approver"] = approver["approver_name"]
                commons.CasePass(msg[Pass]["description"])
                break
        except :
            pass
        i = i + 1
    if content_approver == False :
        commons.WriteOnExcel(msg[Fail])

def ApproverIsSelect(total_approver,approver,result,info_after,content_approver,msg):
    j = 1
    while j <= total_approver :
        try :
            approver_name = commons.GetTextWithI(pr_rq.rq_vc["approver"] , str(j))
            approver_name = commons.ReplaceSpace(approver_name)
            if approver_name == approver["approver_name"][0]:
                content_approver       = True
                result["rs_approver"]  = True
                info_after["approver"] = approver["approver_name"]
                commons.CasePass(msg[Pass]["description"])
                break
        except :
            pass
        j = j + 1
    if content_approver == False :
        commons.WriteOnExcel(msg[Fail])


def check_approver_reason(type_request,approver,info_after,info_before):
    
    content_approver = False
    result = xpath.ParApproverResult()
    if type_request == "all":

        # Check resaon #
        check_reason(info_before,info_after,result,type_request)

        # Check approver #
        if approver["approval_exception"] == True:
            ApproverIsApprovalException(type_request,info_before)

        else:
            
            info_before["approver"] = approver["approver_name"]
            msg = type_vc.MsgDetailApprover(type_request)
            if  commons.IsDisplayedByXpath( pr_rq.rq_vc["approval_exception"]) == True:
                time.sleep(2)
                list_approver  = driver.find_elements_by_xpath( pr_rq.rq_vc["content_vc_approver"])
                total_approver = commons.TotalData(list_approver)

                if approver["result_approver"] == True :
                    ApproverIsApproverLine(total_approver,approver,result,info_after,content_approver,msg)
                else:
                    msg = type_vc.MsgDetailApproverLine(type_request)
                    ApproverIsSelect(total_approver,approver,result,info_after,content_approver,msg)
            else:
                commons.WriteOnExcel(msg[Fail])

        result["info_before"] = info_before
        result["info_after"]  = info_after
        return result

def add_cc():
    msg = pr_rq.SelectCC
    if commons.IsDisplayedByXpath( pr_rq.rq_vc["bt_add_cc"]) == True:
        commons.CasePass(msg[Pass]["description"])
        driver.find_element_by_xpath( pr_rq.rq_vc["dele_all_cc"]).click()
        return True
    else :
        commons.WriteOnExcel(msg[Fail])
        return False

def SelectedCc(selected_cc):
    if selected_cc != False :

        commons.CasePass(pr_rq.CCClickUser)
        time.sleep(1)

        driver.find_element_by_xpath( pr_rq.rq_vc["bt_add_cc"]).click()
        commons.CasePass(pr_rq.CCClickAdd)

        driver.find_element_by_xpath( pr_rq.rq_vc["bt_save"]).click()
        commons.CasePass(pr_rq.CCClickSave)
       
        time.sleep(1) 
        return True
    else :
        commons.WriteOnExcel(pr_rq.ClickCC[Fail])
        return False

def check_saved_cc(selected_cc):
    msg = pr_rq.CCSaveSelect
    if commons.IsDisplayedByXpath(pr_rq.rq_vc["bt_select_cc"]) == True :
        commons.CasePass(msg[Pass]["description"])
        list_cc  = driver.find_elements_by_xpath(pr_rq.rq_vc["list_cc"])
        total_cc = commons.TotalData(list_cc)
        msgo     = pr_rq.CC

        if total_cc == 0:
            commons.WriteOnExcel(msgo[Fail])
        else:
            i = 1
            result_cc = False
            while  i <= total_cc :
                cc_name = commons.GetTextWithI(data["cc_name"] , str(i))
                if cc_name.strip() == selected_cc.strip():
                    result_cc = True
                    commons.CasePass(msgo[Pass]["description"])
                    break
                i = i+1
            if result_cc == False :
                commons.WriteOnExcel(msgo[Fail])
            info_cc(cc_name , selected_cc)
    else:
        commons.WriteOnExcel(msg[Fail])

def view_vacation_date(info_before,vc_rq,type_request,info_after):
    Result = False
    try:
        info_before["vc_date"] = vc_rq["vc_date"]
        Date = commons.GetText(pr_rq.rq_vc["content_vc_date"])
        Date = commons.ReplaceSpace(Date)

        if type_request =="all":
            Date = xpath.VacationDate("d",Date)

        elif type_request == "vc_con":
            Date = commons.ReplaceSpace(vc_rq["vc_date"])
            
        elif type_request == "hour":
            Date = xpath.VacationDate("d",Date)

        else:
            Date = xpath.VacationDate("o",Date) 

        info_after["vc_date"] = Date
        DateO = commons.ReplaceSpace(vc_rq["vc_date"])
        msg   = type_vc.MsgDetailVacationDate(type_request)

        if  Date == DateO :
            Result = True
            commons.CasePass(msg[Pass]["description"])
            
        else:
            commons.WriteOnExcel(msg[Fail])
    except:
        pass
    return Result

def view_detail_number_of_days_used(info_before,use_hour_unit,type_request,info_after):
    result = False
    try :
        info_after["used"]  = commons.GetText(pr_rq.rq_vc["content_vc_use"])
        info_before["used"] = view_detail_used(type_request,use_hour_unit)
        msg                 = type_vc.MsgDetailUsed(type_request)

        if info_after["used"] == info_before["used"] :
            result = True
            commons.CasePass(msg[Pass]["description"])
        else:
            commons.WriteOnExcel(msg[Fail])
    except:
        pass

    return result

def view_detail_request_date(info_before,vc_rq,type_request,info_after):
    result_re_date = False
    try :
        info_before["request_date"] = vc_rq["request_date"]
        info_after["request_date"]  = commons.GetText(pr_rq.rq_vc["content_request_date"])
        msg                         = type_vc.MsgDetailRequestDate(type_request)

        if info_after["request_date"] == vc_rq["request_date"] :
            result_re_date = True
            commons.CasePass(msg[Pass]["description"])
        else:
            commons.WriteOnExcel(msg[Fail])
    except:
        pass
    return result_re_date

def vacation_date_is_used(list_vc):
    for vc_date in list_vc :
        if vc_date["status"] == "Request" or vc_date["status"] == "Approved" :
            if len(vc_date["vc_date"]) > 0 :
                date = vc_date["vc_date"][0]
            else :
                date = vc_date["vc_date"]
            request_date =  datetime.datetime.strptime(date.replace("-", "/"),"%Y/%m/%d")
            commons.ClickLinkText("Request Vacation")
            return request_date
        else :
            return False

    
def view_detail_approver_and_reason(type_request,approver,info_after,info_before):
    rs_resaon = rs_approver = False
    try :
        approver_reason = check_approver_reason(type_request,approver,info_after,info_after)
        rs_resaon   = approver_reason["rs_resaon"]
        rs_approver = approver_reason["rs_approver"]
        info_before = approver_reason["info_before"]
        info_after  = approver_reason["info_after"]  
    except:
        pass
    return rs_resaon , rs_approver , info_before , info_after

def hours_set_from_time_card(type_request):
    # Specific working hours from time card  #  
    hour_use = commons.GetText(pr_rq.rq_vc["hour_use"])
    hour_use = xpath.TcHourUse(hour_use,"d")

    if type_request == "all" :
        if len(hour_use) == 0 :
            return 8
        else :
            hour_use = xpath.TcHour(hour_use)
            if hour_use == 1 :
                return 8
            else :
                return hour_use
    elif type_request == "hour":
        hour_use = commons.GetText(pr_rq.rq_vc["hour_use_h"])
        hour_use = xpath.TcHourUse(hour_use,"h")
        return xpath.TcHour(hour_use) 
    else:
        return 4   
        
def CheckClickedDate(request_date):
    result_click = click_date(request_date)
    if result_click == True:
        commons.CasePass(pr_rq.SelectDate)
    else:
        commons.CasePass(pr_rq.SelectDate)

def GetDateUsed():
    # Go to my vacation to take used date # 
    i         = 1 
    Date_Used = []
    commons.ClickElementWithText("My Vacation Status")
    time.sleep(3)

    rows = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
    total_request = commons.TotalData(rows)
    while i <= total_request:
        if i == 1:
            if commons.IsDisplayedByXpath(pr_rq.rq_vc["check_list_re"]) == True :
                break
            else:
                date = commons.GetTextWithI(pr_rq.rq_vc["request"] , str(i))
                split_date_from_continuous_date(date,Date_Used)
        else:
            date = commons.GetTextWithI(pr_rq.rq_vc["request"] , str(i))
            split_date_from_continuous_date(date,Date_Used)
        i = i + 1
    return Date_Used


def select_date_to_request_leave_for_vacation_consecutive(Date_Used):
    commons.Title("Select Date") 
    list_date = []
    # Get date list #
    date       = datetime.date.today() 
    start_date = choose_start_date(Date_Used,date)
    end_date   = choose_end_date(start_date,Date_Used)
    if end_date == False:
        while end_date == False :
            end_date = choose_end_date(start_date,Date_Used)
            if end_date != False :
                list_date.append(start_date)
                list_date.append(end_date)
                break
            start_date = next_date(start_date)
            start_date = choose_start_date(Date_Used,start_date)
    else:
        list_date.append(start_date)
        list_date.append(end_date)

    # Select date from calendar # 
    commons.ClickElementWithText("Request Vacation")
    for i in list_date:
        request_date  = i
        current_month = commons.GetText(pr_rq.rq_vc["current_month"])[5:None] 
        
        if int(request_date.month) < 10 :
            request_month = "0" + str(request_date.month)
        else :
            request_month = request_date.month
        
        if str(current_month) == str(request_month) :
            CheckClickedDate(request_date)
        else:
            driver.find_element_by_xpath(pr_rq.rq_vc["icon_next_month"]).click()
            CheckClickedDate(request_date)
    
        time.sleep(2)

    # Check if click date is wrong then click again 
    selected_date = commons.GetText(pr_rq.rq_vc["selected_date"])
    if selected_date.rfind("[") > 0:
        selected_date = xpath.SelectDate(selected_date,"y") 
        for date in list_date :
            if selected_date != DateToStr(date) :
                click_date(date)

    return list_date

def select_date_to_request_leave():
    commons.Title("Select Date") 
    result_click = False
    # Go to my vacation to take used days # 
    Date_Used    = GetDateUsed()
    
    # Find unused days , not saturday , not sunday , not holiday to use for request vacation # 
    request_date = RequestDate(Date_Used)
    
    # Select date from calendar # 
    commons.ClickElementWithText("Request Vacation")
    commons.PopupTimeCard()
    current_month = commons.GetText(pr_rq.rq_vc["current_month"])[5:None] 
    if int(request_date.month) < 10 :
        request_month = "0" + str(request_date.month)
    else :
        request_month = request_date.month

    if str(current_month) == str(request_month) :
        result_click = click_date(request_date)
        if result_click == True:
            commons.CasePass(pr_rq.SelectDate)
            return request_date
        else:
            commons.CaseFail(pr_rq.SelectDate)
            return False
    else:
        driver.find_element_by_xpath(pr_rq.rq_vc["icon_next_month"]).click()
        result_click = click_date(request_date)
        if result_click == True:
            commons.CasePass(pr_rq.SelectDate)
            return request_date
        else:
            commons.CaseFail(pr_rq.SelectDate)
            return False

def get_vacation_date_and_status():
    i = 1 
    list_vc   = [] 
    
    # Go to my vacation to take used days # 
    commons.ClickElementWithText("My Vacation Status")
    time.sleep(3)
    rows          = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
    total_request = commons.TotalData(rows)
    while i <= total_request:
        if i == 1:
            if commons.IsDisplayedByXpath(pr_rq.rq_vc["check_list_re"]) == True :
                break
            else:
                date    = commons.GetTextWithI(pr_rq.rq_vc["request"] , str(i))
                vc_date = get_vacation_date(date)
                vc_stat = commons.GetTextWithI(pr_rq.rq_vc["re_status"] , str(i))
                list_vc.append(xpath.ParDateAndStatus(vc_date,vc_stat))

        else:
            date    = commons.GetTextWithI(pr_rq.rq_vc["request"] , str(i))
            vc_date = get_vacation_date(date)
            vc_stat = commons.GetTextWithI(pr_rq.rq_vc["re_status"] , str(i))
            list_vc.append(xpath.ParDateAndStatus(vc_date,vc_stat))
        i = i+1

    return list_vc

def ResultApprover():
    result = False
    approver = select_approver()
    if approver["result_approver"] == True:
        result = True
    else:
        if approver["approval_line"] == True:
            result = True
        if approver["approval_exception"] == True:
            result = True
    return result , approver
        
def available_vacation():

    # Get data of each vacation at available vacation table #  
    i              = 1 
    total_vacation = 0 
    all_vacation   = []
    commons.ClickElementWithText("Request Vacation")
    tbody = driver.find_element_by_xpath(data["available_vacation"]["tbody"])
    rows  = tbody.find_elements_by_tag_name("tr")
    total_vacation = commons.TotalData(rows)

    while i <= total_vacation:
        vacation   = xpath.ParVacation()
        total_days = xpath.VcAvailable(i,"to")
        if total_days != "-" :
            vacation_name             = xpath.VcAvailable(i,"na")
            vacation["total"]         = total_days
            vacation["used"]          = xpath.VcAvailable(i,"us")
            vacation["remain"]        = xpath.VcAvailable(i,"re")
            vacation["expiration"]    = xpath.VcAvailable(i,"ex")
            vacation["start"]         = xpath.VcStart(vacation_name)
            vacation["vacation_name"] = xpath.VcChange(vacation, vacation_name)
            all_vacation.append(vacation)
        else :
            vacation["vacation_name"] = xpath.VcAvailable(i,"na")
            vacation["used"]          = xpath.VcAvailable(i,"us") 
            vacation["total"]         = vacation["remain"]     = "-"
            vacation["start"]         = vacation["expiration"] = "-"
            all_vacation.append(vacation)

        i = i + 1 
    return all_vacation

def total_vacation():
    
    # Total vacation availabel of user #  
    i = 1 
    total_vacation_can_use = 0
    commons.ClickElementWithText("Request Vacation")
    time.sleep(3)

    tbody = driver.find_element_by_xpath(data["available_vacation"]["tbody"])
    rows  = tbody.find_elements_by_tag_name("tr")
    total_vacation =commons.TotalData(rows)
    while i <= total_vacation:
        remain = commons.GetTextWithI(data["available"]["re"] , str(i))
        if str(remain.strip()) != "0" :
            total_vacation_can_use += 1
        i = i+1

    return total_vacation_can_use

def check_number_of_days_off(before,after,hour_use,vc_name,oneday,use_hour_unit,type_request):

    commons.Title("Available Vacation")
    infor_before = infor_after = " " 

    # Check number of days before request #
    for vacation in before:
        if vacation["vacation_name"] == vc_name :
            vc_bf_use         = vacation
            infor_before      = infor(vc_bf_use,"Info Vacation before request",type_request)
            msg_total         = type_vc.MsgBeforeRequest(type_request)
            ResultBeforeOther = xpath.ResultBeforeOther(vc_bf_use,msg_total)
            HourToDaysUsed    = xpath.HourToDaysUsed(vc_bf_use,oneday,hour_use,use_hour_unit,type_request)
            
            if vc_bf_use["total"] != "-" :
                days              = change_hour_to_day(**HourToDaysUsed)
                ResultBefore      = xpath.ResultBefore(vc_bf_use,days,msg_total)
                result_before(**ResultBefore)
            else :
                result_before(**ResultBeforeOther)
            break

    # Check number of days after request # 
    for vacation in after: 
        if vacation["vacation_name"] == vc_name :
            vc_af_use       = vacation
            infor_after     = infor(vc_af_use,"Info Vacation after request ",type_request)
            msg_total       = type_vc.MsgRequestTotal(type_request) 
            msg_used        = type_vc.MsgRequestUsed(type_request)
            msg_remain      = type_vc.MsgRequestRemain(type_request)
            msg_all         = type_vc.MsgAfterRequest(type_request)
            UHourToDaysNone = xpath.UHourToDaysNone(vc_bf_use,oneday,hour_use,use_hour_unit,type_request)
            RHourToDaysNone = xpath.RHourToDaysNone(vc_bf_use,oneday,hour_use,use_hour_unit,type_request)
            
            # vacation type is regular / grant #
            if  vc_bf_use["total"] != "-" : 
                total = result_number(vc_af_use["total"],vc_bf_use["total"],msg_total)
            
                used_before_plus_used = change_hour_to_day(**UHourToDaysNone)
                used = result_number(vc_af_use["used"],used_before_plus_used,msg_used)

                remain_before_minus_used = change_hour_to_day(**RHourToDaysNone)
                remain = result_number(vc_af_use["remain"],remain_before_minus_used,msg_remain)
                
                result_after(total,used,remain,msg_all)

            # vacation type is Other #
            else :
                total = result_number(vc_af_use["total"],"-",msg_total)
                used_before_plus_used =  change_hour_to_day(**UHourToDaysNone)
                used     = result_number(vc_af_use["used"],used_before_plus_used,msg_used)
                remain   = result_number(vc_af_use["remain"],"-",msg_remain)
                result_after(total,used,remain,msg_all)

            break
    
    commons.Content(infor_before)
    commons.Content(infor_after)


def ChooseVacationToCancel(Total_Request):
    # Get all requests that can be use for cancel #
    time.sleep(3)
    i             = 1
    Result_Cancel        = False
    Result_Find          = False
    Time_Cancel          = True 
    Status_Before_Cancel = True
    Request_To_Cancel    = True
    Approved_List = ["Approved"]
    while i<= Total_Request:
        Infor_Request = info_request_list(i)
        if commons.IsDisplayedByXpath(pr_rq.rq_vc["re_ic"] % str(i-1)) == True :
            # If vacation request is approved , check vacation date (> or = ) today => Can cancel #
            if  Infor_Request["status"]  in Approved_List :
                Time        = xpath.TimeComparison(Infor_Request,commons.Today())
                Time_Cancel = TimeComparison(**Time) 
                
            if  Time_Cancel == True :
                Result_Find          = True
                Status_Before_Cancel = Infor_Request["status"]
                Request_To_Cancel    = Infor_Request
                driver.find_element_by_xpath(pr_rq.rq_vc["re_ic"] % str(i-1)).click()

                if  commons.IsDisplayedByXpath(pr_rq.rq_vc["bt_cancel_request"]) == True :  
                    commons.CasePass(pr_rq.ClickOnCancelIcon[Pass]["description"])
                    driver.find_element_by_xpath( pr_rq.rq_vc["bt_cancel_request"]).click()
                    time.sleep(1)

                    if commons.IsDisplayedByXpath(pr_rq.rq_vc["bt_cancel"]) == False :
                        commons.WriteOnExcel(pr_rq.CancelRequest[Pass])
                        Result_Cancel = True
                    else:
                        commons.WriteOnExcel(pr_rq.CancelRequest[Fail])
                    break 
                else:
                    commons.WriteOnExcel(pr_rq.ClickOnCancelIcon[Fail])
        i = i + 1 

    return Result_Find , Result_Cancel , Status_Before_Cancel , Request_To_Cancel

def CheckStatusAfterCancel(Total_Request,Request_To_Cancel,Status_Before_Cancel):
    i = 1
    check_number = False # If cancel successuly => Check number of days off 
    commons.ClickElementWithText("My Vacation Status")
    while i<= Total_Request:
        Infor_Request         = info_request_list(i)
        Find_request_canceled = two_requests_are_the_same(Request_To_Cancel,Infor_Request)

        if  Find_request_canceled == True:
            if  Status_Before_Cancel == "Request" :
                if  Infor_Request["status"] == "Cancelled":
                    commons.WriteOnExcel(pr_rq.StatusOfRequest[Pass])
                    check_number = True
                    break
                else:
                    commons.WriteOnExcel(pr_rq.StatusOfRequest[Fail])
                    break
            else:
                # Status before is "Approved" ,"Approved[1/3],..."
                if  Status_Before_Cancel == "Approved" :
                    commons.ClickElementWithXpath( pr_rq.rq_vc['cancel'] % (i,i-1))
                    time.sleep(2)
                    List_Approver = driver.find_elements_by_xpath( pr_rq.rq_vc["content_vc_approver"])
                    Total_Approver = commons.TotalData(List_Approver)
                    if Total_Approver == 0:
                        if  Infor_Request["status"] == "Cancelled":
                            commons.WriteOnExcel(pr_rq.StatusOfRequest[Pass])
                            check_number = True
                            break
                            
                        else: 
                            commons.WriteOnExcel(pr_rq.StatusOfRequest[Fail])
                            break

                if  Infor_Request["status"] == "User cancel":
                    commons.WriteOnExcel(pr_rq.StatusIsUserCancel[Pass])
                    break
                else:
                    commons.WriteOnExcel(pr_rq.StatusIsUserCancel[Fail])
                    break
            
        i = i + 1 
    return check_number , Infor_Request

def CheckNumberAfterCancel(Infor_Request,number_before_cancel,oneday,use_hour_unit):
    hours_used          = float(re.search(r'\d+',Infor_Request["use"]).group(0)) 
    number_after_cancel = available_vacation()
    hours               = Infor_Request["use"].strip()
    if Infor_Request["vacation_name"].rfind("~") > 0 :
        vc_name = Infor_Request["vacation_name"].replace("\n","")
        vc_name = xpath.VcName(vc_name)
    else :
        vc_name = Infor_Request["vacation_name"]
    if hours == "1D" :
        type_request = "all"
    elif hours == "2D" :
        type_request = "vc_con"
    elif hours == "0.5D" or hours == "4H":
        type_request = "half_day"
    else:
        type_request = "hour"

    check_number_of_days_cancel(number_before_cancel,number_after_cancel,hours_used,vc_name,oneday,use_hour_unit,type_request)

def CancelRequest(oneday,use_hour_unit):
    #try:
    number_before_cancel = available_vacation()
    commons.ClickElementWithText("My Vacation Status")
    time.sleep(3)
    commons.Title("Cancel Request")
    if commons.IsDisplayedByXpath(pr_rq.rq_vc["check_list_re"]) == True :
        commons.WriteOnExcel(pr_rq.NoRequestToCancel[Pass])
       
    else:
        List_Request  = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
        Total_Request = commons.TotalData(List_Request)

        # Get all requests that can be use for cancel #
        Result_Find , Result_Cancel , Status_Before_Cancel , Request_To_Cancel = ChooseVacationToCancel(Total_Request)
        
        # If the request is canceled successfully, check the request has changed status #
        if  Result_Cancel == True:
            Check_Number ,Infor_Request = CheckStatusAfterCancel(Total_Request,Request_To_Cancel,Status_Before_Cancel)

            if  Check_Number == True :
                CheckNumberAfterCancel(Infor_Request,number_before_cancel,oneday,use_hour_unit)
        
        if  Result_Find == False:
            commons.WriteOnExcel(pr_rq.NoRequestToCancel[Pass])
    #except:
        commons.ClickElementWithText("My Vacation Status")



def check_number_of_days_cancel(before,after,hour_use,vc_name,oneday,use_hour_unit,type_request):
    commons.Title("Available Vacation")
    total = used = remain = True 

    # Check number of days before request #
    for vacation in before:
        if vacation["vacation_name"] == vc_name :
            vc_bf_use       = vacation
            infor_before    = infor(vc_bf_use,"Info Vacation before cancel",type_request)
            msg_cancel      = type_vc.MsgBeforeCancel(type_request)
            HourToDaysUsed  = xpath.HourToDaysUsed(vc_bf_use,oneday,hour_use,use_hour_unit,type_request)

            if vc_bf_use["total"] != "-" :
                days = change_hour_to_day(**HourToDaysUsed)
                result_before(vc_bf_use["total"],days,msg_cancel)
            else :
                result_before(vc_bf_use["total"],"-",msg_cancel)
            break

    # Check number of days after request #
    for vacation in after: 
        if vacation["vacation_name"] == vc_name :
            vc_af_use       = vacation
            infor_after     = infor(vc_af_use,"Info Vacation after cancel",type_request)
            msg_total       = type_vc.MsgCancelTotal(type_request)
            msg_used        = type_vc.MsgCancelUsed(type_request)
            msg_remain      = type_vc.MsgCancelRemain(type_request)
            msg_cancel      = type_vc.MsgAfterRequest(type_request)
            UHourToDaysNone = xpath.UCHourToDaysNone(vc_bf_use,oneday,hour_use,use_hour_unit,type_request)
            RHourToDaysNone = xpath.RCHourToDaysNone(vc_bf_use,oneday,hour_use,use_hour_unit,type_request)

            if vc_bf_use["total"] != "-" : 
                total  = result_number(vc_af_use["total"],vc_bf_use["total"],msg_total)
            
                used_before_plus_used = change_hour_to_day(**UHourToDaysNone)
                used   = result_number(vc_af_use["used"],used_before_plus_used,msg_used)

                remain_before_minus_used = change_hour_to_day(**RHourToDaysNone)
                remain = result_number(vc_af_use["remain"],remain_before_minus_used,msg_remain)
                
                result_after(total,used,remain,msg_cancel)

            else :
                total = result_number(vc_af_use["total"],"-",msg_total)
                
                used_before_plus_used = change_hour_to_day(**UHourToDaysNone)
                used = result_number(vc_af_use["used"],used_before_plus_used,msg_used)
                remain = result_number(vc_af_use["remain"],"-",msg_remain)

                result_after(total,used,remain,msg_cancel)
            break

    commons.Content(infor_before)
    commons.Content(infor_after)

def select_approver():

    commons.Scroll()
    bt_select_approver = commons.IsDisplayedByXpath(pr_rq.rq_vc["bt_select_approver"])
    select_approver    = xpath.ParSelectApprover()
    if bt_select_approver == True:
        # Select approver from approver list #
        commons.ClickElementWithXpath(pr_rq.rq_vc["bt_select_approver"])
        commons.ClickElementWithXpath(pr_rq.rq_vc["delete_all"])
        time.sleep(3)

        check_no_approver = driver.find_elements_by_xpath(pr_rq.rq_vc["text_list_ap"])
        if approver_list(check_no_approver) == True :
            
            selected_approver = click_approver()
            
            if selected_approver["is_selected"] == True :
                if add_approver() == True :
                    save_approver(selected_approver["user_name"],select_approver)
        return select_approver

    else:
        # Use approver line #
        if commons.IsDisplayedByXpath(pr_rq.rq_vc["bt_quick_approver"]) == True: 
            list_approver = []
            commons.CasePass(pr_rq.ApproverLine)
            select_approver["result_approver"] = True
            select_approver["approval_line"]   = True
            list_app      = driver.find_elements_by_xpath(pr_rq.rq_vc["list_approver1"])
            total_app     = commons.TotalData(list_app)
            select_approver["approver_name"] = approver_name(total_app,list_approver)
            return select_approver 

        else:
            # Approver is approval exception #
            commons.CasePass(pr_rq.ApproverException)
            select_approver["result_approver"]    = True
            select_approver["approval_exception"] = True
            return select_approver

def function_search():
    try:
        user_name = "TS2"
        result_select_user = {
            "search"    :"False",
            "select_ogr":"False"
            }
        driver.find_element_by_xpath( pr_rq.rq_vc["bt_select_cc"]).click()
        if commons.IsDisplayedByXpath(pr_rq.rq_vc["bt_add_cc"]) == True:
            driver.find_element_by_xpath( pr_rq.rq_vc["dele_all_cc"]).click()

            # Search user # 
            list_departmaent = driver.find_elements_by_xpath( pr_rq.rq_vc["org_search"])
            before_search    = len(list_departmaent)
            firt_department  = commons.GetText(pr_rq.rq_vc["firt_depart"])
            ip_search_user   = driver.find_element_by_xpath( pr_rq.rq_vc["search"])
            driver.implicitly_wait(5)
            ip_search_user.click()
            ip_search_user.send_keys(user_name)
            ip_search_user.send_keys(Keys.RETURN)
            
            if ip_search_user.get_attribute('value') == user_name :
                commons.WriteOnExcel(pr_rq.InputUser[Pass])
                
                time.sleep(1)
                cc = commons.GetText(pr_rq.rq_vc["cc_namea"])
                if cc == "No data." :
                    commons.WriteOnExcel(pr_rq.Search[Pass])
                    
                else:
                    time.sleep(2)
                    list_depart  = driver.find_elements_by_xpath(pr_rq.rq_vc["org_search"])
                    after_search = len(list_depart)
                    if before_search != after_search :
                        commons.WriteOnExcel(pr_rq.Search[Pass])
                        result_select_user["search"] = True
                    else:
                        firt_department1 = commons.GetText(pr_rq.rq_vc["firt_depart"])
                        if firt_department == firt_department1:
                            commons.WriteOnExcel(pr_rq.Search[Fail])
                           
                        else:
                            commons.WriteOnExcel(pr_rq.Search[Pass])
                            result_select_user["search"] = True      
            else:
                commons.WriteOnExcel(pr_rq.InputUser[Fail])

            # Select user from Org # 
            ip_search_user.clear()
            driver.find_element_by_xpath( pr_rq.rq_vc["bt_save"]).click()
            driver.find_element_by_xpath( pr_rq.rq_vc["bt_select_cc"]).click()
            selected_cc = select_user_from_depart() 
            if selected_cc != False :
                commons.WriteOnExcel(pr_rq.SelectUser[Pass])
                result_select_user["select_ogr"] = True      
            else:
                commons.WriteOnExcel(pr_rq.SelectUser[Fail])
            driver.find_element_by_xpath( pr_rq.rq_vc["bt_save"]).click()
        
    except:
        commons.ClickElementWithText("My Vacation Status")

def select_cc_enter_reason():
    # SELECT CC #
    time.sleep(3)
    commons.Scroll()
    
    commons.Title("Select CC")
    time.sleep(3)
    driver.find_element_by_xpath( pr_rq.rq_vc["bt_select_cc"]).click()
    if add_cc() == True :
        selected_cc = select_user_from_depart().replace("(", "").replace(")", "")
        if  SelectedCc(selected_cc) == True :
            check_saved_cc(selected_cc)

    # ADD REASON #
    commons.Scroll()
    commons.Title("Enter Reason")
    if commons.IsDisplayedByXpath( pr_rq.rq_vc["reason"]) == True:
        reason =driver.find_element_by_xpath( pr_rq.rq_vc["reason"])
        #reason.click()
        reason.send_keys(pr_rq.rq_vc["reason_text"])
        if reason.get_attribute('value') ==  pr_rq.rq_vc["reason_text"]:
            commons.CasePass(pr_rq.EnterReason[Pass]["description"])
        else:
            commons.WriteOnExcel(pr_rq.EnterReason[Fail]) 
    else:
        commons.WriteOnExcel(pr_rq.EnterReason[Pass])
       
    
def check_use_hour_unit_half_day(total_vc):
    
    # Choose vacation name to request  #
    use_hour_unit      = False 
    all_vacation       = [] 
    available_vacation = {"available_vacation":""}
    
    if total_vc  == 0 :
        available_vacation["available_vacation"] = 0
        all_vacation.append(available_vacation)

    else:
        i = 1
        available_vacation["available_vacation"] = total_vc
        all_vacation.append(available_vacation)
        while i <= total_vc:
            time.sleep(1)
            usage_settings = xpath.ParUsageSettings()
            driver.find_element_by_css_selector(pr_rq.rq_vc["select_vacation"]).click()
            driver.find_element_by_xpath(data["vc_name"] + str(i) +"]").click()
            vacation = vacation_use_for_request()
            usage_settings["vacation_name"]   = vacation["vacation_name"]
            usage_settings["number_of_days"]  = vacation["number_of_days"]
            usage_settings["number_of_hours"] = vacation["number_of_hours"]

            if commons.IsDisplayedByXpath(pr_rq.rq_vc["hour_unit"]) == True :
                use_hour_unit = True
                usage_settings["use_hour_unit"] = True
            else:
                usage_settings["use_hour_unit"] = False

            if commons.IsDisplayedByXpath(pr_rq.rq_vc["radi_am"]) == True:
                usage_settings["use_half_day"] = True
                if use_hour_unit == True:
                    usage_settings["hour_use"] = hour_used(use_hour_unit,"am")
            else:
                usage_settings["use_half_day"] = False

            all_vacation.append(usage_settings)
            i = i+1
    
    return all_vacation    
            
def select_vacation_use_hour_unit_half_day(total_vc,list_vc_use_half,hour_use,type_vc):
    i = 1
    while i <= total_vc:
        time.sleep(1)
        driver.find_element_by_css_selector( pr_rq.rq_vc["select_vacation"]).click()
        driver.find_element_by_xpath(data["vc_name"] + str(i) +"]").click()
        vacation_name = commons.GetText(pr_rq.rq_vc["vacation_name"])
        vacation_name = xpath.UsChange(vacation_name)
        for vacation in list_vc_use_half:
            if vacation["vacation_name"] == vacation_name :
                if type_vc =="hour":
                    if  float(vacation["number_of_hours"]) >= 1       or \
                        float(vacation["number_of_days"] ) >= hour_use:
                        return vacation_name
                else:
                    if  float(vacation["number_of_hours"]) >= hour_use or \
                        float(vacation["number_of_days"] ) >= hour_use:
                        return vacation_name
        i = i+1      
                        
    return False

def check_result_request():
    try :
        notification = driver.execute_script(data["result"])
        content_notification = notification.split("\n")
        if content_notification[0] =="success":
            return Pass
        else:
            return "noti_error" + content_notification[1]
    except:
        time.sleep(1)
        if commons.IsDisplayedByXpath(data["my_vt"]["vc_history"]) == True:
            return Pass
        else:
            return Fail
    
def CheckCreatedRequest(info_vc,type_request,use_hour_unit,approver):
    
    i      = 1
    result = False
    time.sleep(3)
    try:
        driver.find_element_by_xpath( pr_rq.rq_vc["bt_refresh"]).click()
        List_Request   = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
        Total_Request  = commons.TotalData(List_Request)
        MsgDisplayed   = type_vc.MsgDisplayedRequest(type_request)
        if Total_Request >= 1 :
            while i <= Total_Request:
                vc_rq = xpath.ParVacationRequest()
                if  approver["approval_exception"] == True :
                    vc_rq["status"]   = "Approved"
                    info_vc["status"] = "Approved"

                vc_rq = vacation_request(vc_rq,i)
                if  info_vc["vc_name"] == vc_rq["vc_name"]           and \
                    info_vc["vc_date"] == vc_rq["vc_date"]           and \
                    info_vc["status"]  == vc_rq["status"]            and \
                    info_vc["request_date"] == vc_rq["request_date"]     :

                    result = True
                    commons.WriteOnExcel(MsgDisplayed[Pass])
                    ViewDetailRequest(vc_rq,type_request,use_hour_unit,approver,i)
                    break

                i = i+1

            if result == False:
                commons.WriteOnExcel(MsgDisplayed[Fail])
        else: 
            commons.WriteOnExcel(MsgDisplayed[Fail])
            
    except:
        commons.ClickElementWithText("My Vacation Status")

    return result

def ViewDetailRequest(vc_rq,type_request,use_hour_unit,approver,i):
    commons.Title("View Detail")
    print(data["rq_vc"]["ic_detail"] % (str(i) , str(i-1)))
    commons.ClickElementWithXpath(data["rq_vc"]["ic_detail"] % (str(i) , str(i-1)))
    time.sleep(2)
    Info_Before = xpath.ParInfo()
    Info_After  = xpath.ParInfo()

    # View vacation date #
    Result_Date = view_vacation_date(Info_Before,vc_rq,type_request,Info_After)
    
    # View detail Number of days used #
    Result_Use = view_detail_number_of_days_used(Info_Before,use_hour_unit,type_request,Info_After)
    
    # View detail request date #
    Result_Re_Date = view_detail_request_date(Info_Before,vc_rq,type_request,Info_After)
    
    # View detail approver and reason #
    Result_Resaon , Result_Approver , Info_Before , Info_After = view_detail_approver_and_reason(type_request,approver,Info_After,Info_Before)

    MsgViewDetail = type_vc.MsgViewDetail(type_request)
   
    if type_request == "all":
        if  Result_Date       == True and \
            Result_Use        == True and \
            Result_Re_Date    == True and \
            Result_Resaon     == True and \
            Result_Approver   == True     :
            commons.WriteOnExcel(MsgViewDetail[Pass])
        else:
            commons.WriteOnExcel(MsgViewDetail[Fail])
    else:
        if  Result_Date     == True and \
            Result_Use      == True and \
            Result_Re_Date  == True     :
            commons.WriteOnExcel(MsgViewDetail[Pass])
        else:
            commons.WriteOnExcel(MsgViewDetail[Fail])
            


    infor_detail("Information entered",Info_Before)
    infor_detail("Information saved  ",Info_After )
    
def TimeComparison(request_date,today):

    if request_date.rfind("~") > 0 :
        request_date = request_date[int(request_date.rfind("~")) + 1 : None]

    Request_Date , Today = xpath.Year(today,request_date)
    Request_Date         = datetime.datetime.strptime(Request_Date, "%y/%m/%d")
    Today                = datetime.datetime.strptime(Today, "%y/%m/%d")

    if Request_Date < Today :
        return False  
    else:
        return True

def select_hour_use_hour_unit():
    
    commons.Title("Select Hour")
    driver.find_element_by_xpath( pr_rq.rq_vc["hour_unit"]).click()
    driver.find_element_by_xpath( pr_rq.rq_vc["hour_start"]).click()
    
    start_options = driver.find_elements_by_xpath( pr_rq.rq_vc["start_option"])
    if len(start_options) > 1:
        driver.find_element_by_xpath( pr_rq.rq_vc["sl_hour_start"]).click()
    else:
        commons.WriteOnExcel(pr_rq.HourStart[Pass])
        return False

    driver.find_element_by_xpath( pr_rq.rq_vc["hour_end"]).click()
    end_options = driver.find_elements_by_xpath( pr_rq.rq_vc["end_option"])
    if len(end_options) > 1:
        driver.find_element_by_xpath( pr_rq.rq_vc["sl_hour_end"]).click()
    else:
        commons.WriteOnExcel(pr_rq.HourSEnd[Pass])
        return False
    
    hour_selected = commons.GetText(pr_rq.rq_vc["selected_date"])
    start         = int(hour_selected.rfind("("))
    end           = int(hour_selected.rfind(")")) + 1
    hour_selected = hour_selected[start : end]
    return hour_selected
    
def info_request_list(i):
    
    infor_request       = xpath.ParInforRequest()
    infor_request["no"] = str(i)
    Possition_Icon = commons.IsDisplayedByXpath(data["ic_before"])
    vacation_request["vc_name"]    = xpath.LiVacation(i,"na",Possition_Icon)
    infor_request["vacation_date"] = xpath.LiRequest(i,"da")
    infor_request["use"]           = xpath.LiRequest(i,"us")
    infor_request["request_date"]  = xpath.LiRequest(i,"rd")
    infor_request["status"]        = xpath.LiRequest(i,"st")
    end                            = int(infor_request["vacation_date"].rfind("\n"))
    infor_request["vacation_time"] = infor_request["vacation_date"][None: end]
    return infor_request

def count_all_vacation_request():
    # Get all status from list request vacation #
    
    i = 1
    total_request = 0
    time.sleep(3)

    if commons.IsDisplayedByXpath( pr_rq.rq_vc["check_list_re"]) == False : 
        driver.find_element_by_xpath(data["mn_pro"]["ic_to_end_page"]).click()
        end_page_text = commons.GetText(data["mn_pro"]["page_current"])
        end_page = int(end_page_text)
        driver.find_element_by_xpath(data["mn_pro"]["ic_to_first_page"]).click()
        
        while i <= end_page:
            if i == end_page :
                time.sleep(3)
                total_re = driver.find_elements_by_xpath(data["mn_pro"]["list_re_vc"])
                total_request = total_request + commons.TotalData(total_re)
            else:
                total_request = total_request + 20
            i = i+1
    return total_request

def two_requests_are_the_same(request1,request2):
    if  request1["vacation_name"] == request2["vacation_name"] and \
        request1["vacation_date"] == request2["vacation_date"] and \
        request1["use"]           == request2["use"]           and \
        request1["request_date"]  == request2["request_date"]:
        return True
    else: 
        return False
            
    
def vacation_displayed_in_time_card(date_request,approver):
    try :
        i = 1
        current_date  = datetime.date.today()
        month_request = date_request.month
        current_month = current_date.month
        number_of_clicks = int(month_request - current_month)

        commons.Title("Time Card")
        driver.find_element_by_xpath(data["menu_tc"]).click()
        commons.ClickElementWithText("Timesheets")
        
        if commons.IsDisplayedByXpath(data["tab_calen_tc"]) == True :

            if commons.IsDisplayedByXpath(data["no_work_tc"]) == False :
                commons.WriteOnExcel(pr_rq.NoWorkPolicy[Pass])
            else:
                time.sleep(2)
                driver.find_element_by_css_selector(data["date_tc"]).click()
                if number_of_clicks == 0 :
                    click_date_time_card(date_request)
                else:
                    while i <= number_of_clicks :
                        driver.find_element_by_css_selector(data["ic_next_tc"]).click()
                        click_date_time_card(date_request)

                if approver["approval_exception"] == True :
                    if  commons.IsDisplayedByXpath(data["row_vaca_tc"]) == True :
                        commons.WriteOnExcel(pr_rq.DisplayedVacation[Pass])
                    else:
                        commons.WriteOnExcel(pr_rq.DisplayedVacation[Fail])
                else :
                    commons.WriteOnExcel(pr_rq.NoApproval[Pass])

        else:
            commons.WriteOnExcel(pr_rq.AccessTimeSheets[Fail])

        driver.find_element_by_xpath(data["menu_vc"]).click()
    except:
        driver.find_element_by_xpath(data["menu_vc"]).click()

def time_clockin():
    try :
        time_clock_in =  False
        driver.find_element_by_xpath(data["menu_tc"]).click()
        commons.ClickElementWithText("Timesheets")
        if commons.IsDisplayedByXpath(data["tab_calen_tc"]) == True :
            if commons.IsDisplayedByXpath(data["no_work_tc"]) == True :
                clock_in = commons.GetText(data["time_clock_in"])
                if clock_in.rfind("00") > 0 :
                    time_clock_in = clock_in

                driver.find_element_by_xpath(data["menu_vc"]).click()
                return time_clock_in           
    except:
        driver.find_element_by_xpath(data["menu_vc"]).click()

def collect_clock_in_from_time_card(time_clock_in,type_request,request_date):
    if time_clock_in !=  False :
        clock_in = time_clock_in[None: int(time_clock_in.rfind("("))]
        #clock_in_hour = time_clock_in[None: int(time_clock_in.rfind(":"))]
        clock_in_hour = time_clock_in[None: 2]

        if type_request == "hour":
            time_end = int(clock_in_hour) + 2
            vacation_date = data["tc"]["date_2h"] % (str(request_date) , clock_in , str(time_end))
           
        else:
            time_end = int(clock_in_hour) + 4 
            vacation_date = data["tc"]["date_4h"] % (str(request_date) , clock_in , str(time_end))
           
        return vacation_date

def information_vacation(vacation_request):
    vacation_name = data["info"]["name"]    % vacation_request["vc_name"]
    vacation_date = data["info"]["date"]    % vacation_request["vc_date"]
    request_date  = data["info"]["request"] % vacation_request["request_date"]
    title         = data["info"]["title"]
    content       = data["info"]["data"]    % (title , vacation_name , vacation_date , request_date)
    commons.Content(content)

def update_status_for_request(approver,vacation_request):
    if approver["approval_exception"] == True :
        vacation_request["status"] = "Approved"
    return vacation_request

def click_on_button_to_request():
    driver.find_element_by_xpath( pr_rq.rq_vc["bt_request_be"]).click()
    driver.find_element_by_css_selector( pr_rq.rq_vc["bt_request_af"]).click()

def DateToRequest(VcDate,Vacation_Request,Result_Date,type):

    if VcDate != False: 
        if type != "vc_con":
            Vacation_Request["vc_date"] = data["ty"][type] % str(VcDate)
        else :
            Vacation_Request["vc_date"] = data["ty"]["lt"] % (str(VcDate[0]) , str(VcDate[1]))
    else:
        Result_Date = False
    
    return  Result_Date , Vacation_Request


def choose_date_to_request(Vacation_Request,type,Date_Used):
    Result_Date = True
    ToDay       = str(datetime.date.today())

    if type == "all":
        VcDate = select_date_to_request_leave()
        Result_Date , Vacation_Request = DateToRequest(VcDate,Vacation_Request,Result_Date,type)
        
    elif type == "am" :
        Vacation_Request["request_date"] = ToDay
        VcDate = select_date_to_request_leave()
        Result_Date , Vacation_Request = DateToRequest(VcDate,Vacation_Request,Result_Date,type)

    elif type == "pm" :
        Vacation_Request["request_date"] = ToDay
        VcDate = select_date_to_request_leave()
        Result_Date , Vacation_Request = DateToRequest(VcDate,Vacation_Request,Result_Date,type)
    
    elif type == "hour" :
        Vacation_Request["request_date"] = ToDay
        VcDate =  select_date_to_request_leave()
        Vacation_Request["vc_date"] = VcDate
        if VcDate == False:
            Result_Date = False

    else :
        VcDate = select_date_to_request_leave_for_vacation_consecutive(Date_Used)
        Result_Date , Vacation_Request = DateToRequest(VcDate,Vacation_Request,Result_Date,type)
    

    choose_date = xpath.ParChooseDate(Result_Date,VcDate,Vacation_Request)
    return choose_date

def TypeOfRequest(Type):
    if   Type == "all" :
        commons.Title("Request Vacation : All Day ")
    elif Type == "vc_con" :
        commons.Title("Request Vacation : Vacation Consecutive ")


def choose_vacation_to_request(Vacation_Request,type):
    commons.Title("Vacation Name")
    Result_Select_Vc = False
    Use_Hour_Unit    = False
    Total_Vc         = total_vacation()
    Infor_Vacation   = check_use_hour_unit_half_day(Total_Vc)
    
    # Check there are any vacations using hour unit #
    # Change the way to check the number of leave days #
    i = 1
    while i <= Total_Vc:
        if  Infor_Vacation[i]["use_hour_unit"] == True:
            Use_Hour_Unit = True
            break
        i = i + 1

    # Check the remaining days of each vacation are enough to request #
    i = 1
    while i <= Total_Vc:
        driver.find_element_by_css_selector(pr_rq.rq_vc["select_vacation"]).click()
        driver.find_element_by_xpath(data["vc_name"] + str(i) +"]").click()
        Vc_Use_For_Request = vacation_use_for_request()
        Hour_Use           = hour_used(Use_Hour_Unit,type)
        Remain_Days        = Vc_Use_For_Request["number_of_days"]
        Remain_Hours       = Vc_Use_For_Request["number_of_hours"]
        Vc_Name            = Vc_Use_For_Request["vacation_name"]

        # If vacation is other vacation can use this vacation to request #
        if float(Remain_Days) == 0 and float(Remain_Hours) == 0:
            Result_Select_Vc =  True
            Vacation_Request["vc_name"] = Vc_Name.replace(" ", "")
            commons.CasePass(pr_rq.SelectVacationName)
            commons.CasePass(data["vc_na"] % Vc_Name )
            break

        # If vacation is grant/regular , need to check there are enough days to request #
        else :
            if float(Remain_Days) >= Hour_Use :
                Result_Select_Vc = True
                Vacation_Request["vc_name"] = Vc_Name.replace(" ", "")
                commons.CasePass(pr_rq.SelectVacationName)
                commons.CasePass(data["vc_na"] % Vc_Name )
                break 
        i = i + 1
    # Vacation does not have enough days to choose #
    if Result_Select_Vc == False:
        commons.CasePass(pr_rq.NoVacationToRequest)
    

    Par_Choose_Vacation = {
        "hour_use"         :Hour_Use,
        "result_select_vc" :Result_Select_Vc,
        "use_hour_unit"    :Use_Hour_Unit,
        "vacation_request" :Vacation_Request,
        "vc_name"          :Vc_Name
    }
    
    return Par_Choose_Vacation

def RequestAllDay(Approver,ResultApprover):
    commons.Title("Request Vacation : All Day ")
    type                 = "all"
    RequestDate          = str(datetime.date.today())
    Date_Used             = GetDateUsed()
    VacationRequest      = xpath.par_vacation_request(RequestDate)
    VacationRequest      = update_status_for_request(Approver,VacationRequest)
    SelectVacation       = choose_vacation_to_request(VacationRequest,type)
    ChooseDate           = choose_date_to_request(VacationRequest,type,Date_Used)
    Oneday               = hours_set_from_time_card(type)
    NumberBeforeRequest  = available_vacation()
    ParCreateRequest     = xpath.ParamCreateRequest(VacationRequest,SelectVacation,Approver,type)
    ParNumberOfDays      = xpath.ParNumberOfDays(NumberBeforeRequest,type,SelectVacation,Oneday)

    select_approver()
    select_cc_enter_reason()
    
    commons.Title("Send Request")
    if  bool(SelectVacation["result_select_vc"])  == True and \
        bool(ResultApprover)                      == True and \
        bool(ChooseDate["result_date"])           == True     :

        click_on_button_to_request()
        ResultRequest     = check_result_request()
        if  ResultRequest == Pass :
            information_vacation(VacationRequest)
            commons.WriteOnExcel(pr_rq.RequestVacationAll[Pass])
            requested = CheckCreatedRequest(**ParCreateRequest)
            if requested == True :
                NumberAfterRequest = available_vacation()
                ParNumberOfDays    = commons.AddData(ParNumberOfDays,NumberAfterRequest)
                ParCancelRequest   = xpath.ParCancelRequest(SelectVacation,Oneday)
                check_number_of_days_off(**ParNumberOfDays)
                CancelRequest(**ParCancelRequest)

        elif ResultRequest == Fail:
            commons.WriteOnExcel(pr_rq.RequestVacationAll[Pass])

        else :
            commons.WriteOnExcel(pr_rq.RequestVacationAll[Fail])
    
def RequestVacationConsecutive(Approver,ResultApprover):
    commons.Title("Request Vacation : Vacation Consecutive ")
    type                  = "vc_con"
    RequestDate           = str(datetime.date.today())
    Date_Used             = GetDateUsed()
    VacationRequest       = xpath.par_vacation_request(RequestDate)
    VacationRequest       = update_status_for_request(Approver,VacationRequest)
    SelectVacation        = choose_vacation_to_request(VacationRequest,type)
    ChooseDate            = choose_date_to_request(VacationRequest,type,Date_Used)
    Oneday                = hours_set_from_time_card(type)
    NumberBeforeRequest   = available_vacation()
    ParCreateRequest      = xpath.ParamCreateRequest(VacationRequest,SelectVacation,Approver,type)
    ParNumberOfDays       = xpath.ParNumberOfDays(NumberBeforeRequest,type,SelectVacation,Oneday)
    
    commons.Title("Send Request")
    if  bool(SelectVacation["result_select_vc"])  == True and \
        bool(ResultApprover)                      == True and \
        bool(ChooseDate["result_date"])           == True     :
        
        click_on_button_to_request()
        ResultRequest      = check_result_request()
        if  ResultRequest  == Pass :
            information_vacation(VacationRequest)
            commons.WriteOnExcel(pr_rq.VacationConsecutive[Pass])
            requested = CheckCreatedRequest(**ParCreateRequest)
            if requested == True : 
                NumberAfterRequest = available_vacation()
                ParNumberOfDays    = commons.AddData(ParNumberOfDays,NumberAfterRequest)
                ParCancelRequest   = xpath.ParCancelRequest(SelectVacation,Oneday)
                check_number_of_days_off(**ParNumberOfDays)
                CancelRequest(**ParCancelRequest)

        elif ResultRequest == Fail:
            commons.WriteOnExcel(pr_rq.VacationConsecutive[Pass])
    else :
        commons.WriteOnExcel(pr_rq.VacationConsecutive[Fail])
    
def RequestVacation():
    commons.Title("REQUEST VACATION")
    commons.Title("Select approver")

    TypeApprover = False
    TypeApprover , Approver = ResultApprover()
    commons.ClickElementWithText("Request Vacation")
    TotalVc = total_vacation()
    
    if TotalVc == 0 :
        commons.WriteOnExcel(pr_rq.NoVacationToRequest[Pass])
    else :
       
        RequestAllDay(Approver,TypeApprover)
        RequestVacationConsecutive(Approver,TypeApprover)
       
       
































'''

def submenu_request_vacation():
    vacation_my_functions.request_and_cancel_vacation()
        
def submenu_my_vacation_status():
    
    commons.ClickLinkText("My Vacation Status")
    if commons.IsDisplayedByXpath(pr_rq.rq_vc["vc_history"]) == True : 
        commons.WriteOnExcel(pr_rq.AccessSubmenuMy[Pass])
       
        commons.Title("Delete Request")
        vacation_my_functions.delete_request_my_vc()
       
       
        commons.Title("Filter status")
        vacation_my_functions.filter_status("my") 
        
       
        commons.Title("Tab Information")
        commons.ClickElementWithXpath(pr_rq.rq_vc["tab_information"])
       
        if  commons.IsDisplayedByXpath(pr_rq.rq_vc["text_infor"]) == True : 
            commons.WriteOnExcel(pr_rq.AccessTabInfor[Pass])
        else:
            commons.WriteOnExcel(pr_rq.AccessTabInfor[Fail])
            
        
        commons.Title("Tab Request Status")
        commons.ClickElementWithXpath(pr_rq.rq_vc["tab_re_status"])
       
        if commons.IsDisplayedByXpath(pr_rq.rq_vc["text_re_status"]) == True : 
            commons.WriteOnExcel(pr_rq.AccessTabStatus[Pass])
        else:
            commons.WriteOnExcel(pr_rq.AccessTabStatus[Fail])
        
        
        commons.Title("Tab Vacation History")
        commons.ClickElementWithXpath(pr_rq.rq_vc["tab_vc_history"])
       
        if commons.IsDisplayedByXpath(pr_rq.rq_vc["text_vc_history"]) == True : 
            commons.WriteOnExcel(pr_rq.AccessTabHistory[Pass])

            commons.Title("View detail")
            view.view_detail_request_at_calendar(pr_rq.re_history)
        else:
            commons.WriteOnExcel(pr_rq.AccessTabHistory[Fail])
        
    else:
        commons.WriteOnExcel(pr_rq.AccessSubmenuMy[Pass])

def submenu_vacation_schedule():
    
    commons.Title("III.VACATION SCHEDULE")
    if  commons.IsDisplayedByTextLink("Vacation Schedule") == True:
        driver.find_element_by_link_text("Vacation Schedule").click()

        if commons.IsDisplayedByXpath(pr_rq.rq_vc["vc_schedule"]) == True : 
            commons.WriteOnExcel(pr_rq.AccessTabSchedule[Pass])
            commons.Title("View detail")
            view.view_detail_request_at_calendar(pr_rq.re_detail_depart)
        else:
            commons.WriteOnExcel(pr_rq.AccessTabSchedule[Fail])
    else:
        if commons.IsDisplayedByTextLink("My Dept Vacation") == True:
            driver.find_element_by_link_text("My Dept Vacation").click()

            if commons.IsDisplayedByXpath(pr_rq.rq_vc["check_my_depart"]) == True : 
                commons.WriteOnExcel(pr_rq.AccessTabDepartment[Pass])
                commons.Title("View detail")
                view.view_detail_request_at_calendar(pr_rq.re_detail_depart)
            else:
                commons.WriteOnExcel(pr_rq.AccessTabDepartment[Fail])

def submenu_view_cc():
    
    commons.Title("III.VIEW CC")
    driver.find_element_by_link_text("View CC").click()

    if commons.IsDisplayedByXpath(pr_rq.rq_vc["view_cc"]) == True : 
        commons.WriteOnExcel(pr_rq.AccessSubmenuCc[Pass])
        commons.Title("Filter")
        vacation_my_functions.filter_status("cc")
        
        commons.Title("View detail")
        vacation_my_functions.view_detail_request_cc()
    else:
        commons.WriteOnExcel(pr_rq.AccessSubmenuCc[Fail])

def submenu_request_settlement():
    
    commons.Title("IV.REQUEST SETTLEMENT")
    if commons.IsDisplayedByTextLink("Request Settlement") == True:
        driver.find_element_by_link_text("Request Settlement").click()

        if commons.IsDisplayedByXpath(pr_rq.rq_vc["text_sub_re_settlement"]) == True : 
            commons.WriteOnExcel(pr_rq.AccessSubmenuSettlement[Pass])
            driver.implicitly_wait(5)

            commons.Title("Tab Request History")
            commons.ClickElementWithXpath(pr_rq.rq_vc["tab_re_history"])
            

            if commons.IsDisplayedByXpath(pr_rq.rq_vc["text_re_status"]) == True : 
                commons.WriteOnExcel(pr_rq.AccessTabSettlementHistory[Pass])
                driver.implicitly_wait(5)
            else:
                commons.WriteOnExcel(pr_rq.AccessTabSettlementHistory[Fail])
            
            commons.Title("Tab Request")
            commons.ClickElementWithXpath(pr_rq.rq_vc["tab_re_re"])
           
            if commons.IsDisplayedByXpath(pr_rq.rq_vc["text_re_re"]) == True : 
                commons.WriteOnExcel(pr_rq.AccessTabSettlementRequest[Pass])
                driver.implicitly_wait(5)
            else:
                commons.WriteOnExcel(pr_rq.AccessTabSettlementRequest[Fail])
        else:
            commons.WriteOnExcel(pr_rq.AccessSubmenuSettlement[Fail])  
    else:
        commons.WriteOnExcel(pr_rq.AccessNoSubmenuSettlement[Pass])


def request():
   
    submenu_request_vacation()
    submenu_my_vacation_status()
    submenu_vacation_schedule()
    submenu_view_cc()
    submenu_request_settlement()
    
'''  
    

    
    