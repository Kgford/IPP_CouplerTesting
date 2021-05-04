mcl_RF_Switch_Controller64.dll - 64 bit COM Object for programmers under 64 bit MS operating system

This DLL file can be used by programmers of Visual Basic,VC++,Delphi,C#
LabView or any other program that recognize 64 bit ActiveX COM DLL file.

mcl_rf_switch_Controller64.dll file should be referenced to the program project.

The DLL file include the following Functions:

1. int Connect(Optional *string SN)  :

SN parameter is needed in case of using more than 1 Switch Box.
SN is the Serial Number of the Switch Box and can be ignored if using only one box.


2. void Disconnect()

Recommanded to Disconnect the device while end the program

3. Int GetSwitchesStatus(int StatusRet)     - Return  the switches status port

4. int Read_ModelName(string ModelName) As Integer
The Model Name returned in ModelName parameter.

5. int Read_SN(string SN ) 
The Serial Number returned in SN parameter.

6. int Set_Switch(string SwitchName , int Val) As Integer

SwitchName - parameter for the required switch can be for example "A" or "B" .
Val - parameter can be 0 for DE-ENERGIZED 1 (or greater) for ENERGIZED.
return 1 for success, 2 for 24Volt Power failure or  0 for software failure

7. int Set_SwitchesPort( byte Val ) As Integer

   Setting the entire Switches port .
   Cal parameter is the port valur required. for example if Val=2 than
   Switch B will be ENERGIZED and all other DE-ENERGIZED.
return 1 for success, 2 for 24Volt Power failure or  0 for software failure


8. int Set_Address(Address As Integer)   
   
   set the address of the unit. the address can be any number between 1 to 255.
   return positive number if success.
   
9. int Get_Address() 

   return the device address.
   if fail - return 0.
   
10. Int Get_Available_SN_List(string SN_List)

string SN_List is returned with all avaliable Switch Boxes connected to USB.
the function return the Number of Switch Boxes.

11. Int Get_Available_Address_List(string Add_List)

string Add_List is returned with all avaliable Switch Boxes connected to USB.
the function return the Number of Switch Boxes.  

12. GetDeviceTemperature(Short TSensor) float

Get the device temperature.
Models with a fan have 2 internal temperature sensors.
Tsensor =1 for sensor number 1 or 2 for sensor no 2.
The methode return the temperature value in celsius.

13. GetHeatAlarm
    Return 1 if the temperature of the switches exceeded 50 celsius deg.
    (usually the temperature should not exceeded this value.)
    else return 0.
14. GetFirmware- return the firmware of the Switch Controller.

15. GetUSBConnectionStatus - return the USB Connection Status.
    
16. short Set_2x10( short P1 , short P2 ) 

    Support 10x2 Switch Matrix :
    P1 and P2 can be any valid number between 0 to 10.
    Switching (Connect)  Port P1 to P1 and Port P2 to P2
    The function return 1 if success or 0 if fail.

17. short Get_2x10( short P1 , short P2 ) 
    
    Support 10x2 Switch Matrix :
    P1 and P2 will be returened with the connection status of Ports P1 and P2
    The function return 1 if success or 0 if fail.   
    
18. short Get_24V_Indicator(  ) 
    
    Support 10x2 Switch Matrix :
    The function return 1 if Power Supply exist otherwise return 0

19. Avaliable for RC Models Firmware C3 and above:   
    long GetSwitchCounter(String sw ) As Long
    sw=ascii of Switch Name: 'A' for Switch A 'B' for Switch B ... 'H' for Switch H
    the function return the Cycle Counter of the specific RF Switch or -1 if fail.

20. Avaliable for RC Models Firmware C3 and above:
    int GetAllSwitchCounters( long swc[] ) 
    swc - address of long type array.
    the function return 1 if success and the swc array will include 8 long values of the  8 Switch Counters.
    if fail the function return 0.    
    

21. int Get_2SP4T_State(String sw)
    this function is related to 2SP4T models
    sw =  Switch Name: "A" for Switch A , "B" for Switch B
    the function return the state value of the switches in case of success
    and -1 in case of fail.
    for example if switch A connect to 4 and sw = "A" then the function will return 4

22. int Set_2SP4T_COM_To(int P1,int P2)
    this function is related to 2SP4T models
    P1 and P2 can recieve values between 0 to 4
    this function connect switch A to port P1 
    and switch B to port P2.

23. int Set_2SP4T_COMA_To(int P)
    this function is related to 2SP4T models
    P  can recieve values between 0 to 4
    this function connect switch A to port P 

24. int Set_2SP4T_COMB_To(int P)
    this function is related to 2SP4T models
    P  can recieve values between 0 to 4
    this function connect switch B to port P 

25. int Get_2SP6T_State(String sw)
    this function is related to 2SP6T models
    sw =  Switch Name: "A" for Switch A , "B" for Switch B
    the function return the state value of the switches in case of success
    and -1 in case of fail.
    for example if switch A connect to 4 and sw = "A" then the function will return 4

26. int Set_2SP6T_COM_To(int P1,int P2)
    this function is related to 2SP6T models
    P1 and P2 can recieve values between 0 to 6
    this function connect switch A to port P1 
    and switch B to port P2.

27. int Set_2SP6T_COMA_To(int P)
    this function is related to 2SP6T models
    P  can recieve values between 0 to 6
    this function connect switch A to port P 

28. int Set_2SP6T_COMB_To(int P)
    this function is related to 2SP6T models
    P  can recieve values between 0 to 6
    this function connect switch B to port P 
    
program Example in VB:

Dim sw As New MCL_RF_Switch_Controller64.USB_RF_Switch,Status as integer
Status=sw.Connect
Status = sw.Set_Switch("A", 1)
sw.Disconnect

 

program Example in Visual C++:

MCL_RF_Switch_Controller64::USB_RF_Switch ^sw = gcnew MCL_RF_Switch_Controller64::USB_RF_Switch();
short Status = 0;
System::String ^SN = "";
System::String ^SW_Name1 = "A";
float ReadResult = 0;
Status = sw->Connect(SN);
Statuse=sw->Set_Switch(SW_Name1, 1)
sw->Disconnect();




      