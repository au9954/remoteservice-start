import wmi

#for local machine
#conn=wmi.WMI()

#for remote connection
conn=wmi.WMI("IP/Hostname",user="username",password="password")

#to find wmi_class for specific_keyword
##keyword=input("Enter keyword to find relevant wmi class: ")
##try:
##    for i in conn.classes:
##        if keyword in i:
##            print(i)
##except:
##    print("Oops!No such class")

#to get specific properties/methods of the wmi classess
#for i in conn.win32_service.methods.keys():
#    print("Method: ",i)
##
##for j in conn.win32_service.properties.keys():
##    print("Properties: ",j)


#to display specific services

for var in conn.win32_service():
#if this is the display name of the service
#windows event log service
    if(var.DisplayName=="Windows Event Log"):
#if service state is stopped
        if(var.State=="Stopped"):
#exit code 0 means success
            windowseventlogres,=var.StartService()
            if(windowseventlogres==0):
                print("Windows Event Log Started")
            else:print("Windows Event Log Error Code : ",windowseventlogres)
        else:
            print(var.State)
#dhcp service
    if(var.DisplayName=="DHCP Client"):
        if(var.State=="Stopped"):
            dhcpres,=var.StartService()
            if(dhcpres==0):
                print("DHCP service started")
            else:print("DHCP Client Error Code: ",dhcpres)
        else:
            print(var.State)
#tcp/ip netbios helper service
    if(var.DisplayName=="TCP/IP NetBIOS Helper"):
#if service startup type is not Auto,change it to Auto
        if(var.StartMode!="Auto"):
            var.ChangeStartMode(StartMode="Automatic")
        if(var.State=="Stopped"):
            tcpres,=var.StartService()
            if(tcpres==0):
                print("TCP/IP NetBios Helper service started")
            else:print("TCP/IP Error Code :",tcpres)
        else:
            print(var.State)

#microsoft monitoring agent service
    try:
        if(var.DisplayName=="Microsoft Monitoring Agent"):
            if(var.StartMode!="Auto"):
                var.ChangeStartMode(StartMode="Automatic")
            if(var.State=="Stopped"):
                mmares,=var.StartService()
                if(mmares==0):
                    print("MMA service started")
                else:
                    print("MMA Error Code :",mmares)
                    print("Checking for disk space now------------")
                    for variable in conn.win32_logicaldisk():      #if service failed to start check for disk space in D drive
                        if(variable.DeviceID=="D:"):
                            result=variable.FreeSpace              #store the free disk space value in result variable
                            print("Free space in D drive in Mb : ",int(result)/(1024*1024))  #divide by 1024*1024 to convert value into Mb
            else:print(var.State)
            
    except Exception as e:
        print("Yes",e)


