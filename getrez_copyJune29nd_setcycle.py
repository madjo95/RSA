# -*- coding: cp1252 -*-
# Display Stress Tester(DST) Automation application
# Will detect display count and resolution under a variety of tests(DPMS, reboot, sleep, power cycle)
# MADJO - June 2015


import os
import sys
import time
import thread
import cPickle
import win32api
import win32con
import traceback
import subprocess
import win32com.client
import _winreg as wreg
from win32api import GetSystemMetrics


mon = []
cycles = 0
errors = 0
running = True
num_cycles = int(9999)
dirpath = os.path.dirname(sys.argv[0])
PLINK_PATH = dirpath + "\plink.exe"
RARITAN_HOST = "150.158.46.64"
RARITAN_PASSWORD = "barco123"
RARITAN_USER = "admin"


def ask_yes_no(question):
    """ Ask a yes or no question. """
    response = None
    while response not in ("y", "n"):
        response = raw_input(question).lower()
    return response

def reg_runonce():
    """ Get the path the application was ran from, add reboot argument to path and create Run registry key. """
    path = sys.argv[0]
    arg = "reboot"
    pathname = '%s %s' % (path, arg)
    key = wreg.OpenKey(wreg.HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run",0, wreg.KEY_ALL_ACCESS)
    try:
        wreg.SetValueEx(key, "Barco DST", 0, wreg.REG_SZ, pathname)
    except EnvironmentError:
        print "Encountered problems writing to the Registry..."
    key.Close()

def del_reg_runonce():
    """ Get the path the application was ran from, add reboot argument to path and create Run registry key. """
    path = sys.argv[0]
    arg = "reboot"
    pathname = '%s %s' % (path, arg)
    key = wreg.OpenKey(wreg.HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run",0, wreg.KEY_ALL_ACCESS)
    try:
        wreg.DeleteValue(key, "Barco DST")
    except EnvironmentError:
        pass
        # print "Encountered problems deleting the Registry key..."
    key.Close()
    
def click(x,y):
    """ Move the mouse cursor """
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE,x,y,0,0)

def clr_tmp():
    dirpath = os.path.dirname(sys.argv[0])
    pickle_path = dirpath + "/dst_tmp.dat"
    pickle_file = open(pickle_path, "w+")
    pickle_file.close()

def get_port_number(question):
    response = None
    while response not in ("1", "2", "3", "4", "5", "6", "7", "8"):
        response = raw_input(question)
    return response

def RunRaritanCmd(cmd, port, onoff):
    pdu_cmd = " set /system1/outlet%s powerState=%s" % (port, onoff)
    full_cmd = "%s -ssh %s@%s -pw %s %s" % (PLINK_PATH, RARITAN_USER, RARITAN_HOST, RARITAN_PASSWORD, pdu_cmd)
    putty_pid = subprocess.call(full_cmd, shell=True)

def set_display_power(port, state):
    if state:
        RunRaritanCmd("", port, "on")
    else:
        RunRaritanCmd("", port, "off")

def get_set_cycle(question):
    response = None
    str_lst = map(str,range(1,9999))
    while response not in (str_lst):
        response = raw_input(question)
    return response

def end_cycle_act():
    global mon
    global cycles
    global errors
    global running
    global num_cycles
    cycles = cycles - 1
    print '\n\n\nCycles: %i' % cycles
    print 'Errors: %i\n' % errors
    print "The number of test cycles has been reached."
    num_cycles = int(9999)
    del_reg_runonce()
    running = False
    cycles = 0
    errors = 0
    mon = []
    return

def getdisp_rez():
    """ Enumerate all connected displays and detect their resolution """
    i = 0
    while True:
        try:
            winDev = win32api.EnumDisplayDevices(DevNum=i)
            winSettings = win32api.EnumDisplaySettings(winDev.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
            i = i + 1;
            winMon = win32api.EnumDisplayDevices(winDev.DeviceName, DevNum=0)
            name_res = '%s %sx%s' % (winMon.DeviceString, winSettings.PelsWidth, winSettings.PelsHeight)
            global mon
            mon.append(name_res)
        except:
            break;
    
def dpms():
    global mon
    global cycles
    global errors
    global running
    global num_cycles
    stp_error = ask_yes_no("Do you want to stop when an error is found? (y/n): ")
    if stp_error == "y":
        print "The test will stop on error.\n"
    else:
        print "The test will continue on error.\n"
    set_cycle = ask_yes_no("Do you want to set the number of test cycles? (y/n): ")
    if set_cycle == "y":
        num_cycles = get_set_cycle("\nEnter the number of test cycles to run (1-9999): ")
        num_cycles = int(num_cycles)
        print "The test will run for %i cycles.\n" % (num_cycles)
    else:
        print "\n"
        pass
    while cycles == 0:
        # get attached displays resolution
        getdisp_rez()       
        # cycle counter ++
        cycles = cycles + 1
        # print the test is starting
        print "Starting the DPMS test...\n\n"
        # notify the number of attached displays
        disp_cnt = GetSystemMetrics(80)
        POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
        print POPUP_STR, "\n"
        axWshShell = win32com.client.Dispatch("Wscript.Shell")
        axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
        # print name and resolution of attached displays
        print "**************"
        print "** Displays **"
        print "**************"
        for x in mon:
            print x, "\n"
        #notify the test is starting
        POPUP_STR = "Starting the DPMS test"
        axWshShell.Popup(POPUP_STR,2,u"Starting Test",64)
        time.sleep(4)
    else:
        while running == True and cycles < num_cycles + 1:
            # dpms command
            print '\nPutting the displays into DPMS...\n'
            win32api.PostMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con. SC_MONITORPOWER, 2)
            # timeout
            time.sleep(15)
            # take out of DPMS by moving the mouse
            click(10,10)
            # get display count
            disp_cnt = GetSystemMetrics(80)
            # check for display count errors. if stop on error flag is set stop, if not continue looping
            while disp_cnt != len(mon):
                # print number of attached displays
                POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                print POPUP_STR, "\n"
                axWshShell = win32com.client.Dispatch("Wscript.Shell")
                axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
                # get attached displays resolution
                mon = []
                getdisp_rez()
                # print name and resolution of attached displays
                print "**************"
                print "** Displays **"
                print "**************"
                for x in mon:
                    print x, "\n"
                # print found an error
                print "\nFound an error!\n\n"
                if stp_error == "y":
                    win32api.MessageBox(0, 'The number of attached displays \nhas changed!', 'Error - YDIW', 0x00001010)
                    running = False
                    errors = errors + 1
                    # print cycle and error count
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i' % errors
                    cycles = 0
                    errors = 0
                    mon = []
                    return
                else:
                    # print number of attached displays
                    POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                    print POPUP_STR, "\n"
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
                    # get attached displays resolution
                    mon = []
                    getdisp_rez()
                    # print name and resolution of attached displays
                    print "**************"
                    print "** Displays **"
                    print "**************"
                    for x in mon:
                        print x, "\n"
                    # print found an error
                    print "\nFound an error!\n\n"
                    POPUP_STR = "The number of attached displays has changed!"
                    errors = errors + 1
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Error - YDIW",16)
                    # print cycle and error count
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i' % errors
                    # this will break out of the nested loop
                    disp_cnt = len(mon)
            # no errors found. 
            else:
                # get display count again since it may have been changed to break out of the 'don't stop on error' loop
                disp_cnt = GetSystemMetrics(80)
                time.sleep(6)
                POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                print "\n",POPUP_STR,"\n"
                axWshShell = win32com.client.Dispatch("Wscript.Shell")
                axWshShell.Popup(POPUP_STR,2,u"Display Count Info",64)
            # check/compare resolution
                cur_mon = mon
                mon = []
                # get attached displays resolution
                getdisp_rez()
                # print name and resolution of attached displays
                print "**************"
                print "** Displays **"
                print "**************"
                for x in mon:
                    print x, "\n"
                # check for resolution errors. if stop on error flag is set stop, if not continue looping
                if cur_mon != mon:
                    print "\nFound an error!\n\n"
                    if stp_error == "y":
                        win32api.MessageBox(0, 'One of the attached displays \nresolution has changed!', 'Error - YDIW', 0x00001010)
                        running = False
                        errors = errors + 1
                        # print cycle and error count
                        print 'Cycles: %i' % cycles
                        print 'Errors: %i' % errors
                        cycles = 0
                        errors = 0
                        mon = []
                        return
                    else:
                        POPUP_STR = "One of the attached displays resolution has changed!"
                        errors = errors + 1
                        axWshShell = win32com.client.Dispatch("Wscript.Shell")
                        axWshShell.Popup(POPUP_STR,2,u"Error - YDIW",16)
                        # print cycle and error count
                        print 'Cycles: %i' % cycles
                        print 'Errors: %i\n' % errors
                        # ++cycles
                        time.sleep(4)
                        cycles = cycles + 1
                        continue
                # no errors found. print cycle and error count, ++cycles
                else:
                    time.sleep(4)
                    POPUP_STR = "The attached displays resolution has not changed."
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Resolution Info",64)
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i\n' % errors
                    time.sleep(4)
                    cycles = cycles + 1
    # actions to run after the set_cycle number has been hit
    end_cycle_act()
                
def sleep():
    global mon
    global cycles
    global errors
    global running
    global num_cycles
    stp_error = ask_yes_no("Do you want to stop when an error is found? (y/n): ")
    if stp_error == "y":
        print "The test will stop on error.\n"
    else:
        print "The test will continue on error.\n"
    set_cycle = ask_yes_no("Do you want to set the number of test cycles? (y/n): ")
    if set_cycle == "y":
        num_cycles = get_set_cycle("\nEnter the number of test cycles to run (1-9999): ")
        num_cycles = int(num_cycles)
        print "The test will run for %i cycles.\n" % (num_cycles)
    else:
        print "\n"
        pass
    while cycles == 0:
        # get attached displays resolution
        getdisp_rez()       
        # cycle counter ++
        cycles = cycles + 1
        # print the test is starting
        print "Starting the System Sleep test...\n\n"
        # notify the number of attached displays
        disp_cnt = GetSystemMetrics(80)
        POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
        print POPUP_STR, "\n"
        axWshShell = win32com.client.Dispatch("Wscript.Shell")
        axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
        # print name and resolution of attached displays
        print "**************"
        print "** Displays **"
        print "**************"
        for x in mon:
            print x, "\n"
        #notify the test is starting
        POPUP_STR = "Starting the System Sleep test"
        axWshShell.Popup(POPUP_STR,2,u"Starting Test",64)
    else:
        while running == True and cycles < num_cycles + 1:
            # sleep command
            print '\nPutting the system to sleep...\n'
            os.system("powercfg -h off")
            os.system(r'%windir%\System32\rundll32.exe powrprof.dll, SetSuspendState 0,1,0')
            # timeout
            time.sleep(5)
            # get display count
            disp_cnt = GetSystemMetrics(80)
            # check for display count errors. if stop on error flag is set stop, if not continue looping
            while disp_cnt != len(mon):
                # print number of attached displays
                POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                print "\n", POPUP_STR, "\n"
                axWshShell = win32com.client.Dispatch("Wscript.Shell")
                axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
                # get attached displays resolution
                mon = []
                getdisp_rez()
                # print name and resolution of attached displays
                print "**************"
                print "** Displays **"
                print "**************"
                for x in mon:
                    print x, "\n"
                # print found an error
                print "\nFound an error!\n\n"
                if stp_error == "y":
                    win32api.MessageBox(0, 'The number of attached displays \nhas changed!', 'Error - YDIW', 0x00001010)
                    running = False
                    errors = errors + 1
                    # print cycle and error count
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i' % errors
                    cycles = 0
                    errors = 0
                    mon = []
                    return
                else:
                    # print number of attached displays
                    POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                    print POPUP_STR, "\n"
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
                    # get attached displays resolution
                    mon = []
                    getdisp_rez()
                    # print name and resolution of attached displays
                    print "**************"
                    print "** Displays **"
                    print "**************"
                    for x in mon:
                        print x, "\n"
                    # print found an error
                    print "\nFound an error!\n\n"
                    POPUP_STR = "The number of attached displays has changed!"
                    errors = errors + 1
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Error - YDIW",16)
                    # print cycle and error count
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i' % errors
                    # this will break out of the nested loop
                    disp_cnt = len(mon)
            # no errors found. 
            else:
                # get display count again since it may have been changed to break out of the 'don't stop on error' loop
                disp_cnt = GetSystemMetrics(80)
                time.sleep(6)
                POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                print "\n",POPUP_STR,"\n"
                axWshShell = win32com.client.Dispatch("Wscript.Shell")
                axWshShell.Popup(POPUP_STR,2,u"Display Count Info",64)
            # check/compare resolution
                cur_mon = mon
                mon = []
                # get attached displays resolution
                getdisp_rez()
                # print name and resolution of attached displays
                print "**************"
                print "** Displays **"
                print "**************"
                for x in mon:
                    print x, "\n"
                # check for resolution errors. if stop on error flag is set stop, if not continue looping
                if cur_mon != mon:
                    print "\nFound an error!\n\n"
                    if stp_error == "y":
                        win32api.MessageBox(0, 'One of the attached displays \nresolution has changed!', 'Error - YDIW', 0x00001010)
                        running = False
                        errors = errors + 1
                        # print cycle and error count
                        print 'Cycles: %i' % cycles
                        print 'Errors: %i' % errors
                        cycles = 0
                        errors = 0
                        mon = []
                        return
                    else:
                        POPUP_STR = "One of the attached displays resolution has changed!"
                        errors = errors + 1
                        axWshShell = win32com.client.Dispatch("Wscript.Shell")
                        axWshShell.Popup(POPUP_STR,2,u"Error - YDIW",16)
                        # print cycle and error count
                        print 'Cycles: %i' % cycles
                        print 'Errors: %i\n' % errors
                        # ++cycles
                        time.sleep(4)
                        cycles = cycles + 1
                        continue
                # no errors found. print cycle and error count, ++cycles
                else:
                    time.sleep(4)
                    POPUP_STR = "The attached displays resolution has not changed."
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Resolution Info",64)
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i\n' % errors
                    time.sleep(4)
                    cycles = cycles + 1
    # actions to run after the set_cycle number has been hit
    end_cycle_act()

def reboot():
    global mon
    global key
    global cycles
    global errors
    global running
    global num_cycles
    # clear cycle count in case they persist from previous tests
    cycles = 0
    errors = 0
    # clear temp file
    clr_tmp()
    # create temp file to store var list
    dirpath = os.path.dirname(sys.argv[0])
    pickle_path = dirpath + "/dst_tmp.dat"
    pickle_file = open(pickle_path, "wb+")
    # create list to store vars
    tmp_lst = []
    stp_error = ask_yes_no("Do you want to stop when an error is found? (y/n): ")
    if stp_error == "y":
        print "The test will stop on error.\n"
    else:
        print "The test will continue on error.\n"
    # set custom number of cycles
    set_cycle = ask_yes_no("Do you want to set the number of test cycles? (y/n): ")
    if set_cycle == "y":
        num_cycles = get_set_cycle("\nEnter the number of test cycles to run (1-9999): ")
        num_cycles = int(num_cycles)
        print "The test will run for %i cycles.\n" % (num_cycles)
    else:
        print "\n"
        pass
    while cycles == 0:
        # write to runonce registry key 
        reg_runonce()
        # get attached displays resolution
        getdisp_rez()
        # write stp_error to temp list
        tmp_lst.append(stp_error)
        # write mon [] to temp list
        tmp_lst.append(mon)
        # update cycle count
        cycles = cycles + 1
        # write cycles to temp list
        tmp_lst.append(cycles)
        # write errors to temp list
        tmp_lst.append(errors)
        # write num_cycles to temp list
        tmp_lst.append(num_cycles)
        # print the test is starting
        print "\nStarting the Reboot test...\n\n"
        # notify the number of attached displays
        disp_cnt = GetSystemMetrics(80)
        POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
        print POPUP_STR, "\n"
        axWshShell = win32com.client.Dispatch("Wscript.Shell")
        axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
        # print name and resolution of attached displays
        print "**************"
        print "** Displays **"
        print "**************"
        for x in mon:
            print x, "\n"
        #notify the test is starting
        POPUP_STR = "Starting the Reboot test"
        axWshShell.Popup(POPUP_STR,2,u"Starting Test",64)
        time.sleep(4)
        # write temp list to temp file
        cPickle.dump(tmp_lst, pickle_file)
        pickle_file.close()
        time.sleep(3)
        # reboot command and cycle counter ++
        print '\nRebooting the system...\n'
        os.system("shutdown -r -t 5")
        time.sleep(10)

def reboot_res():
    """ Actions to run after resuming from reboot. """
    global num_cycles
    global running
    global cycles
    global errors
    global mon
    # clear vars
    mon = []
    tmp_lst = []
    stp_error = ""
    cycles = 0
    errors = 0
    # open temp file and assign contents to temp list
    try:
        dirpath = os.path.dirname(sys.argv[0])
        pickle_path = dirpath + "/dst_tmp.dat"
        pickle_file = open(pickle_path, "rb")
        tmp_lst = cPickle.load(pickle_file)
        pickle_file.close()
    except EOFError:
        print "\nError: The temp file is empty!"
        running = False
        errors = 0
        cycles = 0
        mon = []
        main()
    # parse temp list and assign items to variables, stp_error, mon[], cycles, errors
    stp_error = tmp_lst[0]
    cycles = tmp_lst[2]
    errors = tmp_lst[3]
    num_cycles = tmp_lst[4]
    while running == True:
        disp_cnt = GetSystemMetrics(80)
        # print number of attached displays
        POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
        print POPUP_STR, "\n"
        axWshShell = win32com.client.Dispatch("Wscript.Shell")
        axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
        # get attached displays resolution
        getdisp_rez()
        # print name and resolution of attached displays
        print "**************"
        print "** Displays **"
        print "**************"
        for x in mon:
            print x, "\n"
        # timeout
        time.sleep(4)
        # get display count
        disp_cnt = GetSystemMetrics(80)
        # check for display count errors. if stop on error flag is set stop, if not continue looping
        while disp_cnt != len(mon):
            # print found an error
            print "\nFound an error!\n\n"
            if stp_error == "y":
                win32api.MessageBox(0, 'The number of attached displays \nhas changed!', 'Error - YDIW', 0x00001010)
                running = False
                errors = errors + 1
                # print cycle and error count
                print 'Cycles: %i' % cycles
                print 'Errors: %i' % errors
                running = False
                cycles = 0
                errors = 0
                mon = []
                main()
            else:
                POPUP_STR = "The number of attached displays has changed!"
                errors = errors + 1
                axWshShell = win32com.client.Dispatch("Wscript.Shell")
                axWshShell.Popup(POPUP_STR,2,u"Error - YDIW",16)
                # print cycle and error count
                print 'Cycles: %i' % cycles
                print 'Errors: %i' % errors
                # this will break out of the nested loop
                disp_cnt = len(mon)
        # no errors found. 
        else:
            # get display count again since it may have been changed to break out of the 'don't stop on error' loop
            disp_cnt = GetSystemMetrics(80)
            time.sleep(6)
        # make a copy of mon list, clear mon list and temp list
            mon = tmp_lst[1]
            cur_mon = mon
            mon = []
            tmp_lst = []
            # get attached displays resolution
            getdisp_rez()
            # check for resolution errors(compare copy of mon list to mon list). if stop on error flag is set stop, if not continue looping
            if cur_mon != mon:
                print "\nFound an error!\n\n"
                if stp_error == "y":
                    win32api.MessageBox(0, 'One of the attached displays \nresolution has changed!', 'Error - YDIW', 0x00001010)
                    running = False
                    errors = errors + 1
                    # print cycle and error count
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i' % errors
                    running = False
                    cycles = 0
                    errors = 0
                    mon = []
                    main()
                else:
                    POPUP_STR = "One of the attached displays resolution has changed!"
                    errors = errors + 1
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Error - YDIW",16)
                    # print cycle and error count
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i' % errors
                    # ++cycles
                    time.sleep(4)
                    cycles = cycles + 1
                    # write vars to temp list
                    tmp_lst.append(stp_error)
                    tmp_lst.append(mon)
                    tmp_lst.append(cycles)
                    tmp_lst.append(errors)
                    tmp_lst.append(num_cycles)
                    # write temp list to temp file
                    pickle_file = open(pickle_path, "wb+")
                    cPickle.dump(tmp_lst, pickle_file)
                    pickle_file.close()
                    time.sleep(3)
                    # send the reboot command
                    if cycles < num_cycles + 1:
                        print '\n\nRebooting the system...\n'
                        os.system("shutdown -r -t 5")
                        time.sleep(10)
                    else:
                        running = False
            # no errors found. print cycle and error count, ++cycles
            else:
                time.sleep(4)
                POPUP_STR = "The attached displays resolution has not changed."
                axWshShell = win32com.client.Dispatch("Wscript.Shell")
                axWshShell.Popup(POPUP_STR,2,u"Resolution Info",64)
                print 'Cycles: %i' % cycles
                print 'Errors: %i\n' % errors
                time.sleep(4)
            # update the cycle count
            cycles = cycles + 1
            # write vars to temp list
            tmp_lst.append(stp_error)
            tmp_lst.append(mon)
            tmp_lst.append(cycles)
            tmp_lst.append(errors)
            tmp_lst.append(num_cycles)
            # write temp list to temp file
            pickle_file = open(pickle_path, "wb+")
            cPickle.dump(tmp_lst, pickle_file)
            pickle_file.close()
            time.sleep(3)
            # send the reboot command
            if cycles < num_cycles + 1:
                print '\nRebooting the system...\n'
                os.system("shutdown -r -t 5")
                time.sleep(10)
            else:
                break
    # actions to run after the set_cycle number has been hit
    cycles = cycles - 1
    print '\n\n\nCycles: %i' % cycles
    print 'Errors: %i\n' % errors
    print "The number of test cycles has been reached."
    num_cycles = int(9999)
    del_reg_runonce()
    running = False
    cmdargs = ""
    cycles = 0
    errors = 0
    mon = []
    main()
    
def pwrcycle():
    global mon
    global cycles
    global errors
    global running
    global num_cycles
    stp_error = ask_yes_no("Do you want to stop when an error is found? (y/n): ")
    if stp_error == "y":
        print "The test will stop on error.\n"
    else:
        print "The test will continue on error.\n"
    set_cycle = ask_yes_no("Do you want to set the number of test cycles? (y/n): ")
    if set_cycle == "y":
        num_cycles = get_set_cycle("\nEnter the number of test cycles to run (1-9999): ")
        num_cycles = int(num_cycles)
        print "The test will run for %i cycles.\n" % (num_cycles)
    else:
        print "\n"
        pass
    pdu_port = get_port_number("Which PDU port number is the display plugged in to? (1-8): ")
    # turn on the port in case it is off
    print "\n"
    set_display_power(pdu_port, 1)
    time.sleep(8)
    while cycles == 0:
        # get attached displays resolution
        getdisp_rez()       
        # cycle counter ++
        cycles = cycles + 1
        # print the test is starting
        print "\nStarting the Display Power Cycle test...\n\n"
        # notify the number of attached displays
        disp_cnt = GetSystemMetrics(80)
        POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
        print POPUP_STR, "\n"
        axWshShell = win32com.client.Dispatch("Wscript.Shell")
        axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
        # print name and resolution of attached displays
        print "**************"
        print "** Displays **"
        print "**************"
        for x in mon:
            print x, "\n"
        #notify the test is starting
        POPUP_STR = "Starting the Display Power Cycle test"
        axWshShell.Popup(POPUP_STR,2,u"Starting Test",64)
        time.sleep(4)
    else:
        while running == True and cycles < num_cycles + 1:
            # turn off the pdu port
            print "\nTurning off the PDU port...\n\n"
            set_display_power(pdu_port, 0)
            # timeout
            time.sleep(20)
            # turn on the pdu port
            print "\nTurning on the PDU port...\n\n"
            set_display_power(pdu_port, 1)
            # timeout to allow the display to power up
            time.sleep(15)
            # get display count
            disp_cnt = GetSystemMetrics(80)
            # check for display count errors. if stop on error flag is set stop, if not continue looping
            while disp_cnt != len(mon):
                # print number of attached displays
                POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                print POPUP_STR, "\n"
                axWshShell = win32com.client.Dispatch("Wscript.Shell")
                axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
                # get attached displays resolution
                mon = []
                getdisp_rez()
                # print name and resolution of attached displays
                print "**************"
                print "** Displays **"
                print "**************"
                for x in mon:
                    print x, "\n"
                # print found an error
                print "\nFound an error!\n\n"
                if stp_error == "y":
                    win32api.MessageBox(0, 'The number of attached displays \nhas changed!', 'Error - YDIW', 0x00001010)
                    running = False
                    errors = errors + 1
                    # print cycle and error count
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i' % errors
                    cycles = 0
                    errors = 0
                    mon = []
                    return
                else:
                    # print number of attached displays
                    POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                    print POPUP_STR, "\n"
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Display Count",64)
                    # get attached displays resolution
                    mon = []
                    getdisp_rez()
                    # print name and resolution of attached displays
                    print "**************"
                    print "** Displays **"
                    print "**************"
                    for x in mon:
                        print x, "\n"
                    # print found an error
                    print "\nFound an error!\n\n"
                    POPUP_STR = "The number of attached displays has changed!"
                    errors = errors + 1
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Error - YDIW",16)
                    # print cycle and error count
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i\n' % errors
                    # this will break out of the nested loop
                    disp_cnt = len(mon)
            # no errors found. 
            else:
                # get display count again since it may have been changed to break out of the 'don't stop on error' loop
                disp_cnt = GetSystemMetrics(80)
                time.sleep(2)
                POPUP_STR = "There is %i display(s) attached." % (disp_cnt)
                print "\n" + POPUP_STR + "\n"
                axWshShell = win32com.client.Dispatch("Wscript.Shell")
                axWshShell.Popup(POPUP_STR,2,u"Display Count Info",64)
            # make a copy of mon list, clear mon list
                cur_mon = mon
                mon = []
                # get attached displays resolution
                getdisp_rez()
                # print name and resolution of attached displays
                print "**************"
                print "** Displays **"
                print "**************"
                for x in mon:
                    print x, "\n"
                # check for resolution errors(compare copy of mon list to mon list). if stop on error flag is set stop, if not continue looping
                if cur_mon != mon:
                    print "\nFound an error!\n\n"
                    if stp_error == "y":
                        win32api.MessageBox(0, 'One of the attached displays \nresolution has changed!', 'Error - YDIW', 0x00001010)
                        running = False
                        errors = errors + 1
                        # print cycle and error count
                        print 'Cycles: %i' % cycles
                        print 'Errors: %i' % errors
                        cycles = 0
                        errors = 0
                        mon = []
                        return
                    else:
                        POPUP_STR = "One of the attached displays resolution has changed!"
                        errors = errors + 1
                        axWshShell = win32com.client.Dispatch("Wscript.Shell")
                        axWshShell.Popup(POPUP_STR,2,u"Error - YDIW",16)
                        # print cycle and error count
                        print 'Cycles: %i' % cycles
                        print 'Errors: %i\n' % errors
                        # ++cycles
                        cycles = cycles + 1
                        time.sleep(4)
                        continue
                # no errors found. print cycle and error count, ++cycles
                else:
                    time.sleep(4)
                    POPUP_STR = "The attached displays resolution has not changed."
                    axWshShell = win32com.client.Dispatch("Wscript.Shell")
                    axWshShell.Popup(POPUP_STR,2,u"Resolution Info",64)
                    print 'Cycles: %i' % cycles
                    print 'Errors: %i\n' % errors
                    time.sleep(4)
                    cycles = cycles + 1
    # actions to run after the set_cycle number has been hit
    end_cycle_act()
                    
# main
def main():
    global running
    global errors
    global cycles
    global mon
    # display the choice menu
    try:
        choice = None
        while choice != "0":
            print \
            """
            *******************************
            ***           DST           ***                   
            ***                         ***
            *** (DISPLAY STRESS TESTER) ***
            ***                         ***

            0 - Quit
            1 - Start DPMS test
            2 - Start Reboot test
            3 - Start System Sleep test
            4 - Start Disp. Power Cycle test

            -------------------------------
                                     
            *+*    press control-c to   *-*
            *-*      stop the test      *+*  
            """
            # check for passed argument for reboot test
            cmdargs = str(sys.argv[1:])
            if "reboot" in cmdargs and running == True:
                print "\nResuming the Reboot test...\n\n"
                # code for resuming reboot test
                running = True
                cycles = cycles + 1
                try:
                    reboot_res()
                except Exception, err:
                    print Exception, err
                raw_input("\nPress enter to exit.")
            else:
                pass
            choice = raw_input("Choice: ")
            print
            # exit
            if choice == "0":
                print "Good-bye."
            # start DPMS test
            elif choice == "1":
                running = True
                dpms()
            # start system Sleep test
            elif choice == "2":
                running = True
                reboot()
            # start reboot test
            elif choice == "3":
                running = True
                sleep()
            # start power cycle test
            elif choice == "4":
                running = True
                pwrcycle()
            # some unknown choice
            else:
                print "\nSorry, but", choice, "isn't a valid choice."
    # catch control-c
    except KeyboardInterrupt:

        # delete Run registry key if it exists
        del_reg_runonce()
        running = False
        print '\n\nCycles: %i' % cycles
        print 'Errors: %i\n' % errors
        print 'The test was cancelled by a user.'
        num_cycles = int(9999)
        errors = 0
        cycles = 0
        mon = []
        main()

main()
raw_input("\n\nPress the enter key to exit.")

