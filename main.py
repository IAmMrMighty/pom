import os
import ctypes
import time
import sys
import traceback
import getpass
import shutil
from colors import bcolors

#Check for administrator priviledges
is_admin = ctypes.windll.shell32.IsUserAnAdmin()
if is_admin == 0:
    print("This script requires administrator priviledges")
    sys.exit()

#Check if running Microsoft Windows Operating System
def checkos():
    if os.name == 'nt':
        main()
    else:
        print("This script can't be run on any other operating system than Windows")

#Clear Console
def clearconsole():
    os.system('cls')

## VARIABLES
rtset = 0 #Restorepoint set
username = getpass.getuser() #Get logged in user

def systemrestorepoint():
    global rtset
    try:
        lock = open("lock.txt") #Check for systemrestorepoint lock 
        lock.close
    except IOError: #Creating systemrestorepoint and creating lock. So it doesnt create one every start...
        os.system(r'powershell.exe -ExecutionPolicy Bypass -Command "Checkpoint-Computer -Description "POM." -RestorePointType "MODIFY_SETTINGS""')
        with open("lock.txt", 'w') as lock:
            lock.write('')
        rtset = 1
    except FileNotFoundError: #Same
        os.system(r'powershell.exe -ExecutionPolicy Bypass -Command "Checkpoint-Computer -Description "POM." -RestorePointType "MODIFY_SETTINGS""')
        with open("lock.txt", 'w') as lock:
            lock.write('')
        rtset = 1

def main():
    clearconsole()
    global rtsset
    try:
        systemrestorepoint()
        clearconsole()
        if rtset == 1:
            print(bcolors.OKGREEN + "System restore point set.")
            choose()
        elif rtset == 0:
            choose()

    #If user interrupts        
    except KeyboardInterrupt:
            print("\nScript canceled by user. Exiting.")
            sys.exit()
    #If other failure occurs
    except Exception:
            traceback.print_exc(file=sys.stdout)
    
#Optionmenu
choice = 0
def choose():
    global choice
    global username
    try:
        clearconsole()
        print("""
  __  __               _            _                 __  __   _           _       _           
 |  \/  |             | |          | |               |  \/  | (_)         | |     | |          
 | \  / |   __ _    __| |   ___    | |__    _   _    | \  / |  _    __ _  | |__   | |_   _   _ 
 | |\/| |  / _` |  / _` |  / _ \   | '_ \  | | | |   | |\/| | | |  / _` | | '_ \  | __| | | | |
 | |  | | | (_| | | (_| | |  __/   | |_) | | |_| |   | |  | | | | | (_| | | | | | | |_  | |_| |
 |_|  |_|  \__,_|  \__,_|  \___|   |_.__/   \__, |   |_|  |_| |_|  \__, | |_| |_|  \__|  \__, |
                                             __/ |                  __/ |                 __/ |
                                            |___/                  |___/                 |___/ 
""")
## MADE WITH http://www.patorjk.com/software/taag

        print("\n" + bcolors.WARNING + "POM - Performance Optimizer" + bcolors.FAIL + "\n \nThis script will only work on Windows Systems!" + bcolors.OKBLUE + "\n \n[1] Clear only System temporary files \n[2] Clear only Profile temporary files \n[3] Disable SysMain service\n[4] Clear only Prefetch\n[5] Clear Windows Update downloads \n[6] Check Roaming folder size \n[7] Check system drive C \n\n[9] Run everything \n[0] Exit")
        choice = input("> ")
        
        #Cleaning systemtemp
        if choice == "1":
            systemtemp = r"C:\Windows\Temp"
            clearconsole()
            try:
                print(bcolors.OKBLUE + "Clearing Systems temporary files")
                shutil.rmtree(systemtemp, ignore_errors=True)
                print(bcolors.OKGREEN + "\nDONE.")
                time.sleep(5)
                choose()
            except Exception as e:
                print(e)
                choose()

        #Cleaning profile temp
        elif choice == "2":
            profiletemp = "C:\\Users\\" + username + "\\AppData\\Local\\Temp\\"
            clearconsole()
            try:
                print(bcolors.OKBLUE + "Clearing profile temporary files")
                shutil.rmtree(profiletemp, ignore_errors=True)
                print(bcolors.OKGREEN + "\nDONE.")
                time.sleep(5)
                choose()
            except Exception as e:
                print(e)
                choose()

        #Disabling and stopping sysmain service
        elif choice == "3":
            clearconsole()
            print(bcolors.OKBLUE + "Disabling SysMain Service ...")
            os.system("sc config sysmain start= disabled >nul")
            time.sleep(1)
            os.system("net stop SysMain")
            print(bcolors.OKGREEN + "DONE.")
            time.sleep(1)
            choose()

        #Cleaning prefetch
        elif choice == "4":
            prefetch = "C:\\Windows\\Prefetch\\"
            clearconsole()
            try:
                print(bcolors.OKBLUE + "Clearing prefetch")
                if os.path.exists(prefetch):
                    shutil.rmtree(prefetch)
                    print(bcolors.OKGREEN + "\nDONE.")
                    os.mkdir(prefetch)
                    time.sleep(5)
                    choose()
                else:
                    print(bcolors.OKBLUE + "Prefetch folder doesn't exist. Creating ...")
                    os.mkdir(prefetch)
                    time.sleep(5)
                    choose()

            except Exception as e:
                print(e)
                time.sleep(4)
                choose()
        
        #Clear Windows Update Downloads
        elif choice == "5":
            softwaredistribution = 'C:\\Windows\\SoftwareDistribution\\Download\\'
            clearconsole()
            try:
                print(bcolors.OKBLUE + "Clearing Windows Update downloads")
                if os.path.exists(softwaredistribution):
                    shutil.rmtree(softwaredistribution)
                    os.mkdir(softwaredistribution)
                    print(bcolors.OKGREEN + "DONE.")
                    time.sleep(5)
                    choose()
                else:
                    print(bcolors.OKBLUE + "Windows Update Cache folder doesnt exist. Creating...")
                    os.mkdir(softwaredistribution)
                    time.sleep(5)
                    choose()
            except Exception as e:
                print(e)
                time.sleep(5)
                choose()
        
        #Checking Roaming folder size
        elif choice == "6":
            roamingpath = "C:\\Users\\" + username + "\\AppData\\Roaming\\"
            size = 0

            for path, dirs, files in os.walk(roamingpath):
                for f in files:
                    fp = os.path.join(path, f)
                    size += os.path.getsize(fp)
            roamingsize = size / 1024
            roamingsize = roamingsize / 1024
            roamingsize = round(roamingsize)
            if roamingsize >= 600:
                print(bcolors.FAIL + "Roaming folder might be full")
                print(str(roamingsize) + "MB\\600MB")
                time.sleep(5)
                choose()
            else:
                print(bcolors.OKGREEN + "Roaming folder is in good standing.")
                time.sleep(5)
                choose()
        
        elif choice == "7":
            clearconsole()
            print(bcolors.OKBLUE + "Checking system drive for failures ...")
            try:
                os.system("chkdsk c:>nul")
                os.system("sfc /scannow>nul")
                time.sleep(5)
                print(bcolors.OKGREEN + "DONE.")
                time.sleep(5)
                choose()
            except Exception as e:
                print(e)
    


    #RUN EVERYTHING!!
        elif choice == "9":
            clearconsole()
            print(bcolors.OKBLUE + "Running everything ...")
            
            ##VARIABLES
            systemtemp = r"C:\Windows\Temp"
            profiletemp = "C:\\Users\\" + username + "\\AppData\\Local\\Temp\\"
            prefetch = "C:\\Windows\\Prefetch\\"
            softwaredistribution = 'C:\\Windows\\SoftwareDistribution\\Download\\'
            roamingpath = "C:\\Users\\" + username + "\\AppData\\Roaming\\"
            size = 0
            
            #RUN
            clearconsole()
            print(bcolors.OKBLUE + "Clearing Systems temporary files")
            shutil.rmtree(systemtemp, ignore_errors=True)
            print(bcolors.OKGREEN + "\nDONE.")
            time.sleep(5)
            print(bcolors.OKBLUE + "Clearing profile temporary files")
            shutil.rmtree(profiletemp, ignore_errors=True)
            print(bcolors.OKGREEN + "\nDONE.")
            time.sleep(5)
            print(bcolors.OKBLUE + "Disabling SysMain Service ...")
            os.system("sc config sysmain start= disabled >nul")
            time.sleep(1)
            os.system("net stop SysMain")
            print(bcolors.OKGREEN + "DONE.")
            time.sleep(1)
            print(bcolors.OKBLUE + "Clearing prefetch")
            if os.path.exists(prefetch):
                shutil.rmtree(prefetch)
                print(bcolors.OKGREEN + "\nDONE.")
                os.mkdir(prefetch)
                time.sleep(5)
            else:
                print(bcolors.OKBLUE + "Prefetch folder doesn't exist. Creating ...")
                os.mkdir(prefetch)
                time.sleep(5)
            print(bcolors.OKBLUE + "Clearing Windows Update downloads")
            if os.path.exists(softwaredistribution):
                shutil.rmtree(softwaredistribution)
                os.mkdir(softwaredistribution)
                print(bcolors.OKGREEN + "DONE.")
            else:
                print(bcolors.OKBLUE + "Windows Update Cache folder doesnt exist. Creating...")
                os.mkdir(softwaredistribution)
                time.sleep(5)
            for path, dirs, files in os.walk(roamingpath):
                for f in files:
                    fp = os.path.join(path, f)
                    size += os.path.getsize(fp)
            roamingsize = size / 1024
            roamingsize = roamingsize / 1024
            roamingsize = round(roamingsize)
            if roamingsize >= 600:
                print(bcolors.FAIL + "Roaming folder might be full")
                print(str(roamingsize) + "MB\\600MB")
                time.sleep(5)
            else:
                print(bcolors.OKGREEN + "Roaming folder is in good standing.")
                time.sleep(5)
            print(bcolors.OKBLUE + "\nChecking system drive for failures [THIS MIGHT TAKE A WHILE] ...")
            os.system("chkdsk c:>nul")
            os.system("sfc /scannow>nul")
            time.sleep(5)
            print(bcolors.OKGREEN + "DONE.")
            time.sleep(5)
            choose()

        elif choice == "0":
            clearconsole()
            print("Good bye! Have a nice day and thanks for using my script.")
            sys.exit()        

    except KeyboardInterrupt:
        print(bcolors.FAIL + "\nScript canceled by user. Exiting.")
        sys.exit()
    except Exception:
        traceback.print_exc(file=sys.stdout)

#Startup
if __name__ == "__main__":
    clearconsole()
    width = "129"
    height = "34"
    os.system("mode 129,34")
    checkos()
