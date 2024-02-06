import os
import Apply
import Check_Acc_Staus
import IPO_Status
import MyShares_List
import msvcrt
from pynput import keyboard


def start():
    print('MENU: \n')
    print('1. Check Meroshare Account Status.')
    print('2. Get My Shares List.')
    print('3. Check Application Status.')
    print('4. Apply for Issue.')
    print('0. EXIT \n\n')

    key = input('Please Choose a Menu Option : ')
    os.system('cls')
    
    if key == '':
        start()
        
    if int(key) == 0:
        os._exit(0)
    elif int(key) == 1:
        Check_Acc_Staus.start()
    elif int(key) == 2:
        MyShares_List.check_share()
    elif int(key) == 3:
        IPO_Status.start()
    elif int(key) == 4:
        Apply.start()
    else:
        print ('Invalid Input....')
        start()

    start()

if __name__ == '__main__':
    start()




