import os
from files import Apply
from files import Check_Acc_Status
from files import IPO_Status
from files import MyShares_List
from files import GetApplicableIssue


def start():
    print('MENU: \n')
    print('1. Check Meroshare Account Status.')
    print('2. Get My Shares List.')
    print('3. Check Applicable Issue.')
    print('4. Apply for Issue.')
    print('5. Check Application Status.')
    print('0. EXIT \n\n')

    key = input('Please Choose a Menu Option : ')
    os.system('cls')


    if key not in ['0','1','2','3','4','5']:
        print('INVALID INPUT\n')
        start()

    if int(key) == 0:
        os._exit(0)
    elif int(key) == 1:
        Check_Acc_Status.start()
    elif int(key) == 2:
        MyShares_List.check_share()
    elif int(key) == 3:
        GetApplicableIssue.start()
    elif int(key) == 4:
        Apply.start()
    elif int(key) == 5:
        IPO_Status.start()

    os.system('cls')
    start()

if __name__ == '__main__':
    start()




