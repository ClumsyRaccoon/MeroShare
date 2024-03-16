import logging
from tenacity import retry, stop_after_attempt, wait_fixed
import requests
import json
import os
import pandas as pd
from openpyxl import load_workbook
import time

USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.105 Safari/537.36"
logging.basicConfig(format="%(asctime)s %(message)s", level=logging.INFO)
cafile = 'files/cdsc-com-np-chain.pem'


class MeroShare:
    def __init__(
            self,
            name: str = None,
            dpid: int = None,
            username: int = None,
            password: str = None,
            client_id: str = None,
            crn: str = None,
            pin: int = None,
            bank: str = None
    ):

        self.__name = name
        self.__dpid = str(dpid)
        self.__username = str(username)
        self.__password = password
        self.__session = requests.Session()
        self.__auth_token = None
        self.status = None
        self.__capital_id = self.get_capital_id()
        self.__dmat = "130" + self.__dpid + self.__username
        self.client_id = client_id
        self.__crn = crn
        self.__pin = pin
        self.__applicable_issues = None
        self.__account = None
        self.bank = bank

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(2), reraise=True)
    def get_capital_id(self):
        if os.path.exists('files/capitals.json'):
            try:
                return \
                    [response['id'] for response in json.load(open('files/capitals.json')) if
                     response['code'] == self.__dpid][0]
            except Exception:
                print('Could not find capital id cache \n Updating cache and Retrying...')
        with self.__session as sess:
            headers = {
                "Accept": "application/json, text/plain, */*",
                "Accept-Language": "en-US,en;q=0.9",
                "Authorization": "null",
                "Connection": "keep-alive",
                "Origin": "https://meroshare.cdsc.com.np",
                "Referer": "https://meroshare.cdsc.com.np/",
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-site",
                "Sec-GPC": "1",
                "User-Agent": USER_AGENT,
            }
            sess.headers.clear()
            sess.headers.update(headers)

            try:
                sess.options("https://webbackend.cdsc.com.np/api/meroShare/auth/", verify=cafile)
                cap_list = sess.get("https://webbackend.cdsc.com.np/api/meroShare/capital/", verify=cafile).json()
                with open("files/capitals.json", "w") as cap_file:
                    json.dump(cap_list, cap_file)

                return [response['id'] for response in cap_list if response['code'] == self.__dpid][0]
            except Exception as error:
                self.status = 'Error finding capital id'
                logging.warning(f'{self.status} for Acc: {self.__name}')
                logging.info(error)
                return None

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(3), reraise=True)
    def login(self) -> bool:
        assert (
                self.__username and self.__password and self.__dpid
        ), "Username, password and DPID required!"

        if not self.__capital_id:
            logging.warning("Capital ID Problem")
            return False

        try:
            with self.__session as sess:
                data = json.dumps(
                    {
                        "clientId": self.__capital_id,
                        "username": self.__username,
                        "password": self.__password,
                    }
                )

                headers = {
                    "Accept": "application/json, text/plain, */*",
                    "Accept-Language": "en-US,en;q=0.9",
                    "Authorization": "null",
                    "Connection": "keep-alive",
                    "Content-Type": "application/json",
                    "Origin": "https://meroshare.cdsc.com.np",
                    "Referer": "https://meroshare.cdsc.com.np/",
                    "Sec-Fetch-Dest": "empty",
                    "Sec-Fetch-Mode": "cors",
                    "Sec-Fetch-Site": "same-site",
                    "Sec-GPC": "1",
                    "User-Agent": USER_AGENT,
                }
                sess.headers.clear()
                sess.headers.update(headers)
                try:
                    login_req = sess.post(
                        "https://webbackend.cdsc.com.np/api/meroShare/auth/", data=data, verify=cafile)

                    if login_req.status_code == 200:
                        self.__auth_token = login_req.headers.get("Authorization")
                        logging.info(f"Logged in successfully!  Account: {self.__name}!")
                    else:
                        self.status = f'Login Failed! {login_req.status_code}'
                        logging.warning(f"{self.status} Account: {self.__name}")
                        print(f'Login Failed {self.__name}')
                except:
                    print('Login Connection Refused')
                    self.status = 'Login Connection Refused'

        except Exception as error:
            print(error)
            logging.info(error)
            logging.error(
                f"Login request failed! Retrying ({self.login.retry.statistics.get('attempt_number')})!"
            )
            self.status = 'Login Failed'
            return False

        return True

    @retry(stop=stop_after_attempt(5), wait=wait_fixed(3), reraise=True)
    def get_applicable_issues(self):
        try:
            with self.__session as sess:
                data = json.dumps(
                    {
                        "filterFieldParams": [
                            {
                                "key": "companyIssue.companyISIN.script",
                                "alias": "Scrip",
                            },
                            {
                                "key": "companyIssue.companyISIN.company.name",
                                "alias": "Company Name",
                            },
                            {
                                "key": "companyIssue.assignedToClient.name",
                                "value": "",
                                "alias": "Issue Manager",
                            },
                        ],
                        "page": 1,
                        "size": 10,
                        "searchRoleViewConstants": "VIEW_APPLICABLE_SHARE",
                        "filterDateParams": [
                            {
                                "key": "minIssueOpenDate",
                                "condition": "",
                                "alias": "",
                                "value": "",
                            },
                            {
                                "key": "maxIssueCloseDate",
                                "condition": "",
                                "alias": "",
                                "value": "",
                            },
                        ],
                    }
                )

                headers = {
                    "Accept": "application/json, text/plain, */*",
                    "Accept-Language": "en-US,en;q=0.9",
                    "Authorization": self.__auth_token,
                    "Connection": "keep-alive",
                    "Content-Type": "application/json",
                    "Origin": "https://meroshare.cdsc.com.np",
                    "Referer": "https://meroshare.cdsc.com.np/",
                    "Sec-Fetch-Dest": "empty",
                    "Sec-Fetch-Mode": "cors",
                    "Sec-Fetch-Site": "same-site",
                    "Sec-GPC": "1",
                    "User-Agent": USER_AGENT,
                }
                sess.headers.clear()
                sess.headers.update(headers)

                tries = 15
                for i in range(tries):
                    try:
                        issue_req = sess.post(
                            "https://webbackend.cdsc.com.np/api/meroShare/companyShare/applicableIssue/",
                            data=data, verify=cafile)
                        assert issue_req.status_code == 200, "Applicable issues request failed!"
                    except Exception as error:
                        if i < tries - 1:
                            time.sleep(.1)
                            print('Retrying')
                            continue
                        else:
                            self.status = 'Applicable issues request failed! after Retry'
                            print(self.status)
                    break

                self.__applicable_issues = issue_req.json().get("object")
                logging.info(f"Appplicable Issues Obtained! Account: {self.__name}")
                return self.__applicable_issues
        except Exception as error:
            logging.info(error)
            self.status = 'Error Getting Applicable Issue'
            logging.info(f"{self.status}... Retrying!")
            return 0

    def apply(self, share_id: str, qty: int):
        with self.__session as sess:
            try:
                issue_to_apply = None

                if not self.__applicable_issues:
                    self.get_applicable_issues()

                for issue in self.__applicable_issues:
                    if str(issue.get("scrip")) == share_id:
                        issue_to_apply = issue
                        break

                if not issue_to_apply:
                    logging.warning("Provided Script doesn't match any of the applicable issues!")
                    self.status = "No matching applicable issues!"
                    return 0

                share_id = issue_to_apply.get('companyShareId')

                if issue_to_apply.get("action"):
                    status = issue_to_apply.get("action")
                    self.status = "Couldn't apply for issue! - " + status
                    logging.warning(self.status)
                    return 0

                headers = {
                    "Accept": "application/json, text/plain, */*",
                    "Accept-Language": "en-US,en;q=0.9",
                    "Authorization": self.__auth_token,
                    "Connection": "keep-alive",
                    "Content-Type": "application/json",
                    "Origin": "https://meroshare.cdsc.com.np",
                    "Referer": "https://meroshare.cdsc.com.np/",
                    "Sec-Fetch-Dest": "empty",
                    "Sec-Fetch-Mode": "cors",
                    "Sec-Fetch-Site": "same-site",
                    "Sec-GPC": "1",
                    "User-Agent": USER_AGENT,
                }

                sess.headers.update(headers)
                bank_req = sess.get("https://webbackend.cdsc.com.np/api/meroShare/bank/", verify=cafile).json()

                bank_id = None
                for bank_ in bank_req:
                    if bank_['name'] == self.bank:
                        bank_id = bank_["id"]

                if bank_id is None:
                    self.status = "Bank name not found."
                    print(self.status)
                    return 0

                bank_specific_req = sess.get(
                    f"https://webbackend.cdsc.com.np/api/meroShare/bank/{bank_id}", verify=cafile
                )

                bank_specific_response_json = bank_specific_req.json()[0]

                branch_id = bank_specific_response_json.get("accountBranchId")
                account_number = bank_specific_response_json.get("accountNumber")
                customer_id = bank_specific_response_json.get("id")
                acc_type_id = bank_specific_response_json.get("accountTypeId")
        

                data = json.dumps(
                    {
                        "accountBranchId": branch_id,
                        "accountNumber": account_number,
                        "accountTypeId": acc_type_id,
                        "appliedKitta": qty,
                        "bankId": bank_id,
                        "boid": self.__dmat[-8:],
                        "companyShareId": share_id,
                        "crnNumber": self.__crn,
                        "customerId": customer_id,
                        "demat": self.__dmat,
                        "transactionPIN": self.__pin,
                    }
                )

            except Exception as error:
                logging.critical("Apply failed!")
                self.status = "Apply failed! (1)"
                logging.critical(error)
                return 0

            apply_req = sess.post(
                "https://webbackend.cdsc.com.np/api/meroShare/applicantForm/share/apply",
                data=data, verify=cafile
            )

            
            if apply_req.status_code != 201:
                self.status = f"Apply failed!"
                logging.warning(self.status)
                return 0

            logging.info(f"Sucessfully applied! Account: {self.__name}")
            self.status = f"Sucessfully applied! {qty} Kitta"
            return 0


def application_list(sheet, full_list, client_type, Scrip, qty):
    for details in sheet.iter_rows(min_row=2,
                                   min_col=1,
                                   values_only=True):
        if str(details[3]).upper() == 'NO':
            continue

        login_info = {"name": details[1],
                      "username": details[4].replace(" ", ""),
                      "password": details[6],
                      "dpid": int(details[5]) - 13000000,
                      "client_id": client_type + str(details[0]),
                      "crn": details[7],
                      "pin": details[8],
                      "bank": details[9]}

        ms = MeroShare(**login_info)
        ms.login()
        ms.apply(Scrip, qty)

        data = [ms.client_id] + [details[1]] + [str(details[5]) + str(details[4].replace(" ", ""))] + [Scrip] + [
            ms.status]
        full_list.loc[len(full_list)] = data

        time.sleep(.2)
        print('\n\n')

    return (full_list)


def read_excel(Scrip, qty):
    full_list = pd.DataFrame(columns=['Client ID', 'Name', 'Demat', 'Script', 'Application'])

    book = load_workbook(filename='MeroShare Login Details.xlsx', data_only=True)

    full_list = application_list(book['List'], full_list, '', Scrip, qty)
    
    print(full_list)
    full_list.to_excel(f'IPO Applied for {Scrip}.xlsx', index=False)


def start():
    print('NOTE: Code should match with MeroShare include Capitalizations.')
    script = input('Script Code to Apply For : ')
    qty = input('No. of Kitta to Apply : ')
    read_excel(script, qty)
    input('Press Enter to Continue....')


if __name__ == '__main__':
    start()
