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
            pin: int = None
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
            sess.headers.update(headers)

            try:
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
                sess.headers.update(headers)
                try:
                    sess.options("https://webbackend.cdsc.com.np/api/meroShare/auth/", verify=cafile)
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
                issue_req = sess.post(
                    "https://webbackend.cdsc.com.np/api/meroShare/companyShare/applicableIssue/",
                    data=data, verify=cafile
                )
                assert issue_req.status_code == 200, "Applicable issues request failed!"

                self.__applicable_issues = issue_req.json().get("object")
                logging.info(f"Appplicable Issues Obtained! Account: {self.__name}")
                return self.__applicable_issues
        except Exception as error:
            logging.info(error)
            self.status = 'Error Getting Applicable Issue'
            logging.info(f"{self.status}... Retrying!")
            return 0

    @retry(stop=stop_after_attempt(5), wait=wait_fixed(3), reraise=True)
    def get_application_status(self, scrip: str):
        with self.__session as sess:
            try:
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
                data = json.dumps(
                    {
                        "filterFieldParams": [
                            {
                                "key": "companyShare.companyIssue.companyISIN.script",
                                "alias": "Scrip",
                            },
                            {
                                "key": "companyShare.companyIssue.companyISIN.company.name",
                                "alias": "Company Name",
                            },
                        ],
                        "page": 1,
                        "size": 200,
                        "searchRoleViewConstants": "VIEW_APPLICANT_FORM_COMPLETE",
                        "filterDateParams": [
                            {
                                "key": "appliedDate",
                                "condition": "",
                                "alias": "",
                                "value": "",
                            },
                            {
                                "key": "appliedDate",
                                "condition": "",
                                "alias": "",
                                "value": "",
                            },
                        ],
                    }
                )

                tries = 20
                for i in range(tries):
                    try:
                        recent_applied_req1 = sess.post(
                                            "https://webbackend.cdsc.com.np/api/meroShare/applicantForm/active/search/",
                                            data=data, verify=cafile)
                        recent_applied_req = recent_applied_req1.json()['object']
                        assert recent_applied_req1.status_code == 200, "Application request failed!"
                    except:
                        if i < tries-1:
                            time.sleep(.1)
                            print('Retrying application request')
                            continue
                        else:
                            print("Application request failed after Retry")
                            self.status = 'Application rqeuest Failed'
                    break
                

                target_issue = None

                if self.status == 'Application rqeuest Failed':
                    return 
                for issue in recent_applied_req:
                    if issue['scrip'] == scrip:
                        target_issue = issue


                if not target_issue:
                    self.status = "Script not found!"
                    logging.critical(f"{self.status}")

                headers = {
                    "Accept": "application/json, text/plain, */*",
                    "Accept-Language": "en-US,en;q=0.9",
                    "Authorization": self.__auth_token,
                    "Connection": "keep-alive",
                    "Origin": "https://meroshare.cdsc.com.np",
                    "Referer": "https://meroshare.cdsc.com.np/",
                    "Sec-Fetch-Dest": "empty",
                    "Sec-Fetch-Mode": "cors",
                    "Sec-Fetch-Site": "same-site",
                    "Sec-GPC": "1",
                    "User-Agent": USER_AGENT,
                }
                sess.headers.update(headers)

                tries = 20
                for i in range(tries):
                    try:
                        details_req1 = sess.get(
                            f"https://webbackend.cdsc.com.np/api/meroShare/applicantForm/report/detail/{target_issue['applicantFormId']}", verify=cafile
                            )
                        details_req = details_req1.json()
                        self.status = details_req['statusName']
                        assert recent_applied_req1.status_code == 200, "Report request failed!"
                    except:
                        if i < tries-1:
                            time.sleep(.1)
                            print('Retrying Reoprt Request')
                            continue
                        else:
                            print("Report request failed after Retry")
                            self.status = 'Report rqeuest Failed'
                    break
                logging.info(f"Status: {self.status} for {self.__name}")
                return (details_req)

            except Exception as error:
                self.status = "Application status request failed."
                logging.warning(f"Application status request failed! Retrying ({self.get_application_status.retry.statistics.get('attempt_number')})")
                logging.exception(error)
                return 0


def application_list(sheet, full_list, client_type, Scrip):
    for details in sheet.iter_rows(min_row=2,
                                   min_col=1,
                                   values_only=True):
        if str(details[3]).upper()=='NO':
            continue

        login_info = {"name": details[1],
                      "username": details[4].replace(" ", ""),
                      "password": details[6],
                      "dpid": int(details[5]) - 13000000,
                      "client_id": client_type + str(details[0]),
                      "crn": details[7],
                      "pin": details[8]}

        ms = MeroShare(**login_info)
        ms.login()
        ms.get_application_status(Scrip)

        data = [ms.client_id] + [details[1]] + [str(details[5]) + str(details[4].replace(" ",""))] + [Scrip] + [ms.status]
        full_list.loc[len(full_list)] = data

        print('\n')

    return (full_list)


def  read_excel(Script):
    full_list = pd.DataFrame(columns=['Client ID', 'Name', 'Demat', 'Scrip', 'Status'])

    book = load_workbook(filename='MeroShare Login Details.xlsx', data_only=True)

    full_list = application_list(book['List'], full_list, '', Script)

    print(full_list)
    full_list.to_excel(f'Application Status for {Script}.xlsx', index=False)


def start():
    print('NOTE: Code should match with MeroShare include Capitalizations.')
    script = input('Script Code to Check : ')
    read_excel(script)
    input('Press Enter to Continue....')


if __name__ == '__main__':
    start()
