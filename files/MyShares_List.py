import logging
from tenacity import retry, stop_after_attempt, wait_fixed
import requests
import json
import os
import pandas as pd
from openpyxl import load_workbook
import datetime

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
            client_id: str = None
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
                        print(login_req.json()['message'])
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

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(3), reraise=True)
    def get_share_list(self):
        try:
            with self.__session as sess:
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
                data = json.dumps(
                    {
                        "sortBy": "CCY_SHORT_NAME",
                        "demat": [self.__dmat],
                        "clientCode": self.__dpid,
                        "page": 1,
                        "size": 200,
                        "sortAsc": "true"
                    }
                )
                try:
                    details_json = sess.post("https://webbackend.cdsc.com.np/api/meroShareView/myShare/", data=data, verify=cafile).json()

                    pd_list = pd.DataFrame(columns=['Client ID', 'Name', 'DMAT No', 'Script', 'Current Balance', 'Free Balance'])
                    for item in details_json.get("meroShareDematShare"):
                        logging.info(
                            f'Script: {item.get("script")}, CurrentBalance: {item.get("currentBalance")}, for account: {self.__name}'
                        )
                        
                        data = [self.client_id] + [self.__name] + [self.__dmat] + [item['script']] + [item['currentBalance']] + [item['freeBalance']]
                        pd_list.loc[len(pd_list)] = data
                        
                    return pd_list
                except:
                    print('Share list Connection Refused')
                    self.status = 'Share list Connection Refused!'

        except Exception as error:
            logging.info(error)
            logging.error(
                f"Details request failed!!"
            )


def check(sheet, full_list, client_type):
    for details in sheet.iter_rows(min_row=2,
                                   min_col=1,
                                   values_only=True):
        if str(details[2]).upper()=='NO':
            continue

        login_info = {"name": details[1],
                      "username": details[4].replace(" ", ""),
                      "password": details[6],
                      "dpid": int(details[5]) - 13000000,
                      "client_id": client_type + str(details[0])}
        ms = MeroShare(**login_info)
        ms.login()
        if ms.status is None:
            full_list = pd.concat([full_list, ms.get_share_list()])
            if ms.status is not None:
                data = [client_type + str(details[0])] + [details[1]] + [str(details[5]) + str(details[4].replace(' ', ''))] + [ms.status] + [0, 0]
                full_list.loc[len(full_list)] = data
        else:
            data = [client_type + str(details[0])] + [details[1]] + [str(details[5])+str(details[4].replace(' ',''))] + [ms.status] + [0, 0]
            full_list.loc[len(full_list)] = data

    return (full_list)


def check_share():
    full_list = pd.DataFrame(columns=['Client ID', 'Name', 'DMAT No', 'Script', 'Current Balance', 'Free Balance'])

    book = load_workbook(filename='MeroShare Login Details.xlsx', data_only=True)

    full_list = check(book['List'], full_list, '')
    
    print(full_list)
    full_list.to_excel(f'MeroShare - Share List - {datetime.datetime.now().strftime("%d-%b-%Y")}.xlsx', index=False)

    input('Press Enter to Continue....')

if __name__ == '__main__':
    check_share()
