import logging
from tenacity import retry, stop_after_attempt, wait_fixed
import requests
import json
import pandas as pd

logging.basicConfig(format="%(asctime)s %(message)s", level=logging.INFO)

BaseURL_ = "https://webbackend.cdsc.com.np/api"

headers_ = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Authorization": "null",
    "Connection": "keep-alive",
    "Content-Type": "application/json",  # important for login?
    "DNT": "1",
    "Origin": "https://meroshare.cdsc.com.np",
    "Referer": "https://meroshare.cdsc.com.np/",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-site",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36",
}

cap_file = "files/capitals.json"
ca_file = "files/cdsc-com-np-chain.pem"


def update_capital_list():
    response = requests.request(
        "GET", f"{BaseURL_}/meroShare/capital/", headers=headers_
    )
    with open(cap_file, "w") as cap_file_:
        json.dump(response.json(), cap_file_)


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
        bank: str = None,
    ):

        self.__name = name
        self.__dpid = str(dpid)
        self.__username = str(username)
        self.__password = password
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

        self.__session = requests.Session()
        self.__session.headers.update(headers_)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(3), reraise=True)
    def get_capital_id(self):
        updated = False
        while True:
            try:
                return [
                    response["id"]
                    for response in json.load(open(cap_file))
                    if response["code"] == self.__dpid
                ][0]
            except Exception as error:
                logging.info(f"Error finding Capital for Acc: {self.__name}")
                logging.error(error)
                if updated:
                    return None
                logging.info("Updating Capitals List Cache")
                update_capital_list()
                updated = True

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(3), reraise=True)
    def login(self) -> bool:
        assert (
            self.__username and self.__password and self.__dpid
        ), "Username, password and DPID required!"

        if not self.__capital_id:
            logging.warning("Capital ID Error")
            self.status = "Problem Finding Capital"
            return False

        with self.__session as sess:
            data = json.dumps(
                {
                    "clientId": self.__capital_id,
                    "username": self.__username,
                    "password": self.__password,
                }
            )

            try:
                login_req = sess.post(
                    f"{BaseURL_}/meroShare/auth/", data=data, verify=ca_file
                )
                self.status = login_req.json()["message"]

                if login_req.status_code == 200:
                    self.__auth_token = login_req.headers.get("Authorization")
                    logging.info(f"{self.status}  Account: {self.__name}!")
                else:
                    logging.warning(f"{self.status} for Account: {self.__name}")
            except Exception as error:
                logging.info(error)
                logging.info(
                    f"Retrying login: ({self.login.retry.statistics.get('attempt_number')})!"
                )
                self.status = f"Login Failed!! {error}"

        if not self.__auth_token:
            return False

        headers = {"Authorization": self.__auth_token}
        self.__session.headers.update(headers)

        return True

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(3), reraise=True)
    def get_share_list(self):
        with self.__session as sess:
            data = json.dumps(
                {
                    "sortBy": "CCY_SHORT_NAME",
                    "demat": [self.__dmat],
                    "clientCode": self.__dpid,
                    "page": 1,
                    "size": 200,
                    "sortAsc": "true",
                }
            )

            try:
                myShare = sess.post(
                    f"{BaseURL_}/meroShareView/myShare/", data=data, verify=ca_file
                )

                pd_list = pd.DataFrame(
                    columns=[
                        "Client ID",
                        "Name",
                        "DMAT No",
                        "Script",
                        "Current Balance",
                        "Free Balance",
                    ]
                )

                for item in myShare.json()["meroShareDematShare"]:
                    logging.info(
                        f'Account: {self.__name} -> Script: {item.get("script")}, CurrentBalance: {item.get("currentBalance")}, FreeBalance: {item["freeBalance"]}'
                    )
                    data = (
                        [self.client_id]
                        + [self.__name]
                        + [self.__dmat]
                        + [item["script"]]
                        + [item["currentBalance"]]
                        + [item["freeBalance"]]
                    )
                    pd_list.loc[len(pd_list)] = data

                return pd_list
            except Exception as error:
                self.status = "Error Getting MyShare List"
                logging.info(self.status)
                logging.error(error)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(3), reraise=True)
    def get_applicable_issues(self):
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

            try:
                self.__applicable_issues = (
                    sess.post(
                        f"{BaseURL_}/meroShare/companyShare/applicableIssue/",
                        data=data,
                        verify=ca_file,
                    )
                    .json()
                    .get("object")
                )
            except Exception as error:
                logging.error(error)
                self.status = f"Applicable issues request failed! {error}"
                logging.info({self.status})

            logging.info(f"Appplicable Issues Obtained! Account: {self.__name}")
            return self.__applicable_issues

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(3), reraise=True)
    def apply(self, share_id: str, qty: int):
        try:
            with self.__session as sess:
                issue_to_apply = None

                if not self.__applicable_issues:
                    self.get_applicable_issues()

                for issue in self.__applicable_issues:
                    if str(issue.get("scrip")) == share_id:
                        issue_to_apply = issue
                        break

                if not issue_to_apply:
                    logging.warning(
                        "Provided Script doesn't match any of the applicable issues!"
                    )
                    self.status = "No matching applicable issues!"
                    return self.status

                share_id = issue_to_apply.get("companyShareId")

                if issue_to_apply.get("action"):
                    status = issue_to_apply.get("action")
                    self.status = "Couldn't apply for issue! - " + status
                    logging.info(self.status)
                    return self.status

                bank_req = sess.get(
                    f"{BaseURL_}/meroShare/bank/", verify=ca_file
                ).json()

                bank_id = None
                for bank_ in bank_req:
                    if bank_["name"] == self.bank:
                        bank_id = bank_["id"]

                if bank_id is None:
                    self.status = "Bank name not found."
                    print(self.status)
                    return self.status

                bank_specific_req = sess.get(
                    f"{BaseURL_}/meroShare/bank/{bank_id}", verify=ca_file
                )

                bank_specific_response_json = bank_specific_req.json()[0]

                data = json.dumps(
                    {
                        "accountBranchId": bank_specific_response_json.get(
                            "accountBranchId"
                        ),
                        "accountNumber": bank_specific_response_json.get(
                            "accountNumber"
                        ),
                        "accountTypeId": bank_specific_response_json.get(
                            "accountTypeId"
                        ),
                        "appliedKitta": qty,
                        "bankId": bank_id,
                        "boid": self.__dmat[-8:],
                        "companyShareId": share_id,
                        "crnNumber": self.__crn,
                        "customerId": bank_specific_response_json.get("id"),
                        "demat": self.__dmat,
                        "transactionPIN": self.__pin,
                    }
                )

                apply_req = sess.post(
                    f"{BaseURL_}/meroShare/applicantForm/share/apply",
                    data=data,
                    verify=ca_file,
                )

                self.status = apply_req.json()["message"]

                logging.info(self.status)

                if apply_req.status_code == 201:
                    logging.info(
                        f"Application Successful! for account: {self.__name}, {qty} Kitta"
                    )

                return self.status

        except Exception as error:
            logging.info(error)
            self.status = f"Apply failed! - {error}"
            return self.status

    @retry(stop=stop_after_attempt(10), wait=wait_fixed(3), reraise=True)
    def get_application_status(self, scrip: str):
        with self.__session as sess:
            try:
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

                try:
                    recent_applied_req = sess.post(
                        f"{BaseURL_}/meroShare/applicantForm/active/search/",
                        data=data,
                        verify=ca_file,
                    )
                except Exception as error:
                    logging.error(error)
                    raise error

                target_issue = None

                if recent_applied_req.status_code == 200:
                    for issue in recent_applied_req.json()["object"]:
                        if issue["scrip"] == scrip:
                            target_issue = issue
                else:
                    self.status = "Application list request failed."
                    logging.info(self.status)
                    raise self.status

                if not target_issue:
                    self.status = "Script not found!"
                    return self.status

                try:
                    details_req = sess.get(
                        f"{BaseURL_}/meroShare/applicantForm/report/detail/{target_issue['applicantFormId']}",
                        verify=ca_file,
                    ).json()["statusName"]
                    self.status = details_req
                except Exception as error:
                    logging.error(error)
                    self.status = "Report rqeuest Failed"
                    logging.info(self.status)

                logging.info(f"Status: {self.status} for {self.__name}")
                return details_req

            except Exception as error:
                logging.error(error)
                self.status = "Application status request failed."
                logging.warning(
                    f"Application status request failed! Retrying ({self.get_application_status.retry.statistics.get('attempt_number')})"
                )
                return 0
