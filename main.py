import datetime
import json
import win32com.client as client
from requests.auth import HTTPBasicAuth
from Config.html_body import HTML_BODY
import pandas as pd
from jira import JIRA
from configparser import ConfigParser
import requests

config = ConfigParser()
config.read('Config/config.ini')

api_token = config['JIRA_Config']['apitoken']
email = config['JIRA_Config']['email']
server = config['JIRA_Config']['server']
projects = config['Project']['project']
tomail = config['Project']['tomail']

class JiraTool:
    def __init__(self):
        data = pd.read_csv("Config/JIRA FSI Users list.csv")
        filename = self.getfilename()
        self.employeeList = data["User name"].tolist()
        print(self.employeeList)
        self.jira = JIRA(basic_auth=(email, api_token), server=server)
        jquery = f"project in ({projects}) AND status in (Open, Triage, Groomed, \"To do\", \"More Information\", \"In Development\") ORDER BY priority DESC"
        self.tickets = []
        startat = 0
        try:
            while True:
                issues = self.jira.search_issues(jql_str=jquery, maxResults=1000, startAt=startat)
                self.addticketlist(issues)
                startat = startat + 100
                issues = self.jira.search_issues(jql_str=jquery, maxResults=1000, startAt=startat)
                if len(issues) == 0:
                    break
            print(self.tickets.__len__())
            df = pd.DataFrame(data=self.tickets)
            writer = pd.ExcelWriter(filename, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1')
            writer.save()
            self.sendmail(filename)
        except Exception as e:
            print(e)

    def addticketlist(self, issues):
        if issues:
            for issue in issues:
                comments = self.jira.comments(issue, expand="properties")
                if comments:
                    comments.reverse()
                    for comment in comments:
                        if self.getexternalcomment(issue, comment):
                            if comment.author.displayName in self.employeeList:
                                date = abs(self.getdate(comment.created).days)
                                key = issue
                                issueType = str(issue.fields.issuetype)
                                summary = str(issue.fields.summary)
                                priority = str(issue.fields.priority)
                                reporter = str(issue.fields.reporter)
                                assignee = str(issue.fields.assignee)
                                created = self.formatdate(str(issue.fields.created))
                                updated = self.formatdate(str(issue.fields.updated))
                                status = str(issue.fields.status)
                                author = str(comment.author.displayName)
                                temp = {
                                    "Key": key,
                                    "issueType": issueType,
                                    "summary": summary,
                                    "priority": priority,
                                    "reporter": reporter,
                                    "assignee": assignee,
                                    "created": created,
                                    "updated": updated,
                                    "status": status,
                                    "lastExternalCommentDate": date,
                                    "last External Comment Author": author
                                }
                                print(temp)
                                self.tickets.append(temp)
                                break

                pass
        pass

    @staticmethod
    def getexternalcomment(issueKey, commentId):
        """
            gets the type of comment whether it's Internal or Public
        :param issueKey: Issue ID (string)
        :param commentId: Comment ID (string)
        :return: Type of comment ('True/False')
        """
        url = f"{server}/rest/api/3/issue/{issueKey}/comment/{commentId}"
        auth = HTTPBasicAuth(email, api_token)
        headers = {
            "Accept": "application/json",
        }
        response = requests.get(url, headers=headers, auth=auth)
        data = json.loads(response.text)
        if data['jsdPublic']:
           return True
        else:
            return False

    @staticmethod
    def getdate(string):
        """
        converts STRING object to a DATE object in UTC time zone
        :param string: date (string)
        :return: Date object
        """
        year = int(string[0:4])
        month = int(string[5:7])
        day = int(string[8:10])
        hour = int(string[11:13])
        minute = int(string[14:16])
        seconds = int(string[17:19])
        TZD = string[23:28]
        date = datetime.datetime(year=year, month=month, day=day, hour=hour, minute=minute, second=seconds)
        if TZD == "-0800":
            final = date + datetime.timedelta(hours=8)
        else:
            final = date + datetime.timedelta(hours=7)
        return final - datetime.datetime.today()

    @staticmethod
    def diffdate(date1):
        diffday = date1 - datetime.datetime.today()
        return diffday.days

    @staticmethod
    def getmail(user):
        """
             Generates mail ID with the help of parameter
        :param user: name of the person to whom mail ID is to be generated (string)
        :return: Mail ID (string)
        """
        temp = user.replace('.', '').split(" ")
        count = 0
        ename = ""
        n = len(temp)
        for name in temp:
            if count != n - 1:
                ename = ename + name
            else:
                ename = ename + "." + name
            count = count + 1
        email = ename + "@flatironssolutions.com"
        return email.lower()

    @staticmethod
    def formatdate(date):
        """
            To Change the Date Format for our convenience
            :param date: date(string)
        """
        year = int(date[0:4])
        month = int(date[5:7])
        day = int(date[8:10])
        formatedDate = f"{day}-{month}-{year}"
        return formatedDate

    @staticmethod
    def sendmail(filename):
        try:
            outlook = client.Dispatch("Outlook.Application")
            message = outlook.CreateItem(0)
            message.To = tomail
            message.Attachments.Add(filename)
            attachment = message.Attachments.Add("D:\python examples\logo.png")
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_img")
            message.Subject = "Jira Reminder Tool "
            message.Send()
        except Exception as e:
            print(e)
            pass

    @staticmethod
    def getfilename():
        today = datetime.date.today()
        fileName = "TamReport-"+str(today)+".xlsx"
        print(fileName)
        return fileName

if __name__ == "__main__":
    obj = JiraTool()
