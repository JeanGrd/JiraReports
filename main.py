import JiraReports
from jira.client import JIRA

if __name__ == "__main__":

    jira = JIRA(options={'server': ""},
                basic_auth=("", ""))

    jira_XML = JiraReports(jira, "test.xml")

    jira_XML.to_word(document_name="/Users/jean/Desktop/test", landscape=True)
    jira_XML.to_excel(document_name="/Users/jean/Desktop/t")
    print("finished!")
