from JiraReports import JiraReports
from jira.client import JIRA

if __name__ == "__main__":
    jira = JIRA(options={'server': ""},
                basic_auth=("", ""))

    jira_XML = JiraReports(jira, "test.xml")

    jira_XML.to_excel(document_name="/Users/jean/Desktop/testExcel")
    jira_XML.to_word(document_name="/Users/jean/Desktop/testWordNoTemplate", landscape=True)
    jira_XML.to_word_template(document_name="/Users/jean/Desktop/testWordTemplate", path_template_word="Document.docx")

    print("finished!")
