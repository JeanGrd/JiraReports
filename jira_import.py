# -*- coding: utf-8 -*-
"""

---------------------------------------------------------------------------------------------------------------

Remark : 

Written by : Jean Guiraud

---------------------------------------------------------------------------------------------------------------

"""

# !/usr/bin/env python3
# Python version : 3.8

from jira.client import JIRA
import pandas as pd

def just_highest_issues(splitter, n_splitter, list_to_split):

    """

    Get the highest issue from a list

    @splitter = the separator
    @n_splitter = the position of the issue name (example : issue_2 the position is 0)
    @list_to_split = your list where you want to keep the highest issue

    """

    list_to_split.sort()
    list_highest_issue = []

    for val in range(len(list_to_split)):

        cut_string1 = list_to_split[val].split(splitter)

        if val == len(list_to_split) - 1:
            list_highest_issue.append(list_to_split[val])
            break

        cut_string2 = list_to_split[val + 1].split(splitter)

        if cut_string1[n_splitter] != cut_string2[n_splitter]:
            list_highest_issue.append(list_to_split[val])

    return list_highest_issue

def jira_import(jira_issues, information):
        
    jira_import= []
        
    if len(jira_issues) != 0:

        for incremental, all_issues in enumerate(jira_issues):  # Browse all issues in the JSON file
        
            jira_import.append({})
        
            for table_field, table_type in information.items(): 
                                                
                # Condition for the multiple values in the JSON variable
                if table_type == "multiplevalues":  # To check that it is called in the XML file
                                    
                    Fulltext= ""  # Variable that adds the names of the multiple values
                    for multiplevalues in eval("all_issues.fields." + table_field):
                        if Fulltext == "":
                            Fulltext = str(multiplevalues.name)
                        else:
                            Fulltext = Fulltext + "\n" + str(multiplevalues.name)
                                        
                    jira_import[incremental][table_field]= Fulltext
                    
                elif table_type == "link":
                    
                    Fulltext = []
            
                    for link in all_issues.fields.issuelinks:
                        if link.type.inward == table_field:
                            if hasattr(link, "inwardIssue"):
                                cut_string = str(link.inwardIssue.fields.summary).split(
                                    "_")  # To get the name of the issue
            
                                Fulltext.append(str(cut_string[0]) + " Iss. " + str(cut_string[1]))
            
            
                        elif link.type.outward == table_field:
                            if hasattr(link, "outwardIssue"):
                                cut_string = str(link.outwardIssue.fields.summary).split(
                                    "_")  # To get the name of the issue
            
                                Fulltext.append(str(cut_string[0]) + " Iss. " + str(cut_string[1]))
            
                    Fulltext = just_highest_issues(" Iss. ", 0, Fulltext)
                    Text = ""
            
                    for Fulltext in Fulltext :
            
                        if Text == "" :
                            Text += Fulltext
                        else :
                            Text += "\n" + Fulltext
                            
                    jira_import[incremental][table_field]= Text
                        
                elif table_type == "summary":
                    cut_string = str(all_issues.fields.summary).split("_")  # Split variable
                    jira_import[incremental][table_field]= cut_string[2]
                
                # For all other fields
                else:                 
                    jira_import[incremental][table_field]= str(eval("all_issues.fields." + table_field))

    yield jira_import


if __name__ == "__main__": 
    
    jira = JIRA(options={'server': ""}, basic_auth=("", ""))
    jira_issues = jira.search_issues('', maxResults=False)
    
    table= {"fixVersions": "multiplevalues","summary": "summary"}
        
    for jira in jira_import(jira_issues, table):
        print(jira)
