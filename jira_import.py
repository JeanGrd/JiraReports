# -*- coding: utf-8 -*-
"""

---------------------------------------------------------------------------------------------------------------

Remark : This code allows to import jira data which are in JSON format to a pandas 
dataframe. This part of program has been developed for the Moko application.

Written by : Jean Guiraud

---------------------------------------------------------------------------------------------------------------

"""

# !/usr/bin/env python3
# Python version : 3.8

from jira.client import JIRA
import pandas as pd
import xlsxwriter
import docx
import xml.etree.ElementTree as ET
import win32com.client as win32


class import_jira_xml(object):
    
    def __init__ (self, xml):
        
        file = ET.parse(xml)  # To parse the content in a variable
        self.root = file.getroot()  # To go to the beginning of the XML document    
     
    def import_jira(self, excel=False, word=False, word_template=False):
        
        if excel is True:
        
            writer= pd.ExcelWriter("CSAR.xlsx", engine="xlsxwriter")
            workbook = writer.book  # Creating an excel document
            
            # Formatting of the excel document
            header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D8E4BC'})
            Bold = workbook.add_format({'bold': True, 'align': 'center'})
            BoldNoCenter = workbook.add_format({'bold': True})
            # To give the title of the document according to the XML file
            titre = workbook.add_format({'align': 'center', 'bold': True})
            
            worksheet_name = []
        
        if word_template is True:
        
            # open an existing document
            document = docx.Document("CSAR-TEMPLATE.docx")
            tables = document.tables 
        
        if word is True:
            pass
        
        for incremental, all_tables in enumerate(self.root.findall('Table')):
            
            jira = JIRA(options={'server': "https://kodo:8443/"}, basic_auth=("guirauj1", "Z@CtU1404"))
            
            if excel is True:
            
                table= {}
                for columns in all_tables.findall("Column"):
                    table[columns.text]= columns.get("type")
                            
                Name = all_tables.get("name")
                tag = ["/", "*", ":", "[", "]"]
            
                for tags in tag:  # For prohibited characters
                    Name = Name.replace(tags, ' ')
            
                Name = str(incremental+1) + " - " + Name
            
                if len(Name) > 31:  # For character length
                    Name = Name[:31]
            
                worksheet_name.append(Name)
                
            if word_template is True:
                
                for searchtables in range(len(tables)):
                    try:
                        if all_tables.get("keyword") == tables[searchtables].cell(1, 0).text:
                            tableslenght = searchtables
                    except:
                        continue
        
                    searchtables += 1
                    
            if word is True:
                pass
                        
            if all_tables.get("style") == "MultipleFilters":
                for filters in all_tables.findall("Filters"):
                    for nbJQL in filters.findall("JQL"):
                        print(nbJQL.get("name")) 
                        
            if all_tables.get("style") == "Classic":
                
                jira_issues = jira.search_issues(all_tables.find('JQL').text, maxResults=False)
                
                if excel is True:
                         
                    df= pd.DataFrame(self.jira_import(jira_issues, table))
                    df.to_excel(writer, sheet_name=Name, startrow=3, header=False, index=False)
                
                    # To put the name on the Excel sheet
                    worksheet = writer.sheets[Name]
                    worksheet.write(0, 0, all_tables.get('name'), titre)
                
                    for incremental, column in enumerate(all_tables.findall('Column')):  # Creation of all the columns
                        worksheet.write(2, incremental, column.get('name'), header)
                
                    worksheet.set_row(2, 30)
        
                if word_template is True:
     
                    # add the rest of the data frame
                    for i in range(df.shape[0]):
                        if i != 0 :
                            tables[tableslenght].add_row()
                        for j in range(df.shape[-1]):
                            tables[tableslenght].cell(i+1,j).text = str(df.values[i,j])
                            
                if word is True:
                    pass
            
        if excel is True:
            writer.close()
        
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open("C:/Users/guirauj1/Desktop/WanderingStars/GAL_PROJECT/CSAR.xlsx")
    
            for name in worksheet_name:
                ws = wb.Worksheets(name)
                ws.Columns.AutoFit()
        
            wb.Save()
    
        if word_template is True:
            # save the doc
            document.save('df.docx')


    def __just_highest_issues__(self, splitter, n_splitter, list_to_split):
    
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
    
    def jira_import(self, jira_issues, information):
            
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
                
                        Fulltext = self.__just_highest_issues__(" Iss. ", 0, Fulltext)
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
    
        return jira_import
            
    
if __name__ == "__main__": 
        
    jira = import_jira_xml("CSAR.xml")
    jira.import_jira(excel=True, word_template=True)
    
