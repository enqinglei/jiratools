#coding:utf-8
#!/usr/bin/env python
"""
import sys

reload(sys)

sys.setdefaultencoding("utf-8")
"""
from datetime import datetime
from flask import Flask,session, request, flash, url_for, redirect, render_template, abort ,g,send_from_directory,jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_login import login_user , logout_user , current_user , login_required
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import create_engine,and_,or_
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from flask_login import UserMixin
from wtforms import StringField, PasswordField, BooleanField, SubmitField
import xlrd,xlwt
import mysql.connector
import numpy as np
import pandas as pd
from flask import send_file,make_response
from io import BytesIO
import re
import time
import os
from xlrd import xldate_as_tuple
from sqlalchemy import or_
from jira import JIRA
from jira.client import JIRA
import urllib2
import urllib
import webbrowser
#from prontotableupdate import addr_dict,addr_dict1,getjira

import requests
from requests_ntlm import HttpNtlmAuth
from requests.auth import HTTPBasicAuth
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.client_request import ClientRequest
from office365.runtime.utilities.request_options import RequestOptions
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.file import File
import wget

import urllib2

from ntlm import HTTPNtlmAuthHandler




app = Flask(__name__)
#
# <option value="223819">1813</option>
#                                 <option selected="selected" value="211522">1901</option>
#                                 <option value="211523">1902</option>
#                                 <option value="211524">1903</option>
#                                 <option value="211525">1904</option>
#                                 <option value="211526">1905</option>
#                                 <option value="211527">1906</option>
#                                 <option value="211528">1907</option>
#                                 <option value="211529">1908</option>
#                                 <option value="211530">1909</option>
#                                 <option value="211531">1910</option>
#                                 <option value="211532">1911</option>
#                                 <option value="211533">1912</option>
#                                 <option value="223820">1913</option>
#                                 <option value="225511">2001</option>
#                                 <option value="225512">2002</option>
#                                 <option value="225513">2003</option>
#                                 <option value="225514">2004</option>
#                                 <option value="225515">2005</option>
#                                 <option value="225516">2006</option>
#                                 <option value="225517">2007</option>
#                                 <option value="225518">2008</option>
#                                 <option value="225519">2009</option>
#                                 <option value="225520">2010</option>
#                                 <option value="225521">2011</option>
#                                 <option value="225522">2012</option>
#                                 <option value="225523">2013</option>

@app.route('/jira',methods=['GET', 'POST'])
def update2JIRA():
    if request.method == 'GET':
        return render_template('FusionLogin4InternalTaskUpdate.html')
    else:
        username = request.form['username'].strip()
        password = request.form['password'].strip()
        options = {'server': 'https://jiradc.int.net.nokia.com/'}

        jira = JIRA(options, basic_auth=(username, password),max_retries = 2, timeout= 5)
        #Fusion Feature Backlog key is FFB
        # FFB id=u'42894' key =u'FFB'
        issue = jira.issue('FPB-253898')
        issue = jira.issue('FCA_LTE-90')
        dict1 = {'duedate': '2018-12-30'}
        issue.update(dict1)
        # s = issue.fields.customfield_29791
        # t = issue.fields.timetracking.originalEstimate
        parentkey = 'FPB-253899'
        # issue = jira.issue('FCA_LTE-83')
        dict1 = {'customfield_29791': parentkey}
        # dict1 = {'customfield_38699':'LTE'}
        # #customfield_38691 - val value="211474"
        # #customfield_38702 timetracking
        # dict1 = {'customfield_38702': 'CNI-34536-DL Part'}
        # #dict1 = {'customfield_38691': [{'id': u'211474'}]}
        # issue = jira.issue('FCA_LTE-89')
        # issue.update(dict1)
        # t = issue.fields.timetracking
        # #sch = issue.fields.schedule
        # sch = issue.fields.customfield_38751
        # issue = jira.issue('FCA_LTE-83')
        # status = issue.fields.status
        # #dict1 = {'timetracking':{'originalEstimate':'100h','remainingEstimate':'100h'}} #
        # #dict1 = {'timetracking':{'originalEstimate':'100h'}} #
        # dict1 = {'customfield_38751':{'id':'211522'}} #
        # issue.update(dict1)
        # newissue = jira.issue('FCA_LTE-33')
        # project = jira.project('FFB')
        # prjId=project.id
        # project = jira.project('FPB')
        project = jira.project('FCA_LTE')
        prjId = project.id
        # Jira RCA Analysis Subtask
        issueforsubtask = {
            'project': {'id': prjId},
            # 'parent': {"key": "FPB-253898"},
            'summary':'Internal task: MAC000001',
            'issuetype': {'name': 'Epic'},
            'customfield_12791': 'MAC000111',# Epic Name
            'customfield_29791': parentkey, #Parent link
           'customfield_38699':[{'name': 'LTE'}], #System LTE
            'customfield_38700': [{'name': 'eNB'}],  # Pruducts
            'customfield_38702': 'CNI-34536-DL Part', # Item ID
            'customfield_38691': [{'id': u'211474'}],# Fusion Component
            'description': 'Internal Fusion item', #Description
            'timetracking': {'originalEstimate': '100h'},
            'customfield_38751': {'id': '211522'} # Target FB need
           # 'assignee': {'name': 'qmxh38'}
        }
        newissue = jira.create_issue(issueforsubtask)
        prjId = project.id
        MNPRCA = 'MNPRCA'
        TYPE ='5WhyRCA'
        PRID="PR1234"
        PRID = "CAS-95007-C5J8 "
        PRID = "CAS-148023-Z0D4"
        PRID = gPRID
        PRID = "PR1234"
        RcaAnalysis_issues = jira.search_issues(jql_str='project = MNRCA AND summary~'+PRID+' AND type = "Analysis subtask" AND status not in (Resolved, Closed)',maxResults=5)
        Rca5Why_issues = jira.search_issues(jql_str='project = MNRCA AND summary~'+PRID+' AND type = "5WhyRCA"',maxResults=5)
        print RcaAnalysis_issues
        print len(RcaAnalysis_issues)
        link = RcaAnalysis_issues[0].fields.customfield_37064  # SharePoint Link

        #link = "https://nokia.sharepoint.com/:x:/r/sites/LTERCA/_layouts/15/Doc.aspx?sourcedoc=%7B11A7255C-B761-44AC-B631-3549E29AAE69%7D&file=CAS-148023-Z0D4.xlsm&action=default&mobileredirect=true"
        print link
        #link = "https://nokia.sharepoint.com/sites/LTERCA/RcaStore/Forms/AllItems.aspx?RootFolder=%2Fsites%2FLTERCA%2FRcaStore"

        db_url = 'https://jiradc.int.net.nokia.com/secure/EditIssue!default.jspa?id='+issue_id
        ticket_url = 'https://jira3.int.net.nokia.com/browse/' + jira_id
        response = urllib2.urlopen(db_url)
        print response.read()

        todo_item = Todo.query.get(PRID)

        status = RcaAnalysis_issues[0].fields.status
        todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        todo_item.JiraIssueStatus = status
        todo_item.IsRcaCompleted = 'Yes'
        db.session.commit()
        flash('AP item have been successfully imported RCA status also has been updated!')
        return redirect(url_for('show_or_update',PRID=PRID))
    return render_template('ap_new.html')


if __name__ == '__main__':
    #modifyColumn()
    app.run(debug=True,host='0.0.0.0',port=5566)
