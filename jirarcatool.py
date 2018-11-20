#coding:utf-8
#!/usr/bin/env python
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
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
from jira import JIRA
from jira.client import JIRA
import urllib2
import urllib
import webbrowser
from rcatrackingconfig import addr_dict,addr_dict1,getjira,get_login_info

import requests
from requests_ntlm import HttpNtlmAuth
from requests.auth import HTTPBasicAuth
from office365.runtime.client_request import ClientRequest
from office365.runtime.utilities.request_options import RequestOptions
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.file import File
from os.path import basename
import shutil

import ldap
from flask_wtf import Form
from wtforms import TextField, PasswordField
from wtforms.validators import InputRequired
from config import PriorityDict,SourceListDict,BusinessUnitDict,BusinessLineDict,ProductLineDict,CustomerNameDict


app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root@localhost:3306/fddrca?charset=utf8'
app.config['SQLALCHEMY_COMMIT_ON_TEARDOWN'] = True
app.config['SQLALCHEMY_ECHO'] = False
app.config['SECRET_KEY'] = 'secret_key'
app.config['DEBUG'] = True
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False


UPLOAD_FOLDER = 'upload'
UPLOAD_FOLDER1 = 'apupload'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
basedir = os.path.abspath(os.path.dirname(__file__))
ALLOWED_EXTENSIONS = set(['txt', 'png', 'jpg', 'xls', 'JPG', 'PNG', 'xlsx', 'gif', 'GIF','xlsm'])
teams=['chenlong','xiezhen','yangjinyong','zhangyijie','lanshenghai','liumingjing','lizhongyuan','caizhichao','hujun','wangli']

app.config['ldap'] = ''
db = SQLAlchemy(app)

Base=declarative_base()
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

gSSOPWD ='Happy&2018'
def get_ldap_connection():
    # conn = ldap.initialize(app.config['LDAP_PROVIDER_URL'])
    try:
        conn = ldap.initialize('ldap://ed-p-gl.emea.nsn-net.net')
        conn.simple_bind_s('cn=BOOTMAN_Acc,ou=SystemUsers,ou=Accounts,o=NSN', 'Eq4ZVLXqMbKbD4th')
    except:
        return redirect(url_for('login'))

    return conn

class User(db.Model):
    __tablename__ = "jirausers"
    id = db.Column('user_id',db.Integer , primary_key=True)
    username = db.Column('username', db.String(64), unique=True , index=True)
    password = db.Column('password' , db.String(256))
    email = db.Column('email',db.String(128),unique=True , index=True)
    displayName = db.Column(db.String(128))
    lineManagerAccountId = db.Column(db.String(128)) #nsnManagerAccountName
    lineManagerDisplayName = db.Column(db.String(128)) #nsnManagerName
    lineManagerEmail = db.Column(db.String(128)) # Further search thru uid.
    squadGroupName = db.Column(db.String(128))  # Further search thru uid.

    registered_on = db.Column('registered_on' , db.DateTime)
    todos = db.relationship('Todo', backref='user', lazy='select')

    def __init__(self , username, email,displayName,lineManagerAccountId,lineManagerDisplayName,lineManagerEmail,squadGroupName):
        self.username = username
        self.email = email
        self.displayName = displayName
        self.lineManagerAccountId = lineManagerAccountId
        self.lineManagerDisplayName = lineManagerDisplayName
        self.lineManagerEmail = lineManagerEmail
        self.squadGroupName = squadGroupName
        self.registered_on = datetime.utcnow()

    # def get_ldap_connection():
    #     # conn = ldap.initialize(app.config['LDAP_PROVIDER_URL'])
    #     conn = ldap.initialize('ldap://ed-p-gl.emea.nsn-net.net')
    #     conn.simple_bind_s('cn=BOOTMAN_Acc,ou=SystemUsers,ou=Accounts,o=NSN', 'Eq4ZVLXqMbKbD4th')
    #     return conn

    @staticmethod
    def try_login(username, password):
        global gSSOPWD
        try:
            conn = get_ldap_connection()
        except:
            return redirect(url_for('login'))
            return render_template('login.html')
        app.config['ldap'] = conn
        filter = '(uid=%s)'%username
        attrs = ['sn', 'mail', 'cn', 'displayName', 'nsnManagerAccountName']
        base_dn = 'o=NSN'
        try:
            result = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, filter, attrs)
        except:
            return redirect(url_for('login'))
        dn = result[0][0]
        try:
            a = conn.simple_bind_s(dn,password)
            gSSOPWD = password # For privacy policy, cannot save this for other purpose.
            print 'Cookie auto login or Manual login?????'
        except ldap.INVALID_CREDENTIALS:
            print "Your username or password is incorrect."
            sys.exit()
        except ldap.LDAPError, e:
            if type(e.message) == dict and e.message.has_key('desc'):
                print e.message['desc']
            else:
                print e
            sys.exit()
            return redirect(url_for('login'))
        except:
            return redirect(url_for('login'))
        return result, conn
    def is_authenticated(self):
        return True
    def is_active(self):
        return True
    def is_anonymous(self):
        return False
    def get_id(self):
        return unicode(self.id)
    def __repr__(self):
        return '<User %r>' % (self.username)

class Todo(db.Model):
    __tablename__ = 'jirarcatable1'
    JiraIssueId = db.Column('JiraIssueId', db.String(64), primary_key=True)
    JiraIssuePriority = db.Column(db.String(64)) # Parent task has, Child task does not have
    JiraIssueSourceList = db.Column(db.String(64))  # Parent task has, Child task does not have
    JiraIssueBusinessUnit = db.Column(db.String(128))  # Parent task has, Child task does not have
    JiraIssueBusinessLine = db.Column(db.String(128)) # Parent task has, Child task does not have
    JiraIssueProductLine = db.Column(db.String(128))  # Parent task has, Child task does not have
    JiraIssueCustomerName = db.Column(db.String(128))  # Parent task has, Child task does not have
    JiraIssueFeature = db.Column(db.String(128))
    JiraIssueFeatureComponent = db.Column(db.String(128))
    JiraIssueOther = db.Column(db.String(256))  # Parent task has, Child task does not have
    JiraIssueType = db.Column(db.String(64))
    JiraIssueCaseType = db.Column(db.String(64)) # Child task has, Parent task does not have
    JiraIssueStatus = db.Column(db.String(64))
    JiraIssueLabels = db.Column(db.String(256))
    JiraIssueAssignee = db.Column(db.String(128))
    JiraIssueReporter = db.Column(db.String(128))
    PRID = db.Column('PRID', db.String(64))
    PRTitle = db.Column(db.String(1024))
    PRRcaEdaAssessor = db.Column(db.String(128))
    PRRelease = db.Column(db.String(128))
    PRAttached = db.Column(db.String(128))
    PRSeverity = db.Column(db.String(32))
    PRGroupInCharge = db.Column(db.String(64))
    PRProduct = db.Column(db.String(64))
    ReportedBy = db.Column(db.String(64))
    FaultCoordinator = db.Column(db.String(64))
    JiraIssueSummary = db.Column(db.String(512))
    CustomerName = db.Column(db.String(128))
    user_id = db.Column(db.Integer, db.ForeignKey('jirausers.user_id'))

    def __init__(self, JiraIssueId,JiraIssuePriority,JiraIssueSourceList,JiraIssueBusinessUnit,\
                 JiraIssueBusinessLine,JiraIssueProductLine,JiraIssueCustomerName,JiraIssueFeature,JiraIssueFeatureComponent,\
                 JiraIssueOther,JiraIssueType,JiraIssueCaseType,JiraIssueStatus,JiraIssueLabels,JiraIssueAssignee,\
                 JiraIssueReporter,PRID,PRTitle,PRRelease,PRSeverity,PRGroupInCharge,PRAttached,PRRcaEdaAssessor,\
                 PRProduct,ReportedBy,FaultCoordinator,JiraIssueSummary,CustomerName):
        self.JiraIssueId = JiraIssueId
        self.JiraIssuePriority = JiraIssuePriority
        self.JiraIssueSourceList = JiraIssueSourceList
        self.JiraIssueBusinessUnit = JiraIssueBusinessUnit
        self.JiraIssueBusinessLine = JiraIssueBusinessLine
        self.JiraIssueProductLine = JiraIssueProductLine
        self.JiraIssueCustomerName = JiraIssueCustomerName
        self.JiraIssueFeature = JiraIssueFeature
        self.JiraIssueFeatureComponent = JiraIssueFeatureComponent
        self.JiraIssueOther = JiraIssueOther
        self.JiraIssueType = JiraIssueType
        self.JiraIssueCaseType = JiraIssueCaseType
        self.JiraIssueStatus = JiraIssueStatus
        self.JiraIssueLabels = JiraIssueLabels
        self.JiraIssueAssignee = JiraIssueAssignee
        self.JiraIssueReporter = JiraIssueReporter
        self.PRID = PRID
        self.PRTitle = PRTitle
        self.PRRcaEdaAssessor = PRRcaEdaAssessor
        self.PRRelease = PRRelease
        self.PRAttached = PRAttached
        self.PRSeverity = PRSeverity
        self.PRGroupInCharge = PRGroupInCharge
        self.PRProduct = PRProduct
        self.ReportedBy = ReportedBy
        self.FaultCoordinator = FaultCoordinator
        self.JiraIssueSummary = JiraIssueSummary
        self.CustomerName = CustomerName

"""
    conn=mysql.connector.connect(host='localhost',user='root',passwd='',port=3306)
    cur=conn.cursor()
    cur.execute('create database if not exists fddrca')
    conn.commit()
    cur.close()
    conn.close()
    """


@app.route('/',methods=['GET','POST'])
@login_required
def home():
    if request.method=='GET':
        return render_template('home.html',user=User.query.get(g.user.id).displayName)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template('login.html')

    username = request.form['username']
    password = request.form['password']
    remember_me = False
    if 'remember_me' in request.form:
        # remember_me = BooleanField('Keep me logged in')
        remember_me = True
    try:
        result,conn = User.try_login(username, password)
    except ldap.INVALID_CREDENTIALS:
        flash(
            'Invalid username or password. Please try again.',
            'danger')
        return redirect(url_for('login'))
    except:
        flash(
            'Network Issue.Cannot connect to LDAP server, Please check the Network and try again.',
            'danger')
        return redirect(url_for('login'))

    email = result[0][1]['mail'][0]
    displayName = result[0][1]['displayName'][0]
    lineManagerAccountId = result[0][1]['nsnManagerAccountName'][0]
    user = User.query.filter_by(username=username).first()
    if not user:
        filter = '(uid=%s)'%lineManagerAccountId
        attrs = ['sn', 'mail', 'cn', 'displayName', 'nsnManagerAccountName']
        base_dn = 'o=NSN'
        lineResult = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, filter, attrs)
        lineManagerDisplayName = lineResult[0][1]['displayName'][0]
        lineManagerEmail = lineResult[0][1]['mail'][0]
        link = "http://tdlte-report-server.china.nsn-net.net/api/get_user_info?u_id=%s" %username
        r = requests.get(link)
        c = r.json()
        d = c['sg_name']
        squadGroupName = d
        user = User(username,email,displayName,lineManagerAccountId,lineManagerDisplayName,lineManagerEmail,squadGroupName)
        db.session.add(user)
        db.session.commit()
    else:
        lineManagerAccountId = result[0][1]['nsnManagerAccountName'][0]
        filter = '(uid=%s)'%lineManagerAccountId
        attrs = ['sn', 'mail', 'cn', 'displayName', 'nsnManagerAccountName']
        base_dn = 'o=NSN'
        lineResult = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, filter, attrs)
        user.lineManagerDisplayName = lineResult[0][1]['displayName'][0]
        user.lineManagerEmail = lineResult[0][1]['mail'][0]
        link = "http://tdlte-report-server.china.nsn-net.net/api/get_user_info?u_id=%s" %username
        r = requests.get(link)
        c = r.json()
        d = c['sg_name']
        user.squadGroupName = d
        db.session.commit()

    login_user(user, remember=remember_me)
    flash('%s  --Logged in successfully!'%displayName)
    redirect(url_for('home'))
    return redirect(request.args.get('next') or url_for('index'))

@app.route('/logout')
def logout():
    logout_user()
    return redirect(url_for('index'))

@login_manager.user_loader
def load_user(id):
    return User.query.get(int(id))

@app.before_request
def before_request():
    g.user = current_user


@app.route('/index', methods=['GET', 'POST'])
@login_required
def index():
    return redirect(url_for('rca_home'))

@app.route('/loadAssigneeInfo', methods=['GET', 'POST'])
@login_required
def loadAssigneeInfo():
    if request.method == 'POST':
        AssignTo = request.form['AssignTo'].strip()
        user = User.query.filter_by(email = AssignTo).first()
        if user:
            AssignTo = user.displayName
        else:
            try:
                conn = get_ldap_connection()
            except:
                return redirect(url_for('login'))
            filter = '(mail=%s)'%AssignTo
            attrs = ['sn', 'mail', 'cn', 'displayName', 'nsnManagerAccountName']
            base_dn = 'o=NSN'
            #conn = app.config['ldap']
            try:
                result = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, filter, attrs)
            except:
                return redirect(url_for('logout'))
            AssignTo = result[0][1]['displayName'][0]
            email = result[0][1]['mail'][0]
            displayName = result[0][1]['displayName'][0]
            lineManagerAccountId = result[0][1]['nsnManagerAccountName'][0]
            # user = User.query.filter_by(username=username).first()
            # if not user:
            filter = '(uid=%s)' % lineManagerAccountId
            attrs = ['sn', 'mail', 'cn', 'displayName', 'nsnManagerAccountName']
            base_dn = 'o=NSN'
            lineResult = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, filter, attrs)
            lineManagerDisplayName = lineResult[0][1]['displayName'][0]
            lineManagerEmail = lineResult[0][1]['mail'][0]
            link = "http://tdlte-report-server.china.nsn-net.net/api/get_user_info?u_id=%s" % username
            r = requests.get(link)
            c = r.json()
            d = c['sg_name']
            squadGroupName = d
            user = User(username, email, displayName, lineManagerAccountId, lineManagerDisplayName, lineManagerEmail,
                        squadGroupName)
            db.session.add(user)
            db.session.commit()
        return jsonify(AssignTo=AssignTo)

@app.route('/loadProntoInfo', methods=['GET', 'POST'])
@login_required
def loadProntoInfo():
    global gSSOPWD
    if request.method == 'POST':
        PRID = request.form['PRID'].strip()
        todo = Todo.query.filter_by(PRID=PRID).first()
        PRTitle ='This is test AAAA'
        #return jsonify(PRTitle=PRTitle)
        # if todo:
        #     flash('JIRA 5WhyRca Task for this PRID has been create!', 'error')
        #     return render_template('new.html',PRID=PRID)
        url = "https://pronto.inside.nsn.com/prontoapi/rest/api/latest/problemReport/%s" % PRID
        try:
            r = requests.get(url, verify=False, auth=(g.user.username, gSSOPWD))
        except:
            flash(' %s -- Invalid PRID has been input!'%PRID, 'error')
            return render_template('new.html', PRID=PRID)
        a = r.json()
        return jsonify(a)


@app.route('/new', methods=['GET', 'POST'])
@login_required
def new():
    global gSSOPWD
    if request.method == 'POST':
        value = request.form['button']
        if value == 'CreateAndAssignBtn':
            if not request.form['PRID']:
                flash('PRID is required', 'error')
            PRID = request.form['PRID'].strip()
            todo = Todo.query.filter_by(PRID=PRID).first()
            # if todo:
            #     flash('JIRA 5WhyRca Task for this PRID has been create!', 'error')
            #     return render_template('new.html',PRID=PRID)
            url = "https://pronto.inside.nsn.com/prontoapi/rest/api/latest/problemReport/%s" % PRID
            try:
                r = requests.get(url, verify=False, auth=(g.user.username, gSSOPWD))
            except:
                flash(' %s -- Invalid PRID has been input!'%PRID, 'error')
                return render_template('new.html', PRID=PRID)
            a = r.json()
            PRAttached =''
            for s in a['problemReportIds']:
                PRAttached = PRAttached + s
            JiraIssuePriority = request.form['Priority'].strip()  # Parent task has, Child task does not have
            JiraIssueSourceList = request.form['SourceList'].strip()  # Parent task has, Child task does not have
            JiraIssueBusinessUnit = request.form['BusinessUnit'].strip()  # Parent task has, Child task does not have
            JiraIssueBusinessLine = request.form['BusinessLine'].strip()  # Parent task has, Child task does not have
            JiraIssueProductLine = request.form['ProductLine'].strip()  # Parent task has, Child task does not have
            JiraIssueCustomerName = request.form['Customer'].strip()  # Parent task has, Child task does not have
            JiraIssueFeature = request.form['Feature'].strip()
            JiraIssueFeatureComponent = request.form['FeatureComponent'].strip()
            JiraIssueOther = request.form['Other'].strip()  # Parent task has, Child task does not have
            JiraIssueType = '5WhyRCA'
            JiraIssueCaseType = ''
            JiraIssueLabels = request.form['Labels'].strip().split()

            JiraIssueAssignee = request.form['AssignTo'].strip()
            filter = '(uid=%s)' % g.user.username
            attrs = ['sn', 'mail', 'cn', 'displayName', 'nsnManagerAccountName']
            base_dn = 'o=NSN'
            conn = app.config['ldap']
            # Get the assignee name from the inchargegroup mapping table.
            # try:
            #     result = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, filter, attrs)
            # except:
            #     flash('Link Broken with LDAP Server, Need to Re-Login !')
            #     return redirect(url_for('logout'))
            #
            # JiraIssueReporter = result[0][1]['displayName'][0]
            PRID = request.form['PRID'].strip()
            PRTitle = request.form['PRTitle'].strip()
            PRRcaEdaAssessor = JiraIssueAssignee
            PRRelease = request.form['PRRelease'].strip()
            PRAttached = request.form['PRAttached'].strip()
            PRSeverity = request.form['Severity'].strip()
            PRGroupInCharge = request.form['GroupInCharge'].strip()
            PRProduct = request.form['product'].strip()
            ReportedBy = a['reportedBy']
            FaultCoordinator = a['devFaultCoord']
            CustomerName = a['customerName']
            prontolink = "https://pronto.inside.nsn.com/pronto/problemReportSearch.html?freeTextdropDownID=prId&searchTopText={0}".format(PRID)
            # dict2 = {'customfield_37015': prontolink}
            # issue.update(dict2)
            analysislink = str(prontolink)
            # dict1 = {'customfield_37064': analysislink}
            # # issue.fields.customfield_37064 = analysislink # Invalid, verified.
            # issue.update(dict1)
            # ProntoAttached = ''
            link = upload_sharepoint(PRID)
            prontolink="https://pronto.inside.nsn.com/pronto/problemReportSearch.html?freeTextdropDownID=prId&searchTopText={0}".format(PRID)
            # dict2 = {'customfield_37015': prontolink}
            # issue.update(dict2)
            analysislink = str(link)
            # dict1 = {'customfield_37064': analysislink}
            Prontonumber = PRID + ',' + PRAttached
            summaryinfo = PRTitle #'5WhyRca Parent Task for ' + PRID
            issueaddforRCAsubtask = {
                'project': {'id': u'41675'},
                'issuetype': {'id': u'17678'},
                'summary': summaryinfo,
                'customfield_37060': Prontonumber,
                'customfield_37069': {'id': JiraIssueSourceList},  # Source List u'206960'
                'customfield_37460': {'id': JiraIssueBusinessUnit},  # Business Unit{}
                'customfield_37061': [{'id':JiraIssueBusinessLine}],  # 'Business Linef{}',
                'customfield_37062': {'id': JiraIssueProductLine},  # {'Product Line'},
                'customfield_37063': JiraIssueFeatureComponent,  # 'Feature Component'
                'customfield_37015': prontolink,
                'priority': {'id': JiraIssuePriority},
                'labels': JiraIssueLabels, #['AAA', 'BBB'],
                'customfield_37059': PRRelease,  # 'Release Fault Introduced', #softwareRelease
                'customfield_32006': [{'id':JiraIssueCustomerName}],  # {'Customer':}, # less[
                'customfield_38068': JiraIssueOther,  # {Other} #More filled CMCC

                'customfield_37064': analysislink,
                'assignee': {'name':JiraIssueAssignee } #Optional
            }
            jira = getjira()
            issues = jira.search_issues(jql_str='project = MNPRCA AND summary~' + PRID + ' AND type = "5WhyRCA" ',maxResults=5)
            existFlag = True # existFlag not reset ,represent the issue not exist.
            for newissue in issues:
                newissue.update(issueaddforRCAsubtask)
                existFlag = False
            if existFlag: # Flag be reset to False, no need create new issue at all
                newissue = jira.create_issue(issueaddforRCAsubtask)
            JiraIssueId = str(newissue.key)
            JiraIssueStatus = str(newissue.fields.status)
            JiraIssueReporter = str(newissue.fields.reporter)

            JiraIssuePriority = PriorityDict[JiraIssuePriority]  # Parent task has, Child task does not have
            JiraIssueSourceList = SourceListDict[JiraIssueSourceList]  # Parent task has, Child task does not have
            JiraIssueBusinessUnit = BusinessUnitDict[JiraIssueBusinessUnit]  # Parent task has, Child task does not have
            JiraIssueBusinessLine = BusinessLineDict[JiraIssueBusinessLine]  # Parent task has, Child task does not have
            JiraIssueProductLine = ProductLineDict[JiraIssueProductLine]  # Parent task has, Child task does not have
            JiraIssueCustomerName = CustomerNameDict[JiraIssueCustomerName]  # Parent task has, Child task does not have
            JiraIssueSummary = str(newissue.fields.summary)
            CustomerName = CustomerName
            todo_item = Todo.query.get(JiraIssueId)
            if todo_item:  #
                todo_item.JiraIssueFeature = JiraIssueFeature
                todo_item.JiraIssueFeatureComponent = JiraIssueFeatureComponent
                todo_item.JiraIssueOther = JiraIssueOther
                todo_item.JiraIssueType = JiraIssueType
                todo_item.JiraIssueCaseType = JiraIssueCaseType
                todo_item.JiraIssueLabels = JiraIssueLabels
                todo_item.JiraIssueAssignee = JiraIssueAssignee

                todo_item.PRID = PRID
                todo_item.PRTitle = PRTitle
                todo_item.PRRcaEdaAssessor = PRRcaEdaAssessor
                todo_item.PRRelease = PRRelease
                todo_item.PRAttached = PRAttached
                todo_item.PRSeverity = PRSeverity
                todo_item.PRGroupInCharge = PRGroupInCharge
                todo_item.PRProduct = PRProduct
                todo_item.ReportedBy = ReportedBy
                todo_item.FaultCoordinator = FaultCoordinator

                todo_item.JiraIssueId = JiraIssueId
                todo_item.JiraIssueStatus = JiraIssueStatus
                todo_item.JiraIssueReporter = JiraIssueReporter

                todo_item.JiraIssuePriority = JiraIssuePriority
                todo_item.JiraIssueSourceList = JiraIssueSourceList
                todo_item.JiraIssueBusinessUnit = JiraIssueBusinessUnit
                todo_item.JiraIssueBusinessLine = JiraIssueBusinessLine
                todo_item.JiraIssueProductLine = JiraIssueProductLine
                todo_item.JiraIssueCustomerName = JiraIssueCustomerName

                db.session.commit()
            else:
                todo = Todo(JiraIssueId,JiraIssuePriority,JiraIssueSourceList,JiraIssueBusinessUnit,IsLongCycleTime, \
                 JiraIssueBusinessLine,JiraIssueProductLine,JiraIssueCustomerName,JiraIssueFeature,JiraIssueFeatureComponent,\
                 JiraIssueOther,JiraIssueType,JiraIssueCaseType,JiraIssueStatus,JiraIssueLabels,JiraIssueAssignee,\
                 JiraIssueReporter,PRID,PRTitle,PRRelease,PRSeverity,PRGroupInCharge,PRAttached,PRRcaEdaAssessor,\
                 PRProduct,ReportedBy,FaultCoordinator,JiraIssueSummary,CustomerName)
                todo.user = g.user
                db.session.add(todo)
                db.session.commit()
                flash('RCA item was successfully created')
            return redirect(url_for('index'))
    return render_template('new.html',user = User.query.get(g.user.id).displayName)

@app.route('/index_by_assignee',methods=['GET','POST'])
@login_required
def index_by_assignee():
    if request.method=='GET':
        aa = g.user.username
        bb = gSSOPWD
        options = {'server': 'https://jiradc.int.net.nokia.com/'}
        jira = JIRA(options, basic_auth=(aa, bb))
        assignee = g.user.username
        try:
            issues = jira.search_issues('project = MNRCA and assignee = ' + assignee + '')
        except:
            issues =[]
        for issue in issues:
            JiraIssueId = str(issue.key)
            todo_item = Todo.query.get(JiraIssueId)
            if todo_item:  # PR in the rca table,just do the assignment and update assignee.
                # PRID = todo_item.PRID.strip()
                # if PRID in ('5WhyRca Parent Task for CAS-154355-C888-Test','PR1234'):
                #     continue
                # url = "https://pronto.inside.nsn.com/prontoapi/rest/api/latest/problemReport/%s" % PRID
                # try:
                #     r = requests.get(url, verify=False, auth=(g.user.username, gSSOPWD))
                # except:
                #     flash(' %s -- Invalid PRID has been input!' % PRID, 'error')
                #     print PRID
                # a = r.json()
                todo_item.JiraIssueStatus = str(issue.fields.status)
                # todo_item.JiraIssueSummary = issue.fields.summary
                # todo_item.CustomerName = a['customerName']
                db.session.commit()
            else:
                JiraIssueType = str(issue.fields.issuetype)
                if JiraIssueType == '5WhyRCA': #Parent Task
                    JiraIssuePriority = str(issue.fields.priority)
                    #PRAttached = issue.fields.customfield_37060 #PR number
                    JiraIssueSourceList = issue.fields.customfield_37069  # Parent task has, Child task does not have
                    # JiraIssueLabels =  str(issue.fields.lablels)
                    JiraIssueBusinessUnit = str(issue.fields.customfield_37460)
                    JiraIssueBusinessLine = str(issue.fields.customfield_37061)
                    JiraIssueProductLine = str(issue.fields.customfield_37062) # Parent task has, Child task does not have
                    JiraIssueCustomerName = str(issue.fields.customfield_32006)  # Parent task has, Child task does not have
                    #JiraIssueFeature = issue.fields.customfield_32006
                    JiraIssueFeatureComponent = str(issue.fields.customfield_37063)
                    JiraIssueOther = str(issue.fields.customfield_38068) # Parent task has, Child task does not have
                    JiraIssueCaseType = ''
                    PRTitle = issue.fields.summary  # task title

                    # childissues = issue.fields.subtasks
                    # parentissue = issue.fields.parent

                elif JiraIssueType in ('Analysis subtask','Action for RCA','Action for EDA'):
                    parentissue = issue.fields.parent # Get Parent Task
                    PRTitle = parentissue.fields.summary # subtask title
                    JiraIssuePriority = ''
                    JiraIssueSourceList = '' # Parent task has, Child task does not have
                    JiraIssueBusinessUnit = ''
                    JiraIssueBusinessLine = ''
                    JiraIssueProductLine = '' # Parent task has, Child task does not have
                    JiraIssueCustomerName = ''  # Parent task has, Child task does not have
                    JiraIssueFeatureComponent = ''
                    JiraIssueOther = '' # Parent task has, Child task does not have
                    JiraIssueCaseType = ''
                    if JiraIssueType in ('Analysis subtask',):
                        JiraIssueCaseType = str(issue.fields.customfield_10464) # Subtask Only properties

                JiraIssueStatus = str(issue.fields.status)
                JiraIssueAssignee = str(issue.fields.assignee)
                JiraIssueReporter = str(issue.fields.reporter)
                PRRcaEdaAssessor = JiraIssueAssignee
                JiraIssueLabels = issue.fields.labels  # Common for all
                # if not JiraIssueLabels:
                #     JiraIssueLabels = ''
                if not JiraIssueLabels:
                    JiraIssueLabels = ''
                else:
                    LabelItems = JiraIssueLabels
                    a = ''
                    for s in LabelItems:
                        a = a + s + ','
                    JiraIssueLabels = a.strip(',')
                PRID = PRTitle.split(':')[0]
                print '%s'%PRID
                if PRID in ('5WhyRca Parent Task for CAS-154355-C888-Test','PR1234'):
                    continue
                url = "https://pronto.inside.nsn.com/prontoapi/rest/api/latest/problemReport/%s" % PRID
                try:
                    r = requests.get(url, verify=False, auth=(g.user.username, gSSOPWD))
                except:
                    flash(' %s -- Invalid PRID has been input!' % PRID, 'error')
                    print PRID
                a = r.json()

                PRSeverity = a['severity'].strip()
                PRRelease = a['softwareRelease'].strip()
                PRAttached = a['problemReportIds']
                # if not PRAttached:
                #     PRAttached = ''
                if not PRAttached:
                    PRAttached = ''
                else:
                    LabelItems = PRAttached
                    b = ''
                    for s in LabelItems:
                        b = b + s + ','
                        PRAttached = b.strip(',')
                PRGroupInCharge = a['groupIncharge'].strip()
                PRProduct = a['product'].strip()
                ReportedBy = a['reportedBy']
                FaultCoordinator = a['devFaultCoord']
                JiraIssueFeature = a['feature']
                JiraIssueSummary = issue.fields.summary
                CustomerName = a['customerName']
                todo = Todo(JiraIssueId, JiraIssuePriority, JiraIssueSourceList, JiraIssueBusinessUnit, \
                            JiraIssueBusinessLine, JiraIssueProductLine, JiraIssueCustomerName,
                            JiraIssueFeature, JiraIssueFeatureComponent, \
                            JiraIssueOther, JiraIssueType, JiraIssueCaseType, JiraIssueStatus,
                            JiraIssueLabels, JiraIssueAssignee, \
                            JiraIssueReporter, PRID, PRTitle, PRRelease, PRSeverity, PRGroupInCharge,
                            PRAttached, PRRcaEdaAssessor, \
                            PRProduct, ReportedBy, FaultCoordinator,JiraIssueSummary,CustomerName)
                todo.user = g.user
                db.session.add(todo)
                db.session.commit()
                flash('RCA item was successfully created')
        hello = User.query.filter_by(username=g.user.username).first()
        if hello:
            normaluser = hello.displayName
            JiraIssueAssignee = normaluser
            count = Todo.query.filter_by(JiraIssueAssignee=JiraIssueAssignee).count()
            todos = Todo.query.filter_by(JiraIssueAssignee=JiraIssueAssignee).all()
            headermessage = 'All MNRCA JIRA RCA/EDA Items Assigned To Me'
            return render_template('index.html', count=count, headermessage=headermessage,
                                   todos=todos, user=User.query.get(g.user.id).displayName)

@app.route('/index_by_reporter',methods=['GET','POST'])
@login_required
def index_by_reporter():
    if request.method=='GET':
        aa = g.user.username
        bb = gSSOPWD
        options = {'server': 'https://jiradc.int.net.nokia.com/'}
        jira = JIRA(options, basic_auth=(aa, bb))
        assignee = g.user.username
        try:
            issues = jira.search_issues('project = MNRCA and reporter = ' + assignee + '')
        except:
            issues =[]
        for issue in issues:
            JiraIssueId = str(issue.key)
            todo_item = Todo.query.get(JiraIssueId)
            if todo_item:  # PR in the rca table,just do the assignment and update assignee.
                # PRID = todo_item.PRID.strip()
                # if PRID in ('5WhyRca Parent Task for CAS-154355-C888-Test','PR1234'):
                #     continue
                # url = "https://pronto.inside.nsn.com/prontoapi/rest/api/latest/problemReport/%s" % PRID
                # try:
                #     r = requests.get(url, verify=False, auth=(g.user.username, gSSOPWD))
                # except:
                #     flash(' %s -- Invalid PRID has been input!' % PRID, 'error')
                #     print PRID
                # a = r.json()
                todo_item.JiraIssueStatus = str(issue.fields.status)
                # todo_item.JiraIssueSummary = issue.fields.summary
                # todo_item.CustomerName = a['customerName']
                db.session.commit()
            else:
                JiraIssueType = str(issue.fields.issuetype)
                if JiraIssueType == '5WhyRCA': #Parent Task
                    JiraIssuePriority = str(issue.fields.priority)
                    #PRAttached = issue.fields.customfield_37060 #PR number
                    JiraIssueSourceList = issue.fields.customfield_37069  # Parent task has, Child task does not have
                    # JiraIssueLabels =  str(issue.fields.lablels)
                    JiraIssueBusinessUnit = str(issue.fields.customfield_37460)
                    JiraIssueBusinessLine = str(issue.fields.customfield_37061)
                    JiraIssueProductLine = str(issue.fields.customfield_37062) # Parent task has, Child task does not have
                    JiraIssueCustomerName = str(issue.fields.customfield_32006)  # Parent task has, Child task does not have
                    #JiraIssueFeature = issue.fields.customfield_32006
                    JiraIssueFeatureComponent = str(issue.fields.customfield_37063)
                    JiraIssueOther = str(issue.fields.customfield_38068) # Parent task has, Child task does not have
                    JiraIssueCaseType = ''
                    PRTitle = issue.fields.summary  # task title

                    # childissues = issue.fields.subtasks
                    # parentissue = issue.fields.parent

                elif JiraIssueType in ('Analysis subtask','Action for RCA','Action for EDA'):
                    parentissue = issue.fields.parent # Get Parent Task
                    PRTitle = parentissue.fields.summary # subtask title
                    JiraIssuePriority = ''
                    JiraIssueSourceList = '' # Parent task has, Child task does not have
                    JiraIssueBusinessUnit = ''
                    JiraIssueBusinessLine = ''
                    JiraIssueProductLine = '' # Parent task has, Child task does not have
                    JiraIssueCustomerName = ''  # Parent task has, Child task does not have
                    JiraIssueFeatureComponent = ''
                    JiraIssueOther = '' # Parent task has, Child task does not have
                    JiraIssueCaseType = ''
                    if JiraIssueType in ('Analysis subtask',):
                        JiraIssueCaseType = str(issue.fields.customfield_10464) # Subtask Only properties

                JiraIssueStatus = str(issue.fields.status)
                JiraIssueAssignee = str(issue.fields.assignee)
                JiraIssueReporter = str(issue.fields.reporter)
                PRRcaEdaAssessor = JiraIssueAssignee
                JiraIssueLabels = issue.fields.labels  # Common for all
                if not JiraIssueLabels:
                    JiraIssueLabels = ''
                else:
                    LabelItems = JiraIssueLabels
                    a = ''
                    for s in LabelItems:
                        a = a + s + ','
                    JiraIssueLabels = a.strip(',')
                PRID = PRTitle.split(':')[0]
                print '%s'%PRID
                if PRID in ('5WhyRca Parent Task for CAS-154355-C888-Test','PR1234','Light Weight RCA','5WhyRca Parent Task for PR389995'):
                    continue
                url = "https://pronto.inside.nsn.com/prontoapi/rest/api/latest/problemReport/%s" % PRID

                r = requests.get(url, verify=False, auth=(g.user.username, gSSOPWD))
                if r.ok == False:
                    flash(' %s -- Invalid PRID has been input!' % PRID, 'error')
                    print PRID
                a = r.json()

                PRSeverity = a['severity'].strip()
                PRRelease = a['softwareRelease'].strip()
                PRAttached = a['problemReportIds']
                # if not PRAttached:
                #     PRAttached = ''
                if not PRAttached:
                    PRAttached = ''
                else:
                    LabelItems = PRAttached
                    b = ''
                    for s in LabelItems:
                        b = b + s + ','
                        PRAttached = b.strip(',')
                PRGroupInCharge = a['groupIncharge'].strip()
                PRProduct = a['product'].strip()
                ReportedBy = a['reportedBy']
                FaultCoordinator = a['devFaultCoord']
                JiraIssueFeature = a['feature']
                JiraIssueSummary = issue.fields.summary
                CustomerName = 'Internal'
                if ReportedBy == 'Customer':
                    CustomerName = a['customerName']
                todo = Todo(JiraIssueId, JiraIssuePriority, JiraIssueSourceList, JiraIssueBusinessUnit, \
                            JiraIssueBusinessLine, JiraIssueProductLine, JiraIssueCustomerName,
                            JiraIssueFeature, JiraIssueFeatureComponent, \
                            JiraIssueOther, JiraIssueType, JiraIssueCaseType, JiraIssueStatus,
                            JiraIssueLabels, JiraIssueAssignee, \
                            JiraIssueReporter, PRID, PRTitle, PRRelease, PRSeverity, PRGroupInCharge,
                            PRAttached, PRRcaEdaAssessor, \
                            PRProduct, ReportedBy, FaultCoordinator,JiraIssueSummary,CustomerName)
                todo.user = g.user
                db.session.add(todo)
                db.session.commit()
                flash('RCA item was successfully created')


        hello = User.query.filter_by(username=g.user.username).first()
        if hello:
            normaluser = hello.displayName
            JiraIssueAssignee = normaluser
            count = Todo.query.filter_by(JiraIssueReporter=JiraIssueReporter).count()
            todos = Todo.query.filter_by(JiraIssueReporter=JiraIssueReporter).all()
            headermessage = 'All MNRCA JIRA RCA/EDA Reported By Me'
            return render_template('index.html', count=count, headermessage=headermessage,
                                   todos=todos, user=User.query.get(g.user.id).displayName)
admin=['leienqing',]
@app.route('/rca_home',methods=['GET','POST'])
@login_required
def rca_home():
	hello = User.query.filter_by(username=g.user.username).first()
	if hello:
		normaluser = hello.email
		count = Todo.query.filter_by(user_id=g.user.id).count()
		todos = Todo.query.filter_by(user_id=g.user.id).all()
		headermessage = 'All MNRCA Items'
		return render_template('index.html', count=count, headermessage=headermessage,todos= todos,\
							   user=User.query.get(g.user.id).displayName)


class SystemLog(db.Model):
    __tablename__ = 'jirasyslog'
    id = db.Column('Index', db.Integer, primary_key=True)
    loginAccount=db.Column(db.String(64),default='')
    loginLocation = db.Column(db.String(128), default='')
    ipAddress = db.Column(db.String(128),default='')
    browserType = db.Column(db.String(128),default='')
    deviceType = db.Column(db.String(128),default='')
    osType = db.Column(db.String(128),default='')
    operationType = db.Column(db.String(128),default='')
    logTime = db.Column(db.DateTime)
    prIdorApId = db.Column(db.String(128),default='')
    log1 = db.Column(db.String(1024),default='')
    log2 = db.Column(db.String(1024),default='')
    log3 = db.Column(db.String(1024),default='')
    log4 = db.Column(db.String(1024),default='')
    log5 = db.Column(db.String(1024),default='')
    log6 = db.Column(db.String(1024),default='')
    log7 = db.Column(db.String(1024),default='')
    log8 = db.Column(db.String(1024),default='')
    log9 = db.Column(db.String(1024),default='')
    log10 = db.Column(db.String(1024),default='')
    log11 = db.Column(db.String(1024),default='')
    log12 = db.Column(db.String(1024),default='')
    def __init__(self, loginAccount,loginLocation,ipAddress,browserType,deviceType,osType,operationType,prIdorApId, \
                 log1,log2,log3,log4,log5,log6,log7,log8,log9,log10,log11,log12):
        self.loginAccount = loginAccount
        self.loginLocation = loginLocation
        self.ipAddress = ipAddress
        self.browserType = browserType
        self.deviceType = deviceType
        self.osType = osType
        self.operationType = operationType
        self.logTime = datetime.utcnow()
        self.prIdorApId = prIdorApId
        self.log1 = log1
        self.log2 = log2
        self.log3 = log3
        self.log4 = log4
        self.log5 = log5
        self.log6 = log6
        self.log7 = log7
        self.log8 = log8
        self.log9 = log9
        self.log10 = log10
        self.log11 = log11
        self.log12 = log12
        
db.create_all()

def syslog_record(loginAccount,operationType,prIdorApId,log1,log2,log3,log4,log5,log6,log7,log8,log9,log10,log11,log12):
    log_dict = get_login_info()
    """
    log_dict['IP']= remote_ip
    log_dict['BR']= browser
    log_dict['DEV']= device
    log_dict['OS']= os
    log_dict['LOC']= location
    """
    loginAccount = loginAccount
    loginLocation = log_dict['LOC']
    ipAddress = log_dict['IP']
    browserType = log_dict['BR']
    deviceType = log_dict['DEV']
    osType = log_dict['OS']
    operationType = operationType
    loginfo = SystemLog(loginAccount,loginLocation,ipAddress,browserType,deviceType,osType,operationType,prIdorApId,\
                 log1,log2,log3,log4,log5,log6,log7,log8,log9,log10,log11,log12)
    db.session.add(loginfo)
    db.session.commit()

app.config['dbconfig'] = {'host': '10.68.184.123',
                          'port':8080,
                          'user': 'root',
                          'password': 'jupiter111',
                          'database': 'jupiter4', }

app.config['dbconfig1'] = {'host': '127.0.0.1',
                          'user': 'root',
                          'password': '',
                          'database': 'fddrca', }

class UseDatabase:
    def __init__(self, config):
        self.configuration = config
    def __enter__(self):
        self.conn = mysql.connector.connect(**self.configuration)
        self.cursor = self.conn.cursor()
        return self.cursor
    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.conn.commit()
        self.cursor.close()
        self.conn.close()

class UseDatabaseDict:
    def __init__(self, config):
        """Add the database configuration parameters to the object.
        This class expects a single dictionary argument which needs to assign
        the appropriate values to (at least) the following keys:
            host - the IP address of the host running MySQL/MariaDB.
            user - the MySQL/MariaDB username to use.
            password - the user's password.
            database - the name of the database to use.
        For more options, refer to the mysql-connector-python documentation.
        """
        self.configuration = config
    def __enter__(self):
        """
        Connect to database and create a DB cursor.
            cursor = cnx.cursor(dictionary=True)
            cursor.execute("SELECT * FROM country WHERE Continent = 'Europe'")
        Return the database cursor to the context manager.
        """
        self.conn = mysql.connector.connect(**self.configuration)
        self.cursor = self.conn.cursor(dictionary=True)
        return self.cursor

    def __exit__(self, exc_type, exc_value, exc_traceback):
        """Destroy the cursor as well as the connection (after committing).
        """
        self.conn.commit()
        self.cursor.close()
        self.conn.close()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS


@app.route('/api/upload', methods=['POST'], strict_slashes=False)
def api_upload():
    file_dir = os.path.join(basedir, app.config['UPLOAD_FOLDER'])
    file_dir = os.path.join(file_dir, g.user.username)
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
    f = request.files['fileField']
    print f
    print 'fakepath**************'
    if f and allowed_file(f.filename):
        fname = f.filename
        ext = fname.rsplit('.', 1)[1]
        unix_time = int(time.time())
        new_filename = str(unix_time) + '.' + ext
        filename = os.path.join(file_dir, new_filename)
        print filename
        f.save(os.path.join(file_dir, new_filename))
        val = importfromexcel(filename)
        if val == 0:
            flash('PRID has been there !', 'error')
            todo = Todo.query.order_by(Todo.PRID.desc()).first()
            a = re.sub("\D", "", todo.PRID)
            # a=filter(str.isdigit, todo.PRID)
            a = int(a)
            a = a + 1

            b = len(str(a))
            # sr=sr+'m'*(9-len(sr))
            PRID = 'MAC' + '0' * (6 - b) + str(a)
            return render_template('new.html'
                                   , PRID=PRID)
        flash('Pronto item has been successfully imported')
        return redirect(url_for('index'))
    else:
        flash('Invalid Filename!')
        return redirect(url_for('index'))

def modifyColumnType(fieldname):
    with UseDatabase(app.config['dbconfig']) as cursor:
        #alter table user MODIFY new1 VARCHAR(1) -->modify field type
        _SQL = "alter table rcastatus MODIFY `"+fieldname+"` VARCHAR(512)"
        cursor.execute(_SQL)

def addRCAColumn5():
    with UseDatabase(app.config['dbconfig']) as cursor:
        #alter table user MODIFY new1 VARCHAR(1) -->modify field type
        _SQL = "alter table rcastatus ADD column (LteCategory VARCHAR(32),CustomerOrInternal VARCHAR(32),JiraFunctionArea VARCHAR(32),\
        TriggerScenarioCategory VARCHAR(32),FirstFaultEcapePhase VARCHAR(32),FaultIntroducedRelease VARCHAR(32),TechnicalRootCause VARCHAR(1024),\
        TeamAssessor VARCHAR(64),EdaCause VARCHAR(1024),RcaRootCause5WhyAnalysis VARCHAR(2048))"
        cursor.execute(_SQL)

def addApColumn5():
    with UseDatabase(app.config['dbconfig']) as cursor:
        #alter table user MODIFY new1 VARCHAR(1) -->modify field type
        _SQL = "alter table apstatus ADD column (InjectionRootCauseEdaCause VARCHAR(1024),RcaEdaCauseType VARCHAR(32),\
        RcaEdaActionType VARCHAR(32),TargetRelease VARCHAR(128))"
        cursor.execute(_SQL)

def addColumn():
    with UseDatabase(app.config['dbconfig']) as cursor:
        #alter table user MODIFY new1 VARCHAR(1) -->modify field type
        _SQL = "alter table rcastatus modify column JiraIssueAssignee VARCHAR(128)"
        cursor.execute(_SQL)

def compare_time(start_t,end_t):
    s_time = time.mktime(time.strptime(start_t,'%Y-%m-%d'))                        
    #get the seconds for specify date
    e_time = time.mktime(time.strptime(end_t,'%Y-%m-%d'))
    if float(s_time) >= float(e_time):
        return True
    return False 

def comparetime(start_t,end_t):
    s_time = time.mktime(time.strptime(start_t,'%Y-%m-%d'))                        
    #get the seconds for specify date
    e_time = time.mktime(time.strptime(end_t,'%Y-%m-%d'))
    if(float(e_time)- float(s_time)) > float(86400):
        print ("@@@float(e_time)- float(s_time))=%f"%(float(e_time)- float(s_time)))
        return True
    return False 

def leap_year(y):
    if (y % 4 == 0 and y % 100 != 0) or y % 400 == 0:
        return True
    else:
        return False
        
def days_in_month(y, m): 
    if m in [1, 3, 5, 7, 8, 10, 12]:
        return 31
    elif m in [4, 6, 9, 11]:
        return 30
    else:
        if leap_year(y):
            return 29
        else:
            return 28
            
def days_this_year(year): 
    if leap_year(year):
        return 366
    else:
        return 365
            
def days_passed(year, month, day):
    m = 1
    days = 0
    while m < month:
        days += days_in_month(year, m)
        m += 1
    return days + day

def dateIsBefore(year1, month1, day1, year2, month2, day2):
    """Returns True if year1-month1-day1 is before year2-month2-day2. Otherwise, returns False."""
    if year1 < year2:
        return True
    if year1 == year2:
        if month1 < month2:
            return True
        if month1 == month2:
            return day1 < day2
    return False

def daysBetweenDates(year1, month1, day1, year2, month2, day2):
    if year1 == year2:
        return days_passed(year2, month2, day2) - days_passed(year1, month1, day1)
    else:
        sum1 = 0
        y1 = year1
        while y1 < year2:
            sum1 += days_this_year(y1)
            y1 += 1
        return sum1-days_passed(year1,month1,day1)+days_passed(year2,month2,day2)

"""
    ip_set = [int(i) for i in ip_addr.split('.')]
    ip_number = (ip_set[0] << 24) + (ip_set[1] << 16) + (ip_set[2] << 8) + ip_set[3]
    return ip_number
    ext = fname.rsplit('.', 1)[1]
    
"""
def daysBetweenDate(start,end):
    year1=int(start.split('-',2)[0])
    month1=int(start.split('-',2)[1])
    day1=int(start.split('-',2)[2])

    year2=int(end.split('-',2)[0])
    month2=int(end.split('-',2)[1])
    day2=int(end.split('-',2)[2])
    print ("daysBetweenDates(year1, month1, day1, year2, month2, day2)=%d"%daysBetweenDates(year1, month1, day1, year2, month2, day2))
    return daysBetweenDates(year1, month1, day1, year2, month2, day2)

def insert_item(team,internaltask_sheet,i):
    PRID = internaltask_sheet.cell_value(i+1,2)
    print PRID
    PRTitle = internaltask_sheet.cell_value(i+1,9)
    PRReportedDate = internaltask_sheet.cell_value(i+1,5)    
    PRClosedDate =internaltask_sheet.cell_value(i+1,44)
    PROpenDays=daysBetweenDate(PRReportedDate,PRClosedDate)
    PRRcaCompleteDate =''

    PRRelease = internaltask_sheet.cell_value(i+1,6)
    PRAttached = internaltask_sheet.cell_value(i+1,27)
    
    if daysBetweenDate(PRReportedDate,PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'
    
    #IsLongCycleTime= internaltask_sheet.cell_value(i+1,34)
    IsCatM =''
    IsRcaCompleted='No'
    LongCycleTimeRcaIsCompleted='No'
    NoNeedDoRCAReason =''
    RootCauseCategory = ''
    FunctionArea = ''

    CodeDeficiencyDescription = ''
    CorrectionDescription = ''
    RootCause=''
    LongCycleTimeRootCause=''

    IntroducedBy=''
    Handler = team
    todo_item = Todo.query.get(PRID)
    registered_user = Todo.query.filter_by(PRID=PRID).all()
    if len(registered_user) ==0:
	print 'OK#################################################'
	todo = Todo(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,PRRelease,PRAttached,IsLongCycleTime,\
                 IsCatM,IsRcaCompleted,NoNeedDoRCAReason,RootCauseCategory,FunctionArea,CodeDeficiencyDescription,\
		 CorrectionDescription,RootCause,IntroducedBy,Handler)
        #g.user=Todo.query.get(team)
	todo.user = g.user
	print("todo.user = g.user=%s"%todo.user)
	hello = User.query.filter_by(username=team).first()
	todo.user_id=hello.id
	print("todo.user_id=hello.user_id=%s"%hello.id)
	db.session.add(todo)
	db.session.commit()
    else:
        print ("registered_user.PRTitle=%s"%PRID)

    registered_user = TodoLongCycleTimeRCA.query.filter_by(PRID=PRID).all()
    if len(registered_user) ==0 and IsLongCycleTime=='Yes':
	print 'OK#################################################'
	todo = TodoLongCycleTimeRCA(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,IsLongCycleTime,\
                 IsCatM,LongCycleTimeRcaIsCompleted,LongCycleTimeRootCause,NoNeedDoRCAReason,Handler)
        #g.user=Todo.query.get(team)
	todo.user = g.user
	print("todo.user = g.user=%s"%todo.user)
	hello = User.query.filter_by(username=team).first()
	todo.user_id=hello.id
	print("todo.user_id=hello.user_id=%s"%hello.id)
	db.session.add(todo)
	db.session.commit()
    else:
        print ("registered_user.PRTitle=%s"%PRID)
    
def importfromexcel(filename):
    workbook = xlrd.open_workbook(filename)
    internaltask_sheet=workbook.sheet_by_name(r'PR List')
    rows=internaltask_sheet.row_values(0)
    nrows=internaltask_sheet.nrows
    ncols=internaltask_sheet.ncols
    #modifyColumnType('PRTitle')
    print str(nrows)+"*********"
    print ncols
    k=0
    #teams=['chenlong','xiezhen','yangjinyong','zhangyijie','lanshenghai','liumingjing','lizhongyuan','caizhichao','hujun','wangli']
    for i in range(nrows-1):
        ClosedEnter= internaltask_sheet.cell_value(i+1,44)
        PRState= internaltask_sheet.cell_value(i+1,8)
        team= internaltask_sheet.cell_value(i+1,63)
        print ClosedEnter
        if PRState=='Closed':
            time='2018-1-1'
            if compare_time (ClosedEnter,time):
                if team in teams:
                    print team
                    k=k+1
                    teamname=teams[teams.index(team)]
                    print teamname
                    insert_item(team,internaltask_sheet,i)
      
    val=1
    print k
    print 'closed pr'
    return val
    
def get_apid():
    todoap=TodoAP.query.order_by(TodoAP.APID.desc()).first()
    if todoap is not None:
        a = re.sub("\D", "", todoap.APID)
        a=int(a)
        a=a+1 
        b=len(str(a))
    else:
        a=1
        b=1

    APID='AP'+'0'*(6-b)+ str(a)
    return APID
def update_rca(PRID,internaltask_sheet):
    todo_item = Todo.query.get(PRID)
    if todo_item is None:
        flash('Please check the PR, seems it is not in the Formal RCA PR list!', 'error')
        return False
    todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))    
    todo_item.IsRcaCompleted = 'Yes'
    aa=findRootCauseIndex(16,'Root Cause',internaltask_sheet)
    todo_item.RootCause = internaltask_sheet.cell_value(aa,11)
    #todo_item.RootCause  = internaltask_sheet.cell_value(20,11)
    s2 ='Triggering scenario category: select appropriate item from the list'
    s1 = internaltask_sheet.row_values(6)[1]
    if s1.strip() == s2 :
        todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(7)[3]
        todo_item.CorrectionDescription  = internaltask_sheet.row_values(8)[3]
    else:
        todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(6)[3]
        todo_item.CorrectionDescription  = internaltask_sheet.row_values(7)[3]
    """
    with UseDatabase(app.config['dbconfig']) as cursor:
        _SQL = "alter table rcastatus MODIFY column FunctionArea VARCHAR(1024)"
        cursor.execute(_SQL)

    todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(6)[3]
    todo_item.CorrectionDescription  = internaltask_sheet.row_values(7)[3]
	"""
    db.session.commit()
    return True

def update_longcycletimerca(PRID,internaltask_sheet):
    todo_item = TodoLongCycleTimeRCA.query.get(PRID)
    if todo_item is None:
        flash('No in LongCycleTimeRca PR list!', 'error')
        return False
    todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))    
    todo_item.LongCycleTimeRcaIsCompleted = 'Yes'
    bb=findLongCycleRootCauseIndex(16,'Root Cause',internaltask_sheet)
    todo_item.LongCycleTimeRootCause  = internaltask_sheet.cell_value(bb,11)

    db.session.commit()
    return True

def insert_ap(internaltask_sheet,index,ApCategory):
    PRID = internaltask_sheet.cell_value(2,1).strip()
    if len(PRID)>15:
        prlist  = PRID.split('\n')
        for PRID in prlist:
            count = Todo.query.filter_by(PRID=PRID).count()
            if count:
                PRID = PRID
                break
    APCompletedOn=''
    IsApCompleted='No'
    APDescription = internaltask_sheet.cell_value(index,12)
    if len(APDescription.strip())!=0:
        APCreatedDate=time.strftime('%Y-%m-%d',time.localtime(time.time()))
        APAssingnedTo=internaltask_sheet.cell_value(index,15)
        #ctype:0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
        ctype= internaltask_sheet.cell(index,17).ctype
        APDueDate= internaltask_sheet.cell_value(index,17)
        # New added for JIRA
        InjectionRootCauseEdaCause = internaltask_sheet.cell_value(index, 11)  # RCA/EDA RootCause
        RcaEdaCauseType = internaltask_sheet.cell_value(index, 13).strip()
        RcaEdaActionType = internaltask_sheet.cell_value(index, 14)
        TargetRelease ="" #internaltask_sheet.row_values(12)[3]
        CustomerAp = "No"
        ApCategory = ApCategory  # RCA/EDA
        ShouldHaveBeenDetected=''
        if ApCategory =='EDA':
            s2 = 'Triggering scenario category: select appropriate item from the list'
            s1 = internaltask_sheet.row_values(6)[1]
            if s1.strip() == s2:
                ShouldHaveBeenDetected = internaltask_sheet.cell_value(index, 18)
        ApJiraId = ""
        if ctype==3:
            date = datetime(*xldate_as_tuple(APDueDate, 0))
            APDueDate = date.strftime('%Y-%m-%d')
        else:
            APDueDate = APDueDate
        #APDueDate=APDueDate.strftime('%Y-%d-%m')
        APID=get_apid()
        QualityOwner=User.query.get(g.user.id).username #g.user.id
        todoap = TodoAP(APID,PRID,APDescription,APCreatedDate,APDueDate,APCompletedOn,IsApCompleted,APAssingnedTo,QualityOwner,\
                        InjectionRootCauseEdaCause, RcaEdaCauseType, RcaEdaActionType, TargetRelease, CustomerAp, \
                        ApCategory,ShouldHaveBeenDetected,ApJiraId)
        todoap.user = g.user
        db.session.add(todoap)
        db.session.commit()

def insert_rca5why(internaltask_sheet,index):
    PRID = internaltask_sheet.cell_value(2,1).strip()
    if len(PRID)>15:
        prlist  = PRID.split('\n')
        for PRID in prlist:
            count = Todo.query.filter_by(PRID=PRID).count()
            if count:
                PRID = PRID
                break
    Why1 = internaltask_sheet.cell_value(index,2)
    Why2 = internaltask_sheet.cell_value(index,4)
    Why3= internaltask_sheet.cell_value(index,6)
    Why4 = internaltask_sheet.cell_value(index,8)
    Why5 = internaltask_sheet.cell_value(index,10)
    rca5why = Rca5Why(PRID,Why1,Why2,Why3,Why4,Why5)
    rca5why.pr_id = PRID
    db.session.add(rca5why)
    db.session.commit()

def findIndex(index_start,target_String,internaltask_sheet):
    APDescription = internaltask_sheet.cell_value(index_start,12)
    for i in range(20):
        APDescription = internaltask_sheet.cell_value(index_start+i,12)
        if APDescription==target_String:
            return index_start+i+1
			
def findRootCauseIndex(index_start,target_String,internaltask_sheet):
    APDescription = internaltask_sheet.cell_value(index_start,11)
    for i in range(10):
        APDescription = internaltask_sheet.cell_value(index_start+i,11)
        if APDescription==target_String:
            return index_start+i+1
			
def findLongCycleRootCauseIndex(index_start,target_String,internaltask_sheet):
    APDescription = internaltask_sheet.cell_value(index_start,11)
    for i in range(10):
        APDescription = internaltask_sheet.cell_value(index_start+i,11)
        if APDescription==target_String:
            return index_start+i+2
        
def find5whyIndex(index_start,target_String,internaltask_sheet):
    APDescription = internaltask_sheet.cell_value(index_start,1)
    for i in range(10):
        APDescription = internaltask_sheet.cell_value(index_start+i,1)
        if APDescription==target_String:
            return index_start+i+1
        
def import_ap_fromexcel(filename):
    workbook = xlrd.open_workbook(filename)
    internaltask_sheet=workbook.sheet_by_name(r'RcaEda')
    rows=internaltask_sheet.row_values(0)
    nrows=internaltask_sheet.nrows
    ncols=internaltask_sheet.ncols
    print ("nrows==%d"%nrows)
    print ("#############nrows==%d"%ncols)
    PRID = internaltask_sheet.cell_value(2,1).strip()
    if len(PRID)>15:
        prlist  = PRID.split('\n')
        for PRID in prlist:
            count = Todo.query.filter_by(PRID=PRID).count()
            if count:
                PRID = PRID
                break
    APCompletedOn=''
    IsApCompleted=''
    QualityOwner=User.query.get(g.user.id).username #g.user.id
    todo=TodoAP.query.filter_by(PRID = PRID).order_by(TodoAP.APID.asc()).first()
    state=0
    a=update_rca(PRID,internaltask_sheet)
    if a is False:
        print ("PRID=%s is not in the PR list of RCA candidate"%PRID)
        return False
    a=update_longcycletimerca(PRID,internaltask_sheet)
    if a is False:
        print ("PRID=%s is not in the LongCycleTime PR list"%PRID)
        #return False
    b=find5whyIndex(16,'Root Cause Analysis',internaltask_sheet)
    insert_rca5why(internaltask_sheet,b)
    insert_rca5why(internaltask_sheet,b+1)
    insert_rca5why(internaltask_sheet,b+2)
    print("import_ap_fromexcel.PRID=%s"%PRID)
    prid=str(PRID)
    """
    with UseDatabase(app.config['dbconfig']) as cursor:
        _SQL = "select * from apstatus where PRID=`"+prid+""
        cursor.execute(_SQL)
        contents = cursor.fetchall()
        if contents
        _SQL = "delete from apstatus where PRID=`"+prid+""
        cursor.execute(_SQL)
    """
    todo_item = TodoAP.query.filter_by(PRID = prid).all()
    if len(todo_item)!=0 and g.user.username in admin:
        for item in todo_item:
            db.session.delete(item)
            print("item.APID=%s,item.PRID=%s"%(item.APID,item.PRID)) +"&*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
        db.session.commit()            
        #TodoAP.query.filter(PRID == PRID).delete()
        #db.session.query(TodoAP).filter(PRID == prid).delete()
    else:
        print "PR's APs has already been there,please check with the Admin"

    if state==0:
        ApCategory = 'RCA'
        b=findIndex(16,'Action Proposal',internaltask_sheet)
        for i in range(4):
            index=b+i
            insert_ap(internaltask_sheet,index,ApCategory)
        ApCategory = 'EDA'        
        b=findIndex(30,'Action Proposal',internaltask_sheet)
        for i in range(3):
            index=b+i
            insert_ap(internaltask_sheet,index,ApCategory)
        c=findIndex(b,'Action Proposal',internaltask_sheet)                
        for i in range(3):
            index=c+i
            insert_ap(internaltask_sheet,index,ApCategory)
        d=findIndex(c,'Action Proposal',internaltask_sheet)
        for i in range(3):
            index=d+i
            insert_ap(internaltask_sheet,index,ApCategory)
        e=findIndex(d,'Action Proposal',internaltask_sheet)
        for i in range(3):
            index=e+i
            insert_ap(internaltask_sheet,index,ApCategory)
        f=findIndex(e,'Action Proposal',internaltask_sheet)
        for i in range(3):
            index=f+i
            insert_ap(internaltask_sheet,index,ApCategory)
        gg=findIndex(f,'Action Proposal',internaltask_sheet)
        if gg is None:
            return True
        for i in range(3):
            index=gg+i
            insert_ap(internaltask_sheet,index,ApCategory)
        h=findIndex(gg,'Action Proposal',internaltask_sheet)
        if h is None:
            return True
        for i in range(3):
            index=h+i
            insert_ap(internaltask_sheet,index,ApCategory)
    return True       

class Excel():  
    def export(self):

        output = BytesIO() 

        writer = pd.ExcelWriter(output, engine='xlwt')
        workbook = writer.book

        worksheet= workbook.add_sheet('sheet1',cell_overwrite_ok=True)
        col=0
        row=1
        pattern = xlwt.Pattern() # Create the Pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        #style = xlwt.XFStyle() # Create the Pattern
        font = xlwt.Font() # Create the Font
        font.name = 'Times New Roman'
        font.bold = True
        #font.underline = True
        #font.italic = True
        style = xlwt.XFStyle() # Create the Style
        style.font = font # Apply the Font to the Style
        style.pattern = pattern # Add Pattern to Style
        columns=['PRID','PRTitle','PRReportedDate','PRClosedDate','PROpenDays','IsLongCycleTime','PRRcaCompleteDate','PRRelease','PRAttached','IsCatM','IsRcaCompleted',\
                 'NoNeedDoRCAReason','RootCauseCategory','FunctionArea','CodeDeficiencyDescription','CorrectionDescription','RootCause','LongCycleTimeRootCause','IntroducedBy','Handler']
        for item in columns:
            worksheet.col(col).width = 4333 # 3333 = 1" (one inch).
            worksheet.write(0, col,item,style)
            col+=1
        style = xlwt.XFStyle()
        style.num_format_str = 'M/D/YY' # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
        alignment = xlwt.Alignment() # Create Alignment
        alignment.horz = xlwt.Alignment.HORZ_JUSTIFIED # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED,HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        alignment.vert = xlwt.Alignment.VERT_TOP # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        #style = xlwt.XFStyle() # Create Style
        style.alignment = alignment # Add Alignment to Style
        if g.user.username in admin:
            todos=Todo.query.order_by(Todo.Handler.desc()).all()
        else:
            todos = Todo.query.filter_by(user_id=g.user.id).order_by(Todo.Handler.desc()).all()
        nrows=len(todos)
        print ('nrows==%s' %(nrows))
        for row in range(nrows):
            row1=row+1
            worksheet.write(row1,0,todos[row].PRID)
            worksheet.write(row1,1,todos[row].PRTitle,style)
            worksheet.write(row1,2,todos[row].PRReportedDate)
            worksheet.write(row1,3,todos[row].PRClosedDate)
            worksheet.write(row1,4,todos[row].PROpenDays)            
            worksheet.write(row1,5,todos[row].IsLongCycleTime)
            worksheet.write(row1,6,todos[row].PRRcaCompleteDate)
            worksheet.write(row1,7,todos[row].PRRelease)
            worksheet.write(row1,8,todos[row].PRAttached,style)
            worksheet.write(row1,9,todos[row].IsCatM)
            worksheet.write(row1,10,todos[row].IsRcaCompleted)
            worksheet.write(row1,11,todos[row].NoNeedDoRCAReason)
            worksheet.write(row1,12,todos[row].RootCauseCategory)
            worksheet.write(row1,13,todos[row].FunctionArea)
            worksheet.write(row1,14,todos[row].CodeDeficiencyDescription)
            worksheet.write(row1,15,todos[row].CorrectionDescription)
            worksheet.write(row1,16,todos[row].RootCause)
            if todos[row].IsLongCycleTime is 'Yes':
                todo=TodoLongCycleTimeRCA.query.filter_by(PRID = todos[row].PRID).first()
                worksheet.write(row1,17,todo.LongCycleTimeRootCause)
            else:
                worksheet.write(row1,17,'N/A')
            worksheet.write(row1,18,todos[row].IntroducedBy)
            worksheet.write(row1,19,todos[row].Handler)            

        writer.close() 
        output.seek(0) 
        return output            

class apExcel():  
    def export(self):

        output = BytesIO() 

        writer = pd.ExcelWriter(output, engine='xlwt')
        workbook = writer.book

        worksheet= workbook.add_sheet('sheet1',cell_overwrite_ok=True)
        col=0
        row=1
        pattern = xlwt.Pattern() # Create the Pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        #style = xlwt.XFStyle() # Create the Pattern
        font = xlwt.Font() # Create the Font
        font.name = 'Times New Roman'
        font.bold = True
        #font.underline = True
        #font.italic = True
        style = xlwt.XFStyle() # Create the Style
        style.font = font # Apply the Font to the Style
        style.pattern = pattern # Add Pattern to Style

        columns=['APID','PRID','APDescription','APCreatedDate','APDueDate','APCompletedOn','IsApCompleted','APAssingnedTo','QualityOwner']
        for item in columns:
            worksheet.col(col).width = 4333 # 3333 = 1" (one inch).
            worksheet.write(0, col,item,style)
            col+=1
        style = xlwt.XFStyle()
        style.num_format_str = 'M/D/YY' # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
        alignment = xlwt.Alignment() # Create Alignment
        alignment.horz = xlwt.Alignment.HORZ_JUSTIFIED # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED,HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        alignment.vert = xlwt.Alignment.VERT_TOP # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        #style = xlwt.XFStyle() # Create Style
        style.alignment = alignment # Add Alignment to Style
        todos=TodoAP.query.filter_by(user_id = g.user.id).order_by(TodoAP.APID.asc()).all()
        nrows=len(todos)
        print ('nrows==%s' %(nrows))
        for row in range(nrows):
            r=row+1
            worksheet.write(r,0,todos[row].APID)
            worksheet.write(r,1,todos[row].PRID,style)
            worksheet.write(r,2,todos[row].APDescription,style)
            worksheet.write(r,3,todos[row].APCreatedDate)
            worksheet.write(r,4,todos[row].APDueDate,style)
            worksheet.write(r,5,todos[row].APCompletedOn,style)
            worksheet.write(r,6,todos[row].IsApCompleted)
            worksheet.write(r,7,todos[row].APAssingnedTo,style)
            worksheet.write(r,8,todos[row].QualityOwner)
            
            """
            for co in columns:
                column=columns.index(co)
                cellvalue=todos[row][column]
                worksheet.write(row,column,cellvalue)
            print('row===%s,index===%s'%(column,cellvalue))
            """
 
        #worksheet.set_column('A:E', 20)  

        writer.close() 
        output.seek(0) 
        return output            

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('rca_dashboard.html')

@app.route("/weather", methods=["GET"])
def weather():
    if request.method == "GET":
        count = Todo.query.filter_by(user_id=g.user.id, IsRcaCompleted='No',\
                                     NoNeedDoRCAReason='').order_by(Todo.PRClosedDate.asc()).count()
        count1 = Todo.query.filter_by(user_id=g.user.id, \
                                     NoNeedDoRCAReason='',IsRcaCompleted='Yes').order_by(Todo.PRClosedDate.asc()).count()
    return jsonify(rcadone=count1,rcaundone=count,user=g.user.username)

@app.route("/totoaltribercastatus", methods=["GET"])
def totoaltribercastatus():
    if request.method == "GET":
        count = Todo.query.filter_by(LteCategory = 'FDD',IsRcaCompleted='No',NoNeedDoRCAReason='').order_by(Todo.PRClosedDate.asc()).count()
        count1 = Todo.query.filter_by(LteCategory = 'FDD',IsRcaCompleted='Yes').order_by(Todo.PRClosedDate.asc()).count()
    return jsonify(rcadone=count1,rcaundone=count,user='Tribe RCA Status')

@app.route("/totaltribeapstatus", methods=["GET"])
def totaltribeapstatus():
    if request.method == "GET":
        data=[]
        count = TodoAP.query.filter_by(IsApCompleted='No',).count()
        count1 = TodoAP.query.filter_by(IsApCompleted='Yes').count()
        aa=(count,count1)
        data.append(aa)
    return jsonify(rcadone=count1,rcaundone=count,user='Tribe AP Status')

@app.route("/rcastatus", methods=["GET"])
def rcastatus():
    if request.method == "GET":
        data=[]
        for team in teams:
            count = Todo.query.filter_by(Handler=team, IsRcaCompleted='No', \
                                         NoNeedDoRCAReason='').order_by(Todo.PRClosedDate.asc()).count()
            count1 = Todo.query.filter_by(Handler=team, \
                                          NoNeedDoRCAReason='', IsRcaCompleted='Yes').order_by(
                                            Todo.PRClosedDate.asc()).count()
            aa=(team,count,count1)
            data.append(aa)

    return jsonify(team = [x[0] for x in data],
                   rcaundone = [x[1] for x in data],
                   rcadone = [x[2] for x in data])



rootcauses=['Function Implementation Error','Coding Error','EFS Misunderstanding','EFS Missing','EFS Error',\
            'Change Alignment Missing','Interface Misunderstanding','Correction Inheritance Missing','Feature Missing',\
            'Change Request Issue','Third Party SW bug']
@app.route("/rootcausedistribution", methods=["GET"])
def rootcausedistribution():
    if request.method == "GET":
        data=[]
        for rootcause in rootcauses:
            count = Todo.query.filter_by(RootCauseCategory=rootcause).order_by(Todo.PRClosedDate.asc()).count()
            aa=(rootcause,count)
            data.append(aa)
        data.sort(key=lambda x: x[1], reverse=True)

    return jsonify(rootcause = [x[0] for x in data],count = [x[1] for x in data],user='Root Cause')

@app.route("/rootcausedistributioncatm", methods=["GET"])
def rootcausedistributioncatm():
    if request.method == "GET":
        data=[]
        for rootcause in rootcauses:
            count = Todo.query.filter_by(RootCauseCategory=rootcause,IsCatM = 'Yes').order_by(Todo.PRClosedDate.asc()).count()
            aa=(rootcause,count)
            data.append(aa)
        data.sort(key=lambda x: x[1], reverse=True)

    return jsonify(rootcause = [x[0] for x in data],count = [x[1] for x in data],user='Root Cause')

functionareas=['DL TD Scheduling','UL TD Scheduling','DL FD Scheduling','UL FD Scheduling','DL HARQ',\
            'UL HARQ','DL LA','UL LA','RACH','SysInfo Scheduling','DL CA','UL CA','CnP','RLF','PM Counter',\
            'TTITrace','Third Party SW']
@app.route("/functionareadistribution", methods=["GET"])
def functionareadistribution():
    if request.method == "GET":
        data=[]
        for functionarea in functionareas:
            count = Todo.query.filter_by(FunctionArea=functionarea).order_by(Todo.PRClosedDate.asc()).count()
            aa=(functionarea,count)
            data.append(aa)
        data.sort(key=lambda x:x[1],reverse=True)

    return jsonify(functionarea = [x[0] for x in data],count = [x[1] for x in data],user='Function Area')

@app.route("/functionareadistributioncatm", methods=["GET"])
def functionareadistributioncatm():
    if request.method == "GET":
        data=[]
        for functionarea in functionareas:
            count = Todo.query.filter_by(FunctionArea=functionarea,IsCatM = 'Yes').order_by(Todo.PRClosedDate.asc()).count()
            aa=(functionarea,count)
            data.append(aa)
        data.sort(key=lambda x:x[1],reverse=True)

    return jsonify(functionarea = [x[0] for x in data],count = [x[1] for x in data],user='Function Area')

@app.route("/apstatus", methods=["GET"])
def apstatus():
    if request.method == "GET":
        data=[]
        for team in teams:
            count = TodoAP.query.filter_by(QualityOwner=team, IsApCompleted='No',).count()
            count1 = TodoAP.query.filter_by(QualityOwner=team,IsApCompleted='Yes').count()
            aa=(team,count,count1)
            data.append(aa)

    return jsonify(team = [x[0] for x in data],
                   rcaundone = [x[1] for x in data],
                   rcadone = [x[2] for x in data])

month={'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun','07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'}
@app.route("/teamapstatus", methods=["GET"])
@login_required
def teamapstatus():
    if request.method == "GET":
        data = []
        keys=month.keys()
        keys.sort()
        for k in keys:
            aparrival = 0
            apdone = 0
            apundone = 0
            count = TodoAP.query.filter_by(QualityOwner=g.user.username, IsApCompleted='No',).count()
            count1 = TodoAP.query.filter_by(QualityOwner=g.user.username,IsApCompleted='Yes').count()
            items = TodoAP.query.filter_by(QualityOwner=g.user.username).all()
            for item in items:
                a = item.APCreatedDate
                ext = a.rsplit('-', 2)[1]
                if k == ext:
                    aparrival = aparrival + 1
                    if item.IsApCompleted == 'Yes':
                        apdone = apdone + 1
                    else:
                        apundone = apundone + 1
            aa=(month[k],aparrival,apdone,apundone)
            data.append(aa)

    return jsonify(month = [x[0] for x in data],
                   aparrival = [x[1] for x in data],
                   apdone = [x[2] for x in data],
                   apundone=[x[3] for x in data],user=g.user.username)

#JIRA

@app.route("/teamjirarcaontimedeliverystatus", methods=["GET"])
@login_required
def teamjirarcaontimedeliverystatus():
    jiraissuestatus = ['Open','In Progress','Reopened']
    if request.method == "GET":
        data = []
        items = Todo.query.filter(Todo.Handler == g.user.username,\
                Todo.JiraRcaBeReqested == 'Yes',Todo.JiraIssueStatus.in_(jiraissuestatus)).order_by(
            Todo.PRClosedDate.asc()).all()
        count = Todo.query.filter(Todo.Handler==g.user.username,\
                Todo.JiraRcaBeReqested == 'Yes',Todo.JiraIssueStatus.in_(jiraissuestatus)).count()

        for item in items:
            a = item.PRID
            b = item.JiraRcaDeliveryOnTimeRating
            c = item.JiraRcaPreparedQualityRating
            aa=(a,b)
            data.append(aa)

    return jsonify(prid= [x[0] for x in data],
                   ontimerating = [x[1] for x in data],
                   #qualityrating=[x[2] for x in data],
                          user = g.user.username)

@app.route("/teamjirarcaqualitystatus", methods=["GET"])
@login_required
def teamjirarcaqualitystatus():
    jiraissuestatus = ['Resolved','Closed']
    if request.method == "GET":
        data = []
        items = Todo.query.filter(Todo.Handler == g.user.username,\
                Todo.JiraRcaBeReqested == 'Yes',Todo.JiraIssueStatus.in_(jiraissuestatus)).order_by(
            Todo.PRClosedDate.asc()).all()
        count = Todo.query.filter(Todo.Handler==g.user.username,\
                Todo.JiraRcaBeReqested == 'Yes',Todo.JiraIssueStatus.in_(jiraissuestatus)).count()

        for item in items:
            a = item.PRID
            b = item.JiraRcaDeliveryOnTimeRating
            c = item.JiraRcaPreparedQualityRating
            aa=(a,c)
            data.append(aa)

    return jsonify(prid= [x[0] for x in data],
                   qualityrating = [x[1] for x in data],
                   #qualityrating=[x[2] for x in data],
                          user = g.user.username)

@app.route("/averageontimerating", methods=["GET"])
def averageontimerating():
    if request.method == "GET":
        data=[]
        for team in teams:
            items = Todo.query.filter(Todo.Handler == team,Todo.JiraRcaBeReqested == 'Yes').all()
            count = Todo.query.filter(Todo.Handler == team,Todo.JiraRcaBeReqested == 'Yes').count()
            if count:
                sum=0
                for item in items:
                    b = item.JiraRcaDeliveryOnTimeRating
                    c = item.JiraRcaPreparedQualityRating
                    if b:
                        sum = sum + b
                a = sum/count
            else:
                a=10

            aa=(team,a)
            data.append(aa)

    return jsonify(team = [x[0] for x in data],averageontimerating = [x[1] for x in data])

@app.route("/averagequalityrating", methods=["GET"])
def averagequalityrating():
    if request.method == "GET":
        data=[]
        for team in teams:
            items = Todo.query.filter(Todo.Handler == team, Todo.JiraRcaBeReqested == 'Yes').all()
            count = Todo.query.filter(Todo.Handler == team, Todo.JiraRcaBeReqested == 'Yes').count()
            if count:
                sum = 0
                for item in items:
                    b = item.JiraRcaDeliveryOnTimeRating
                    c = item.JiraRcaPreparedQualityRating
                    if c:
                        sum = sum + c
                a = sum / count
            else:
                a = 10

            aa = (team, a)
            data.append(aa)

        return jsonify(team=[x[0] for x in data], averageontimerating=[x[1] for x in data])


@app.route("/teamjirarcastatus", methods=["GET"])
def teamjirarcastatus():
    if request.method == "GET":
        count = Todo.query.filter_by(user_id=g.user.id, IsRcaCompleted='No',JiraRcaBeReqested = 'Yes',\
                                     ).order_by(Todo.PRClosedDate.asc()).count()
        count1 = Todo.query.filter_by(user_id=g.user.id, \
                                      JiraRcaBeReqested='Yes',IsRcaCompleted='Yes').order_by(Todo.PRClosedDate.asc()).count()
    return jsonify(rcadone=count1,rcaundone=count,user=g.user.username)

@app.route("/tribercastatus", methods=["GET"])
def tribercastatus():
    if request.method == "GET":
        data=[]
        for team in teams:
            count = Todo.query.filter_by(Handler=team, IsRcaCompleted='No', \
                                         JiraRcaBeReqested='Yes').order_by(Todo.PRClosedDate.asc()).count()
            count1 = Todo.query.filter_by(Handler=team, \
                                          JiraRcaBeReqested='Yes', IsRcaCompleted='Yes').order_by(
                                            Todo.PRClosedDate.asc()).count()
            aa=(team,count,count1)
            data.append(aa)

    return jsonify(team = [x[0] for x in data],
                   rcaundone = [x[1] for x in data],
                   rcadone = [x[2] for x in data])

@app.route("/tribejiraapstatus", methods=["GET"])
def tribejiraapstatus():
    if request.method == "GET":
        data=[]
        for team in teams:
            count = TodoAP.query.filter_by(QualityOwner=team, IsApCompleted='No',CustomerAp = 'Yes').count()
            count1 = TodoAP.query.filter_by(QualityOwner=team,IsApCompleted='Yes',CustomerAp = 'Yes').count()
            aa=(team,count,count1)
            data.append(aa)

    return jsonify(team = [x[0] for x in data],
                   rcaundone = [x[1] for x in data],
                   rcadone = [x[2] for x in data])

month={'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun','07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'}
@app.route("/jirateamapstatus", methods=["GET"])
@login_required
def jirateamapstatus():
    if request.method == "GET":
        data = []
        keys=month.keys()
        keys.sort()
        for k in keys:
            aparrival = 0
            apdone = 0
            apundone = 0
            count = TodoAP.query.filter_by(QualityOwner=g.user.username, IsApCompleted='No',CustomerAp = 'Yes').count()
            count1 = TodoAP.query.filter_by(QualityOwner=g.user.username,IsApCompleted='Yes',CustomerAp = 'Yes').count()
            items = TodoAP.query.filter_by(QualityOwner=g.user.username).all()
            for item in items:
                a = item.APCreatedDate
                ext = a.rsplit('-', 2)[1]
                if k == ext:
                    aparrival = aparrival + 1
                    if item.IsApCompleted == 'Yes':
                        apdone = apdone + 1
                    else:
                        apundone = apundone + 1
            aa=(month[k],aparrival,apdone,apundone)
            data.append(aa)

    return jsonify(month = [x[0] for x in data],
                   aparrival = [x[1] for x in data],
                   apdone = [x[2] for x in data],
                   apundone=[x[3] for x in data],user=g.user.username)
#End JIRA
@app.route('/api/apupload', methods=['POST'], strict_slashes=False)
def api_ap_upload():
    app.config['UPLOAD_FOLDER'] = 'apupload' 
    file_dir = os.path.join(basedir, app.config['UPLOAD_FOLDER'])
    file_dir = os.path.join(file_dir,g.user.username)
    if not os.path.exists(file_dir):
        os.makedirs(file_dir) 
    f=request.files['fileField']
    print f
    if f and allowed_file(f.filename): 
        fname=f.filename
        """
        ext = fname.rsplit('.', 1)[1] 
        unix_time = int(time.time())
        new_filename = str(unix_time)+'.'+ext
        filename=os.path.join(file_dir, new_filename)
        print filename
        f.save(os.path.join(file_dir, new_filename))
        """
        filename=os.path.join(file_dir, fname)
        print filename
        f.save(filename)
        a=import_ap_fromexcel(filename)
        if a is True: 
            flash('AP item have been successfully imported')
        else:
            flash('APs importing failed')
        return redirect(url_for('ap_home'))
    else:
        flash('Invalid Filename!')
        return redirect(url_for('ap_home'))



@app.route('/fromexcel',methods=['GET', 'POST'])
@login_required
def fromexcel():
    if g.user.username not in admin:
        flash('You are not permitted to import before you have passed the Audit!', 'error')
        return redirect(url_for('rca_home'))
    if request.method == 'GET':
        return render_template('fromexcel.html')
    else:
        print request.form['textfield']
        filename=request.form['textfield']
        val=importfromexcel(filename)
        if val==0:
            flash('PRID has been used! please use the recommanded one', 'error')
            todo=Todo.query.order_by(Todo.PRID.desc()).first()
            a = re.sub("\D", "", todo.PRID)
                    #a=filter(str.isdigit, todo.PRID)
            a=int(a)
            a=a+1
                    
            b=len(str(a))
                    #sr=sr+'m'*(9-len(sr))
            PRID='MAC'+'0'*(6-b)+ str(a)
            return render_template('new.html'
                                    ,PRID=PRID)           
        flash('Todo item was successfully imported')
        return redirect(url_for('index'))
    
@app.route('/importapfromexcel',methods=['GET', 'POST'])
@login_required
def importapfromexcel():
    if request.method == 'GET':
        return render_template('ap_fromexcel.html')
    else:
        filename=request.files['fileField']
        print filename
        import_ap_fromexcel(filename)
        flash('AP item have been successfully imported RCA status also has been updated!')
        return redirect(url_for('apindex'))
    return render_template('ap_new.html')

#JIRA Update
shouldhavebeenfound = []
def insert_ap_jira(internaltask_sheet,index,ApCategory):
    PRID = internaltask_sheet.cell_value(2,1).strip()
    if len(PRID)>15:
        prlist  = PRID.split('\n')
        for PRID in prlist:
            count = Todo.query.filter_by(PRID=PRID).count()
            if count:
                PRID = PRID
                break
    APCompletedOn=''
    IsApCompleted='No'
    APDescription = internaltask_sheet.cell_value(index,12)
    if len(APDescription.strip())!=0:
        APCreatedDate=time.strftime('%Y-%m-%d',time.localtime(time.time()))
        APAssingnedTo=internaltask_sheet.cell_value(index,15)
        #ctype:0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
        ctype= internaltask_sheet.cell(index,17).ctype
        APDueDate= internaltask_sheet.cell_value(index,17)
        if ctype==3:
            date = datetime(*xldate_as_tuple(APDueDate, 0))
            APDueDate = date.strftime('%Y-%m-%d')
        else:
            APDueDate = APDueDate
        #APDueDate=APDueDate.strftime('%Y-%d-%m')

        QualityOwner=User.query.get(g.user.id).username #g.user.id

        #New added for JIRA
        InjectionRootCauseEdaCause = internaltask_sheet.cell_value(index,11) # RCA/EDA RootCause
        RcaEdaCauseType = internaltask_sheet.cell_value(index,13).strip()
        RcaEdaActionType = internaltask_sheet.cell_value(index,14)
        TargetRelease = ""
        CustomerAp = "Yes"
        ApCategory = ApCategory # RCA/EDA
        ShouldHaveBeenDetected = internaltask_sheet.cell_value(index,18)
        ApJiraId = ""
        #End of added new field for JIRA
        global PR_AP_Dict
        if len(PR_AP_Dict[PRID]): # Reuse Old APID
            APID = PR_AP_Dict[PRID].pop(0)
            todo_item = TodoAP.query.get(APID)
            todo_item.APDescription = APDescription
            todo_item.InjectionRootCauseEdaCause = InjectionRootCauseEdaCause
            todo_item.APDueDate = APDueDate
            todo_item.APAssingnedTo = APAssingnedTo
            db.session.commit()
        else: # Add New APID
            APID = get_apid()
            todoap = TodoAP(APID,PRID,APDescription,APCreatedDate,APDueDate,APCompletedOn,IsApCompleted,APAssingnedTo,QualityOwner,\
                            InjectionRootCauseEdaCause, RcaEdaCauseType, RcaEdaActionType, TargetRelease, CustomerAp,ApCategory,\
                            ShouldHaveBeenDetected,ApJiraId)
            todoap.user = g.user
            db.session.add(todoap)
            db.session.commit()

    global shouldhavebeenfound
    ShouldHaveBeenDetected = internaltask_sheet.cell_value(index, 18)
    if ShouldHaveBeenDetected:
        shouldhavebeenfound.append(FaultShouldHaveBeenFound[ShouldHaveBeenDetected])


def update_rca_jira(PRID,internaltask_sheet,RcaSubtaskJiraId):
    todo_item = Todo.query.get(PRID)
    if todo_item is None:
        flash('Please check the PR, seems it is not in the Formal RCA PR list!', 'error')
        return False
    #todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    #todo_item.IsRcaCompleted = 'Yes'
    aa=findRootCauseIndex(16,'Root Cause',internaltask_sheet)
    todo_item.RootCause = internaltask_sheet.cell_value(aa,11).encode('utf-8')
    #todo_item.RootCause  = internaltask_sheet.cell_value(20,11)
    s2 ='Triggering scenario category: select appropriate item from the list'
    s1 = internaltask_sheet.row_values(6)[1]
    if s1.strip() == s2 :
        todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(7)[3]
        todo_item.CorrectionDescription  = internaltask_sheet.row_values(8)[3]
        #self.LteCategory = LteCategory
        #self.CustomerOrInternal = CustomerOrInternal
        #self.JiraFunctionArea = JiraFunctionArea
        todo_item.TriggerScenarioCategory = internaltask_sheet.row_values(6)[3]
        #self.FirstFaultEcapePhase = FirstFaultEcapePhase
        todo_item.FaultIntroducedRelease = internaltask_sheet.row_values(10)[3]
        #self.TechnicalRootCause = TechnicalRootCause
        #self.TeamAssessor = TeamAssessor
        #self.EdaCause = EdaCause
        #self.RcaRootCause5WhyAnalysis = RcaRootCause5WhyAnalysis
       # self.JiraRcaBeReqested = JiraRcaBeReqested
        #self.JiraIssueStatus = JiraIssueStatus
        #self.JiraIssueAssignee = JiraIssueAssignee
        #self.JiraRcaPreparedQualityRating = JiraRcaPreparedQualityRating
        #self.JiraRcaDeliveryOnTimeRating = JiraRcaDeliveryOnTimeRating
    else:
        todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(6)[3]
        todo_item.CorrectionDescription  = internaltask_sheet.row_values(7)[3]
    """
    with UseDatabase(app.config['dbconfig']) as cursor:
        _SQL = "alter table rcastatus MODIFY column FunctionArea VARCHAR(1024)"
        cursor.execute(_SQL)

    todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(6)[3]
    todo_item.CorrectionDescription  = internaltask_sheet.row_values(7)[3]
	"""
    todo_item.RcaSubtaskJiraId = RcaSubtaskJiraId
    db.session.commit()
    return True

PR_AP_Dict={}

def readworkbook(filename):
    _conn_status = True
    _conn_retry_count = 0
    while _conn_status:
        try:
            print 'Reading Excel...'
            workbook = xlrd.open_workbook(filename)
            _conn_status = False
            return workbook
        except:
            _conn_retry_count += 1
            print ("_conn_retry_count=%d"%_conn_retry_count)
        print 'Excel Reading Error'
        time.sleep(8)
        contiue

def import_ap_fromexcel_jira(filename,RcaSubtaskJiraId):
    #workbook = xlrd.open_workbook(filename)
    workbook = readworkbook(filename)
    internaltask_sheet = workbook.sheet_by_name(r'RcaEda')
    rows = internaltask_sheet.row_values(0)
    nrows = internaltask_sheet.nrows
    ncols = internaltask_sheet.ncols
    print ("nrows==%d" % nrows)
    print ("#############nrows==%d" % ncols)
    PRID = internaltask_sheet.cell_value(2,1).strip()
    if len(PRID)>15:
        prlist  = PRID.split('\n')
        for PRID in prlist:
            count = Todo.query.filter_by(PRID=PRID).count()
            if count:
                PRID = PRID
                break
    APCompletedOn = ''
    IsApCompleted = ''
    QualityOwner = User.query.get(g.user.id).username  # g.user.id
    todo = TodoAP.query.filter_by(PRID=PRID).order_by(TodoAP.APID.asc()).first()
    state = 0
    a = update_rca_jira(PRID, internaltask_sheet,RcaSubtaskJiraId)
    if a is False:
        print ("PRID=%s is not in the PR list of RCA candidate" % PRID)
        return False

    b = find5whyIndex(16, 'Root Cause Analysis', internaltask_sheet)
    prid = str(PRID)

    todo_items = TodoAP.query.filter_by(PRID=PRID).order_by(TodoAP.APID.asc()).all()
    global PR_AP_Dict
    aplist=[]
    PR_AP_Dict[PRID]=[]
    if todo_items:
        for item in todo_items:
            aplist.append(item.APID)
            PR_AP_Dict[PRID].append(item.APID)

    if state == 0:
        b = findIndex(16, 'Action Proposal', internaltask_sheet)
        for i in range(4):
            index = b + i
            ApCategory = 'RCA'
            insert_ap_jira(internaltask_sheet, index,ApCategory)

        b = findIndex(30, 'Action Proposal', internaltask_sheet)
        ApCategory = 'EDA'
        for i in range(3):
            index = b + i
            insert_ap_jira(internaltask_sheet, index,ApCategory)
        c = findIndex(b, 'Action Proposal', internaltask_sheet)
        for i in range(3):
            index = c + i
            insert_ap_jira(internaltask_sheet, index,ApCategory)
        d = findIndex(c, 'Action Proposal', internaltask_sheet)
        for i in range(3):
            index = d + i
            insert_ap_jira(internaltask_sheet, index, ApCategory)
        e = findIndex(d, 'Action Proposal', internaltask_sheet)
        for i in range(3):
            index = e + i
            insert_ap_jira(internaltask_sheet, index, ApCategory)
        f = findIndex(e, 'Action Proposal', internaltask_sheet)
        for i in range(3):
            index = f + i
            insert_ap_jira(internaltask_sheet, index, ApCategory)
        gg = findIndex(f, 'Action Proposal', internaltask_sheet)
        if gg is None:
            return True
        for i in range(3):
            index = gg + i
            insert_ap_jira(internaltask_sheet, index, ApCategory)
        h = findIndex(gg, 'Action Proposal', internaltask_sheet)
        if h is None:
            return True
        for i in range(3):
            index = h + i
            insert_ap_jira(internaltask_sheet, index, ApCategory)
    return True


JIRA_STATUS = {'Closed': 2, 'Reopened' : 3, 'In Progress' : 4, 'Resolved' : 5}

TriggeringCategory = {  'Installation & Startup': u'219294',
                        'SW Upgrade':u'219295',
                        'SW Fallback':u'219296',
                        'Process of Configuration / Reconfigurationu':u'219297',
                        'Not Supported Configuration':u'219298',
                        'OAM Operations':u'219299',
                        'Feature Interaction & interoperability':u'219300',
                        'OAM Robustness (high Load / stressful scenarios/long duration)':u'219301',
                        'Telecom Robustness (high Load / mobility /stressful scenarios/long duration)':u'219302',
                        'Debug (Counters/Alarms/Trace/Resource monitoring)':u'219303',
                        'Reset &amp; Recovery':u'219304',
                        'HW Failure':u'219305',
                        'Required customer/vendor specific equipment':u'219306',
                        'Unknown triggering scenario':u'219307'
                        }

FaultShouldHaveBeenFound = {'Unit/System Component Test (UT/SCT)':{'id':u'219533'},
                            'Entity Test - Regression':{'id':u'219534'},
                            'Entity Test - New Feature':{'id':u'219535'},
                            'System Test - Regression':{'id':u'219536'},
                            'System Test - New Feature':{'id':u'219537'},
                            'Interoperability Testing':{'id':u'219538'},
                            'Performance Testing':{'id':u'219539'},
                            'Stability Testing':{'id':u'219540'},
                            'Negative/Adversarial Testing':{'id':u'219541'},
                            'Network Level Testing':{'id':u'219542'},
                            'First Office Application':{'id':u'219543'}}

RcaCauseType = {'None':								        {'id':u'-1'},
                'Standards':								{'id':u'192644'},
                'HW Problem':                               {'id':u'192645'},
                '3rd Party Products':                       {'id':u'192646'},
                'Missing Requirement':                      {'id':u'192647'},
                'Incorrect Requirements':                   {'id':u'192648'},
                'Unclear Requirements':                     {'id':u'192649'},
                'Changing Requirements':                    {'id':u'192650'},
                'Feature Interaction':                      {'id':u'192651'},
                'Design Deficiency':                        {'id':u'192652'},
                'Lack of Design Detail':                    {'id':u'192653'},
                'Defective Fix':                            {'id':u'192654'},
                'Feature  Enhancement':                     {'id':u'192655'},
                'Code Standards':                           {'id':u'192656'},
                'Configuration Mgmt - Merge':               {'id':u'192658'},
                'Documentation Issue':                      {'id':u'192659'},
                'Dependencies from Other Documentation':    {'id':u'192660'},
                'Code Error':                               {'id':u'192661'},
                'Code Complexity':                          {'id':u'192662'}
                }

EdaCauseType = {
                'None':								                       {'id':u'-1'},
                'Traceability to Requirements':							   {'id':u'193960'},
                'incorrect Testcase':                                      {'id':u'193961'},
                'Scenario Not predicted - Customer Specific Test Scope':   {'id':u'193962'},
                'Module/Unit tests - Test Scope':                          {'id':u'193963'},
                'Entity/Feature Tests - Test Scope':                       {'id':u'193964'},
                'Regression Tests -Test Scope':                            {'id':u'193965'},
                'Network Verification  -- Test Scope':                     {'id':u'193966'},
                'Solution Verification - Test Scope':                      {'id':u'193967'},
                'Performance - Test Scope':                                {'id':u'193968'},
                'Non-Functianl  - Test Scope':                             {'id':u'193969'},
                'Maintenance Test Scope':                                  {'id':u'193970'},
                'E2E  Verification -- Test Scope':                         {'id':u'193971'},
                'Test Plan':                                               {'id':u'193972'},
                'Test Strategy --  Content missing/not covered/Configuration not tested': {'id':u'193973'},
                'Missing Test case review':                                {'id':u'193974'},
                'Unknown/Unclear test configurations':                     {'id':u'193975'},
                'Unforeseen upgrade Paths':                                {'id':u'193976'},
                'Unclear Load Model':                                {'id':u'193977'},
                'Missing Compatibility Tests':                       {'id':u'193978'},
                'Code review missed':                                {'id':u'193979'},
                'Configuration missing/misunderstood/incomplete':              {'id':u'219275'},
                'Scenario missing/misunderstood/incomplete':                   {'id':u'219276'},
                'Extraordinary scenario - Very difficult to recreate defect':  {'id':u'219277'},
                'Customer documentation not reviewed/tested':                  {'id':u'219278'},
                'Functional Testcase missing/misunderstood/incomplete':        {'id':u'219279'},
                'Non-Functional Testcase missing/misunderstood/incomplete':    {'id':u'219280'},
                'Regression Testcase missing/misunderstood/incomplete':        {'id':u'219281'},
                'Robustness Testcase missing/misunderstood/incomplete':        {'id':u'219282'},
                'Test Blocked : Lack of Hardware/Test Tools':                  {'id':u'219283'},
                'Test Blocked : Needed SW not available':                      {'id':u'219284'},
                'Test run too early in the release':                           {'id':u'219285'},
                'Planned Test run post delivery':                              {'id':u'219286'},
                'Planned Test not run':                                {'id':u'219287'},
                'Out of Testing Scope':                                {'id':u'219288'},
                'insufficient testing of changes/fixes':                                {'id':u'219289'},
                'insufficient test duration/iterations':                                {'id':u'219290'},
                'Fix was available but was not delivered':                              {'id':u'219291'},
                'incorrect analysis of test result, marked passed when actually failed': {'id':u'219292'},
                'intermittent defect':                                  {'id':u'219293'}}


#JiraSubTask Case Type
JiraSubTaskCaseType = {'RCA':u'219575','EDA':u'219576','RCA and EDA':'219892'}
#Analysis subtask:19800,Action for RCA:19804,Action for EDA:19806
JiraChildIssueType = {'Analysis subtask':u'19800','Action for RCA':u'19804','Action for EDA':u'19806'}

JiraChildIssueType ={'RCA':u'19804','EDA':u'19806'}

def sharepointfile(PRID):
    SHAREPOINT_COMMON_USER = 'enqing.lei@nokia-sbell.com'
    SHAREPOINT_COMMON_PWD = 'kkcqqgfmmxkgjjmd'
    #site_url = 'https://nokia.sharepoint.com/sites/LTERCA/_api/web'
    #link = 'https://nokia.sharepoint.com/'
    #url ='https://nokia.sharepoint.com/sites/LTERCA/'
    site_url ='https://nokia.sharepoint.com/sites/LTERCA/'
    local_file ='c://python/1.xlsm'
    item_list_url = "https://nokia.sharepoint.com/:x:/r/sites/LTERCA/RcaStore/" + PRID + ".xlsm"
    app.config['UPLOAD_FOLDER'] = 'apupload'
    file_dir = os.path.join(basedir, app.config['UPLOAD_FOLDER'])
    file_dir = os.path.join(file_dir,g.user.username)
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
    filename = PRID + '.xlsm'
    filename = os.path.join(file_dir, filename)

    ctx_auth = AuthenticationContext(site_url)
    print "Authenticate credentials"
    if ctx_auth.acquire_token_for_user(SHAREPOINT_COMMON_USER, SHAREPOINT_COMMON_PWD):
        ctx = ClientContext(site_url, ctx_auth)
        request = ClientRequest(ctx_auth)
        options = RequestOptions(site_url)
        ctx_auth.authenticate_request(options)
        ctx.ensure_form_digest(options)
        result = requests.get(url=item_list_url,
                              headers=options.headers,
                              auth=options.auth)
        print result.content
        if result.ok:
            with open(filename, 'wb') as file:
                file.write(result.content)
            return filename
        return False
    return False



@app.route('/update2JIRA',methods=['GET', 'POST'])
def update2JIRA():
    if request.method == 'GET':
        return render_template('JIRAlogin.html')
    else:
        global gPRID
        PRID = gPRID
        username = request.form['username']
        password = request.form['password']

        options = {'server': 'https://jiradc.int.net.nokia.com/'}
        try:
            jira = JIRA(options, basic_auth=(username, password))
        except:
            flash('Invalid login,Please login with your own JIRA account and password!!!')
            return render_template('JIRAlogin.html')
        RcaAnalysis_issues = jira.search_issues(jql_str='project = MNRCA AND summary~'+PRID+' AND type = "Analysis subtask" AND status not in (Resolved, Closed)',maxResults=5)
        Rca5Why_issues = jira.search_issues(jql_str='project = MNRCA AND summary~'+PRID+' AND type = "5WhyRCA"',maxResults=5)
        print RcaAnalysis_issues
        print len(RcaAnalysis_issues)
        link = RcaAnalysis_issues[0].fields.customfield_37064  # SharePoint Link
        print link
        """
        url = "https://nokia.sharepoint.com/:x:/r/sites/LTERCA/RcaStore/"+PRID+".xlsm"
        a = webbrowser.open(url)
        time.sleep(8)
        filename ="c:\users\qmxh38\Downloads\\"+PRID+".xlsm"
        """
        RcaSubtaskJiraId=RcaAnalysis_issues[0].id
        RcaSubtaskJiraKey = RcaAnalysis_issues[0].key
        Rca5WhyTaskJiraKey = Rca5Why_issues[0].key
        filename = sharepointfile(PRID)
        if filename != False:
            import_ap_fromexcel_jira(filename,RcaSubtaskJiraKey)
        else:
            flash('SharePoint Excel Error!!!!!')
            return render_template('JIRAlogin.html')

        #Jira RCA Analysis Subtask
        issueaddforRCAsubtask = {
            'project': {'id': u'41675'},
            'parent':{"key": "MNRCA-15563"},
            'issuetype': {'id': u'19800'},
            'summary': 'testSubtask  RCA for PR1234',
            'customfield_10464':{'id':u'219575'},# Case Type
            'assignee':'enqing.lei@nokia-sbell.com'
        }

        todos = TodoAP.query.filter_by(PRID = PRID, IsApCompleted='No').all()
        if len(todos) !=0:
            for todo_item in todos:
                ChildSubtaskType = JiraChildIssueType[todo_item.ApCategory]

                if todo_item.ApCategory == 'RCA':
                    TitleSummary = 'RCA Action For ' + PRID
                else:
                    TitleSummary = 'EDA Action For ' + PRID
                """
                    if todo_item.ShouldHaveBeenDetected:
                        shouldhavebeenfound.append(FaultShouldHaveBeenFound[todo_item.ShouldHaveBeenDetected])
                if shouldhavebeenfound:
                    dict1 = {'customfield_37199': shouldhavebeenfound}
                    Rca5Why_issues[0].update(dict1)
                """
                ProposedActionDescription = todo_item.APDescription
                TargetDate = todo_item.APDueDate
                RootCause = todo_item.InjectionRootCauseEdaCause
                JiraIssueAssignee = todo_item.APAssingnedTo
                with UseDatabaseDict(app.config['dbconfig']) as cursor:
                    _SQL = "select * from user_info where email = '" + JiraIssueAssignee + "'"
                    cursor.execute(_SQL)
                    items = cursor.fetchall()
                if len(items) != 0:
                    emailname = items[0]['accountId'].encode('utf-8')
                else:# Cannot find assignee, assign to TeamAssessor
                    todos = Todo.query.get(todo_item.PRID)
                    JiraIssueAssignee = todos .TeamAssessor
                    #JiraIssueAssignee='kefeng.liu@nokia-sbell.com'
                    with UseDatabaseDict(app.config['dbconfig']) as cursor:
                        _SQL = "select * from user_info where email = '" + JiraIssueAssignee + "'"
                        cursor.execute(_SQL)
                        items = cursor.fetchall()
                        emailname = items[0]['accountId'].encode('utf-8')
                    #emailname = 'bqgk83'# Xia Shujie
                if todo_item.ApCategory == 'RCA':
                    CauseType = RcaCauseType[todo_item.RcaEdaCauseType]
                    issueaddforrca_action = {
                        'project': {'id': u'41675'},
                        'parent':{"key": Rca5WhyTaskJiraKey},
                        'issuetype': {'id': ChildSubtaskType},# Child subtask type
                        'summary': TitleSummary,
                        #'customfield_37543':{'id':'193986'},# Action Type conflict with excel
                        'customfield_37480':ProposedActionDescription,
                        'customfield_37070':TargetDate,# Target Date
                        #'customfield_38021':'TL18', # Target Release conflict with excel
                        'customfield_37755':RootCause,#InjectionRootCauseEdaCause
                        'customfield_38019':CauseType,# RCA Cause Type,conflict with excel
                        'assignee':{'name':emailname}
                    }
                    newissue = jira.create_issue(issueaddforrca_action)
                elif todo_item.ApCategory == 'EDA':
                    CauseType = EdaCauseType[todo_item.RcaEdaCauseType]
                    issueaddforrca_action = {
                        'project': {'id': u'41675'},
                        'parent':{"key": Rca5WhyTaskJiraKey},
                        'issuetype': {'id': ChildSubtaskType},# Child subtask type
                        'summary': TitleSummary,
                        #'customfield_37543':{'id':'193986'},# Action Type conflict with excel
                        'customfield_37472':ProposedActionDescription,
                        'customfield_37058':TargetDate,# Target Date
                        #'customfield_38021':'TL18', # Target Release conflict with excel
                        'customfield_37470':RootCause,#InjectionRootCauseEdaCause
                        'customfield_38091':CauseType,# RCA Cause Type,conflict with excel
                        'assignee':{'name':emailname}
                    }
                    if todo_item.ApJiraId == '':
                        newissue = jira.create_issue(issueaddforrca_action)
                    else:
                        JiraIssueId = todo_item.ApJiraId
                        print JiraIssueId
                        # issue = jira.issue('MNRCA-15563')
                        newissue = jira.issue(JiraIssueId)
                        newissue.update(issueaddforrca_action)

                print newissue.key
                JiraIssueId = str(newissue.key)
                APID = todo_item.APID
                todo_item = TodoAP.query.get(APID)
                todo_item.ApJiraId = JiraIssueId
                todo_item.APAssingnedTo = JiraIssueAssignee
                db.session.commit()

        todo_item = Todo.query.get(PRID)
        TriggeringType = TriggeringCategory[todo_item.TriggerScenarioCategory]
        Rca5Why_issues[0].update(customfield_39290={'id': TriggeringType})
        global shouldhavebeenfound
        if shouldhavebeenfound: # Should have been found update
            dict1 = {'customfield_37199': shouldhavebeenfound}
            Rca5Why_issues[0].update(dict1)
            shouldhavebeenfound = []
        status = RcaAnalysis_issues[0].fields.status
        print status
        if status != 'Closed':
            jira.transition_issue(RcaAnalysis_issues[0], JIRA_STATUS['Resolved'])
        status = RcaAnalysis_issues[0].fields.status
        todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        todo_item.JiraIssueStatus = status
        todo_item.IsRcaCompleted = 'Yes'
        db.session.commit()
        flash('AP item have been successfully imported RCA status also has been updated!')
        return redirect(url_for('show_or_update',PRID=gPRID))
    return render_template('ap_new.html')

@app.route('/toexcel',methods=['GET', 'POST'])
@login_required
def toexcel():
    if request.method == 'GET':
        return render_template('toexcel.html')
    else:
        output=Excel().export()
        resp = make_response(output.getvalue()) 
        resp.headers["Content-Disposition"] ="attachment; filename=rca_pronto_list.xls"
        resp.headers['Content-Type'] = 'application/x-xlsx'
        flash('Export to excel successfully,Getting the result at bottom left!')
        return resp

@app.route('/toapexcel',methods=['GET', 'POST'])
@login_required
def toapexcel():
    if request.method == 'GET':
        return render_template('ap_toexcel.html')
    else:
        output=apExcel().export()
        resp = make_response(output.getvalue()) 
        resp.headers["Content-Disposition"] ="attachment; filename=rca_ap_list.xls"
        resp.headers['Content-Type'] = 'application/x-xlsx'
        flash('Export to excel successfully,Getting the result at bottom left!')
        return resp

# @app.route('/',methods=['GET','POST'])
# #@login_required
# def home1():
#     if request.method=='GET':
#         return render_template('home.html')
#     else:
#         return redirect(url_for('rca_home'))




@app.route('/customer_rca_done',methods=['GET','POST'])
@login_required
def customer_rca_done():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter(Todo.IsRcaCompleted== 'Yes',Todo.JiraRcaBeReqested == 'Yes').order_by(Todo.PRClosedDate.asc()).count()
            todos = Todo.query.filter(Todo.IsRcaCompleted == 'Yes',Todo.JiraRcaBeReqested == 'Yes').order_by(Todo.PRClosedDate.asc()).all()
            headermessage='All Customer RCA Done'
            return render_template('customer_index.html',count= count,headermessage = headermessage,
                                   todos = todos,user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count=Todo.query.filter_by(user_id = g.user.id,JiraRcaBeReqested = 'Yes',IsRcaCompleted ='Yes').order_by(Todo.PRClosedDate.asc()).count()
            headermessage = 'All Customer RCA Done'
            todos = Todo.query.filter_by(user_id=g.user.id,JiraRcaBeReqested = 'Yes',IsRcaCompleted = 'Yes').\
                order_by(Todo.PRClosedDate.asc()).all()
            return render_template('customer_index.html',count=count,headermessage = headermessage,
                               todos= todos,\
                               user=User.query.get(g.user.id).username + '  Logged in')

@app.route('/customer_rca_undone',methods=['GET','POST'])
@login_required
def customer_rca_undone():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter_by(JiraRcaBeReqested = 'Yes',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).count()
            todos = Todo.query.filter_by(JiraRcaBeReqested = 'Yes', IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).all()
            headermessage = 'All Customer RCA UnDone'
            return render_template('customer_index.html',count= count,headermessage = headermessage,
                                   todos = todos,user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count = Todo.query.filter_by(user_id = g.user.id,\
                    JiraRcaBeReqested='Yes',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).count()
            todos = Todo.query.filter_by(user_id=g.user.id, \
                    JiraRcaBeReqested='Yes', IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).all()
            headermessage = 'All Customer RCA UnDone'
            return render_template('customer_index.html',count=count,headermessage = headermessage,
                               todos=todos,user=User.query.get(g.user.id).username + '  Logged in')

@app.route('/rca_done',methods=['GET','POST'])
@login_required
def rca_done():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter(Todo.IsRcaCompleted== 'Yes').order_by(Todo.PRClosedDate.asc()).count()
            todos = Todo.query.filter(Todo.IsRcaCompleted == 'Yes').order_by(Todo.PRClosedDate.asc()).all()
            headermessage='All RCA Done'
            return render_template('index.html',count= count,headermessage = headermessage,
                                   todos = todos,user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count=Todo.query.filter_by(user_id = g.user.id).filter(Todo.IsRcaCompleted=='Yes').order_by(Todo.PRClosedDate.asc()).count()
            headermessage = 'All RCA Done'
            todos = Todo.query.filter_by(user_id=g.user.id).filter(Todo.IsRcaCompleted=='Yes').\
                order_by(Todo.PRClosedDate.asc()).all()
            return render_template('index.html',count=count,headermessage = headermessage,
                               todos= todos,\
                               user=User.query.get(g.user.id).username + '  Logged in')

@app.route('/rca_undone',methods=['GET','POST'])
@login_required
def rca_undone():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter_by( NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).count()
            headermessage = 'All RCA UnDone'
            return render_template('index.html',count= count,headermessage = headermessage,
                                   todos=Todo.query.filter_by( \
                                                              NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(
                                       Todo.PRClosedDate.asc()).all(), \
                                   user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).count()
            headermessage = 'All RCA UnDone'
            return render_template('index.html',count=count,headermessage = headermessage,
                               todos=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).all(),\
                               user=User.query.get(g.user.id).username + '  Logged in')

@app.route('/rca_noneed',methods=['GET','POST'])
@login_required
def rca_noneed():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter(and_(Todo.NoNeedDoRCAReason != '',Todo.IsRcaCompleted == 'No')).order_by(Todo.PRClosedDate.asc()).count()
            todos = Todo.query.filter(and_(Todo.NoNeedDoRCAReason != '', Todo.IsRcaCompleted == 'No')).order_by(
                Todo.PRClosedDate.asc()).all()
            headermessage = 'All RCA NoNeed'
            return render_template('index.html',count= count,headermessage = headermessage,
                                   todos=todos,\
                                   user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).count()
            count = Todo.query.filter(and_(Todo.user_id == g.user.id,Todo.NoNeedDoRCAReason != '',Todo.IsRcaCompleted == 'No')).order_by(Todo.PRClosedDate.asc()).count()
            todos = Todo.query.filter(and_(Todo.user_id == g.user.id,Todo.NoNeedDoRCAReason != '', Todo.IsRcaCompleted == 'No')).order_by(
                Todo.PRClosedDate.asc()).all()
            headermessage = 'All RCA NoNeed'
            return render_template('index.html',count=count,headermessage = headermessage,
                               todos= todos,\
                               user=User.query.get(g.user.id).username + '  Logged in')


@app.route('/rca_jira',methods=['GET','POST'])
@login_required
def rca_jira():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter(Todo.JiraRcaBeReqested== 'Yes').order_by(Todo.PRClosedDate.asc()).count()
            todos = Todo.query.filter(Todo.JiraRcaBeReqested == 'Yes').order_by(Todo.PRClosedDate.asc()).all()
            headermessage='All RCA OnJIRA'
            return render_template('index.html',count= count,headermessage = headermessage,
                                   todos = todos,user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count=Todo.query.filter_by(user_id = g.user.id).filter(Todo.JiraRcaBeReqested=='Yes').order_by(Todo.PRClosedDate.asc()).count()
            headermessage = 'All RCA OnJIRA'
            todos = Todo.query.filter_by(user_id=g.user.id).filter(Todo.JiraRcaBeReqested=='Yes').\
                order_by(Todo.PRClosedDate.asc()).all()
            return render_template('index.html',count=count,headermessage = headermessage,
                               todos= todos,\
                               user=User.query.get(g.user.id).username + '  Logged in')

@app.route('/longcycletimerca_home',methods=['GET','POST'])
@login_required
def longcycletimerca_home():
    if request.method=='GET':
        if g.user.username in admin:
            return render_template('longcycletimerca_index.html',count= TodoLongCycleTimeRCA.query.count(),
                               todos=TodoLongCycleTimeRCA.query.filter_by(NoNeedDoRCAReason='').order_by(TodoLongCycleTimeRCA.PRClosedDate.asc()).all(),user=User.query.get(g.user.id).username + '  Logged in')
        else:
            return render_template('longcycletimerca_index.html',
                                   todos=TodoLongCycleTimeRCA.query.filter_by(user_id=g.user.id,
                                                                              NoNeedDoRCAReason='').order_by(
                                       TodoLongCycleTimeRCA.PRClosedDate.asc()).all(),
                                   user=User.query.get(g.user.id).username + '  Logged in')


@app.route('/ap_home',methods=['GET','POST'])
@login_required
def ap_home():
    if request.method=='GET':
        if g.user.username in admin:
            count=TodoAP.query.order_by(TodoAP.APID.asc()).count()
            todos = TodoAP.query.order_by(TodoAP.APID.asc()).all()
            username = g.user.username
            return render_template('ap_index.html',count=count,username=username,
                               todos=TodoAP.query.order_by(TodoAP.APID.asc()).all())
        elif g.user.username in teams:
            PRID= TodoAP.PRID
            username = g.user.username
            filename = PRID+'.xlsm'
            count=TodoAP.query.filter_by(user_id=g.user.id).order_by(TodoAP.APID.asc()).count()
            #gEvidenceFileName = "http://n-5cg5010gn7.nsn-intra.net:5001/" + 'apupload' + '/'+ username + '/' + fname
            return render_template('ap_index.html',count=count,username = username,
                                   todos=TodoAP.query.filter_by(user_id=g.user.id).order_by(TodoAP.APID.asc()).all())
        else:
            hello = User.query.filter_by(username=g.user.username).first()
            if hello:
                normaluser = hello.email
                count = TodoAP.query.filter_by(APAssingnedTo=normaluser).order_by(TodoAP.APID.asc()).count()
                return render_template('ap_index.html', count=count,
                                       todos=TodoAP.query.filter_by(APAssingnedTo=normaluser).order_by(
                                           TodoAP.APID.asc()).all())

    else:
        output=Excel().export()
        resp = make_response(output.getvalue()) 
        resp.headers["Content-Disposition"] ="attachment; filename=ap_list.xls"
        resp.headers['Content-Type'] = 'application/x-xlsx'
        return resp



@app.route('/newlongcycletimerca', methods=['GET', 'POST'])
@login_required
def newlongcycletimerca():
    if request.method == 'POST':
        if not request.form['PRID']:
            flash('PRID is required', 'error')
        elif not request.form['PRTitle']:
            flash('PRTitle is required', 'error')
        else:
            PRID=request.form['PRID'].strip()

            todo=TodoLongCycleTimeRCA.query.filter_by(PRID=PRID).first()
            if todo is not None:
                flash('PRID has been used!', 'error')
                return render_template('longcycletimerca_new.html'
                                       ,PRID=PRID)
            PRID = request.form['PRID']
            PRTitle = request.form['PRTitle']
            PRClosedDate  = request.form['PRClosedDate']
            PRReportedDate = request.form['PRReportedDate']
            PROpenDays=daysBetweenDate(PRReportedDate,PRClosedDate)
            PRRcaCompleteDate = request.form['PRRcaCompleteDate']            
            IsCatM  = request.form['IsCatM']
            LongCycleTimeRcaIsCompleted = request.form['LongCycleTimeRcaIsCompleted']
            if daysBetweenDate(PRReportedDate,PRClosedDate) > 14:
                IsLongCycleTime = 'Yes'
            else:
                IsLongCycleTime = 'No'
            NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']

            LongCycleTimeRootCause = request.form['LongCycleTimeRootCause']

            Handler  = request.form['Handler']
            
            todo = TodoLongCycleTimeRCA(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,IsLongCycleTime,\
                                         IsCatM,LongCycleTimeRcaIsCompleted,LongCycleTimeRootCause,NoNeedDoRCAReason,Handler)
            todo.user = g.user
            db.session.add(todo)
            db.session.commit()
            flash('LongCycleTime RCA item was successfully created')
            return redirect(url_for('longcycletimercaindex'))

    return render_template('longcycletimerca_new.html')


@app.route('/newap', methods=['GET', 'POST'])
@login_required
def newap():
    if request.method == 'POST':
        if not request.form['APID']:
            flash('APID is required', 'error')
        elif not request.form['PRID']:
            flash('PRID is required', 'error')
        elif not request.form['APDescription']:
            flash('APDescription is required', 'error')
        else:
            PRID=request.form['PRID']
            APID=request.form['APID']     
            todo=TodoAP.query.filter_by(APID=APID).first()
            if todo is not None:
                flash('APID has been used!', 'error')
                return render_template('ap_new.html',APID=get_apid())
            
            APDescription = request.form['APDescription']
            APCreatedDate = request.form['APCreatedDate']
            APDueDate  = request.form['APDueDate']
            APCompletedOn = request.form['APCompletedOn']           
            IsApCompleted  = request.form['IsApCompleted']
            APAssingnedTo = request.form['APAssingnedTo']
            QualityOwner  = request.form['QualityOwner']

            todo = TodoAP(APID,PRID,APDescription,APCreatedDate,APDueDate,APCompletedOn,IsApCompleted,APAssingnedTo,QualityOwner)
            todo.user = g.user
            db.session.add(todo)
            db.session.commit()
            flash('AP item was successfully created')
            return redirect(url_for('apindex'))

    return render_template('ap_new.html',APID=get_apid())

def update_rca_team(PRID,internaltask_sheet):
    todo_item = Todo.query.get(PRID)
    todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))    
    todo_item.IsRcaCompleted = 'Yes'
    s1 = internaltask_sheet.row_values(6)[1]
    if s1.strip() == 'Triggering scenario category: select appropriate item from the list':
        todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(7)[3]
        todo_item.CorrectionDescription  = internaltask_sheet.row_values(8)[3]
    else:
        todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(6)[3]
        todo_item.CorrectionDescription  = internaltask_sheet.row_values(7)[3]	    
    aa=findRootCauseIndex(16,'Root Cause',internaltask_sheet)
    todo_item.RootCause = internaltask_sheet.cell_value(aa,11)
    db.session.commit()

def update_longcycletimercatable(PRID,request):
    todo_item = TodoLongCycleTimeRCA.query.get(PRID)

    todo_item.PRTitle = request.form['PRTitle']
    todo_item.PRClosedDate  = request.form['PRClosedDate']
    todo_item.PRReportedDate = request.form['PRReportedDate']
    todo_item.PROpenDays = daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate)

    todo_item.PRRcaCompleteDate = request.form['PRRcaCompleteDate']

    if daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'
    todo_item.IsCatM  = request.form['IsCatM']
    todo_item.LongCycleTimeRcaIsCompleted = request.form['IsRcaCompleted']
    #todo_item.NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']
    todo_item.Handler  = request.form['Handler']
    team=request.form['Handler']

    hello = User.query.filter_by(username=team).first()
    todo_item.user_id=hello.id      
    db.session.commit()


gPRID='CAS-xxx'

@app.route('/todos/<PRID>', methods = ['GET' , 'POST'])
@login_required
def show_or_update(PRID):
    global gPRID
    gPRID = PRID
    todo_item = Todo.query.get(PRID)
    if request.method == 'GET':
        #loginAccount,operationType,prIdorApId,log1,log2,log3,log4,log5,log6,log7,log8,log9,log10,log11,log12
        # syslog_record(g.user.username,"reading",todo_item.PRID,todo_item.IsCatM,todo_item.IsRcaCompleted,\
        #               todo_item.RootCauseCategory,todo_item.FunctionArea,todo_item.IntroducedBy,todo_item.Handler,\
        #               todo_item.TeamAssessor,todo_item.JiraIssueAssignee,'','','','')
        #log_request2(get_login_info())
        # if todo_item.user_id == g.user.id or g.user.username in admin:
        #     print ("todo_item.user_id=%d"%todo_item.user_id)
        #     print ("g.user.id=%d"%g.user.id)
        #     print "before"
        HIs = {'a':'1','b':'2'}
        return render_template('view.html',todo=todo_item)
        # else:
        #     print "after"
        #     print ("todo_item.user_id=%d"%todo_item.user_id)
        #     print ("g.user.id=%d"%g.user.id)
        #     flash('This PR is not under your account,You are not authorized to edit this item, Please login with correct account', 'error')
        #     return redirect(url_for('logout'))
    elif request.method == 'POST':
        value = request.form['button']
        if value == 'Update':
            #if g.user.username in admin:
            if todo_item.user.id == g.user.id or g.user.username in admin:
                todo_item.PRID = request.form['PRID'].strip()
                PRID=todo_item.PRID

                return redirect(url_for('index'))
            else:
                flash('You are not authorized to edit this item','error')
                return redirect(url_for('show_or_update',PRID=PRID))
        elif value == 'Delete'and g.user.username in admin:
            todo_item = Todo.query.get(PRID)
            syslog_record(g.user.username, value, todo_item.PRID, todo_item.IsCatM, todo_item.IsRcaCompleted, \
                          todo_item.RootCauseCategory, todo_item.FunctionArea, todo_item.IntroducedBy,
                          todo_item.Handler, \
                          todo_item.TeamAssessor, todo_item.JiraIssueAssignee, '', '', '', '')
            db.session.delete(todo_item)
            todo_longcycle_item = TodoLongCycleTimeRCA.query.get(PRID)
            if todo_longcycle_item:
                db.session.delete(todo_longcycle_item)
            db.session.commit()
            return redirect(url_for('index'))
        flash('You are not authorized to Delete this item', 'error')
        return redirect(url_for('show_or_update', PRID=PRID))

def update_rcatable(PRID,request):
    todo_item = Todo.query.get(PRID)
    todo_item.PRTitle = request.form['PRTitle']
    todo_item.PRClosedDate  = request.form['PRClosedDate']
    todo_item.PRReportedDate = request.form['PRReportedDate']
    todo_item.PROpenDays = daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate)

    todo_item.PRRcaCompleteDate = request.form['PRRcaCompleteDate']


    if daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'
    todo_item.IsCatM  = request.form['IsCatM']
    todo_item.IsRcaCompleted = request.form['LongCycleTimeRcaIsCompleted']
    todo_item.NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']

    todo_item.Handler  = request.form['Handler']
    team=request.form['Handler']

    hello = User.query.filter_by(username=team).first()
    todo_item.user_id=hello.id      
    db.session.commit()

@app.route('/todolongcycletimercas/<PRID>', methods = ['GET' , 'POST'])
@login_required
def show_or_updatelongcycletimerca(PRID):
    todo_item = TodoLongCycleTimeRCA.query.get(PRID)
    if request.method == 'GET':
        return render_template('longcycletimerca_view.html',todo=todo_item)
    if todo_item.user.id == g.user.id:
        todo_item.PRID = request.form['PRID'].strip()
        PRID=todo_item.PRID
        todo_item.PRTitle = request.form['PRTitle']
        todo_item.PRClosedDate  = request.form['PRClosedDate']
        todo_item.PRReportedDate = request.form['PRReportedDate']
        todo_item.PROpenDays = daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate)
        todo_item.PROpenDays = daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate)
  
        todo_item.PRRcaCompleteDate = request.form['PRRcaCompleteDate']

        todo_item.IsLongCycleTime  = request.form['IsLongCycleTime']

        todo_item.IsCatM  = request.form['IsCatM']
        todo_item.LongCycleTimeRcaIsCompleted = request.form['LongCycleTimeRcaIsCompleted']

        todo_item.NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']

        todo_item.LongCycleTimeRootCause = request.form['LongCycleTimeRootCause']
       
        todo_item.Handler  = request.form['Handler']
        team=request.form['Handler']
	hello = User.query.filter_by(username=team).first()
	todo_item.user_id=hello.id

        db.session.commit()

        update_rcatable(PRID,request)
        
        return redirect(url_for('longcycletimercaindex'))
    
    flash('You are not authorized to edit this todo item','error')
    return redirect(url_for('show_or_updatelongcycletimerca',PRID=PRID))

@app.route('/uploadevidence2jira',methods=['GET', 'POST'], strict_slashes=False)
@login_required
def uploadevidence2jira():
    global gEvidenceFileName
    global gAPID
    if request.method == 'GET':
        print "Impossible!"
        return redirect(url_for('apindex'))
    else:
        username = request.form['username']
        password = request.form['password']
        options = {'server': 'https://jiradc.int.net.nokia.com/'}
        #jira = JIRA(options, basic_auth=(aa, bb))
        try:
            jira = JIRA(options, basic_auth=(username, password))
        except:
            flash('Invalid login,Please login with your own JIRA account and password!!!')
            return render_template('JIRAlogin4ApUpdate.html')
        todo_item = TodoAP.query.get(gAPID)
        JiraIssueId = todo_item.ApJiraId
        print JiraIssueId
        #issue = jira.issue('MNRCA-15563')
        issue = jira.issue(JiraIssueId)
        rcaapevidence = gEvidenceFileName
        dict1 = {'customfield_38032': rcaapevidence}
        dict2 =  {'customfield_38032': rcaapevidence}
        issue.update(dict1)
        jira.transition_issue(issue,JIRA_STATUS['Resolved'])
        #todo_item = TodoAP.query.get(gAPID)
        todo_item.APCompletedOn = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        todo_item.IsApCompleted = 'Yes'
        db.session.commit()
        flash('AP Status has been successfully updated!!!')
        return redirect(url_for('show_or_updateap', APID=gAPID))
        #return redirect(url_for('apindex'))

gEvidenceFileName=''

@app.route('/uploadevidence',methods=['GET', 'POST'], strict_slashes=False)
@login_required
def uploadevidence():
    global gEvidenceFileName
    if request.method == 'GET':
        #TaskId = request.form['TaskId']
        return render_template('uploadevidence.html')
    else:
        global gAPID
        file_dir = os.path.join(basedir, 'evidence')
        file_dir = os.path.join(file_dir, gAPID)
        if not os.path.exists(file_dir):
            os.makedirs(file_dir)
        f = request.files['fileField']
        #if f and allowed_file(f.filename):
        fname = f.filename
        filename = os.path.join(file_dir, fname)
        f.save(filename)
        gEvidenceFileName = filename
        """
        todo_item = TodoAP.query.get(gAPID)
        todo_item.APCompletedOn = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        todo_item.IsApCompleted = 'Yes'
        db.session.commit()
        """
        gEvidenceFileName = "http://n-5cg5010gn7.nsn-intra.net:5001/"+'evidence'+'/'+gAPID+'/'+fname
        #flash('Internal task material has been successfully uploaded!!')
        return render_template('JIRAlogin4ApUpdate.html')
        #return redirect(url_for('index'))


gAPID  = 'AP000001'

@app.route('/todoaps/<APID>', methods = ['GET' , 'POST'])
@login_required
def show_or_updateap(APID):
    global gAPID
    gAPID = APID
    todo_item = TodoAP.query.get(APID)
    if request.method == 'GET':
        if todo_item.user_id == g.user.id or g.user.username in admin:
            print ("todo_item.user_id=%d"%todo_item.user_id)
            print ("g.user.id=%d"%g.user.id)
            print "before"
            return render_template('ap_view.html', todo=todo_item)
        else:
            print "after"
            print ("todo_item.user_id=%d"%todo_item.user_id)
            print ("g.user.id=%d"%g.user.id)
            flash('This AP is not under your account,You are not authorized to edit this item, Please login with correct account', 'error')
            return redirect(url_for('logout'))
        #return render_template('ap_view.html',todo=todo_item)
    elif request.method == 'POST':
        value = request.form['button']
        if value == 'Update':
            #if g.user.username in admin:
            #if todo_item.user.id == g.user.id or g.user.username in admin:
            if todo_item.user.id == g.user.id or g.user.username in admin:
                todo_item = TodoAP.query.get(APID)
                todo_item.PRID = request.form['PRID'].strip()
                todo_item.APDescription  = request.form['APDescription']
                todo_item.ApJiraId = request.form['ApJiraId']
                todo_item.APCreatedDate = request.form['APCreatedDate']
                todo_item.APDueDate = request.form['APDueDate']
                todo_item.APCompletedOn  = request.form['APCompletedOn']
                todo_item.IsApCompleted = request.form['IsApCompleted']
                todo_item.APAssingnedTo  = request.form['APAssingnedTo']
                todo_item.QualityOwner  = request.form['QualityOwner']
                team = request.form['QualityOwner']
                hello = User.query.filter_by(username=team).first()
                todo_item.user_id = hello.id
                db.session.commit()
                return redirect(url_for('apindex'))
            else:
                flash('You are not authorized to edit this item','error')
                return redirect(url_for('show_or_updateap',APID=APID))
        elif value == 'Delete' and g.user.username in admin:
            todo_item = TodoAP.query.get(APID)
            db.session.delete(todo_item)
            db.session.commit()
            return redirect(url_for('apindex'))
        flash('You are not authorized to Delete this item', 'error')
    #flash('You are not authorized to edit this todo item','error')
    return redirect(url_for('show_or_updateap',APID=APID))


def sleeptime(hour,min,sec):
    return hour*3600+min*60+sec



def modifyColumn():
    with UseDatabase(app.config['dbconfig']) as cursor:
        _SQL = "alter table apstatus MODIFY RcaEdaActionType VARCHAR(128)"
        #_SQL = "alter table apstatus ADD column (ApCategory VARCHAR(32),ShouldHaveBeenDetected VARCHAR(32))"
        cursor.execute(_SQL)

@app.route('/syslog_view',methods=['GET','POST'])
@login_required
def syslog_view():
    contents = SystemLog.query.order_by(SystemLog.id.asc()).all()
    titles = ('Index','loginAccount','loginLocation','IP', 'Browser', 'DeviceType','OS','operationType','Time',\
              'prIdorApId','log1','log2','log3','log4','log5','log6','log7','log8','log9','log10','log11','log12')

    return render_template('viewlog.html',
                           the_title='View Log',
                           the_row_titles=titles,
                           the_data=contents,)

def upload_sharepoint(PRID):
    basedir = os.path.abspath(os.path.dirname(__file__))
    SHAREPOINT_COMMON_USER = 'enqing.lei@nokia-sbell.com'
    SHAREPOINT_COMMON_PWD = 'kkcqqgfmmxkgjjmd'
    site_url ='https://nokia.sharepoint.com/sites/LTERCA/'
    ctx_auth = AuthenticationContext(site_url)
    print "Authenticate credentials"
    if ctx_auth.acquire_token_for_user(SHAREPOINT_COMMON_USER, SHAREPOINT_COMMON_PWD):
        ctx = ClientContext(site_url, ctx_auth)
        options = RequestOptions(site_url)
        ctx_auth.authenticate_request(options)
        ctx.ensure_form_digest(options)
        path = get_5whyrca_template(PRID)
        link = upload_binary_file(path,ctx_auth)
        return link
    return ctx_auth.get_last_error()


def get_5whyrca_template(PRID):
    file_dir = os.path.join(basedir, '5WhyRcaTemplate')
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
    templatefilename = 'RCA_EDA_Analysis_Template_LTE_BL-mode.xlsm'
    fullpatholdtemplatefilename = 'D://01-TMT/00-Formal RCA/'+ templatefilename
    #fullpatholdtemplatefilename = os.path.join(file_dir, templatefilename)
    #fullpatholdtemplatefilename = 'D:\\01-TMT\\00-Formal RCA' + templatefilename
    filename = PRID + '.xlsm'
    fullpathnewtemplatefilename = os.path.join(file_dir, filename)
    shutil.copyfile(fullpatholdtemplatefilename, fullpathnewtemplatefilename)
    return fullpathnewtemplatefilename


def upload_binary_file(file_path, ctx_auth):
    """Attempt to upload a binary file to SharePoint"""
    PRID = "CAS-148023-8888"
    base_url = 'https://nokia.sharepoint.com/sites/LTERCA/'
    folder_url = "RcaStore"
    file_name = basename(file_path)
    files_url ="{0}/_api/web/GetFolderByServerRelativeUrl('{1}')/Files/add(url='{2}', overwrite=true)"
    full_url = files_url.format(base_url, folder_url, file_name)
    file_link = "{0}/_api/web/GetFileByServerRelativePath('{1}')/Files/add(url='{2}', overwrite=true)"
    file_full_link = file_link.format(base_url, folder_url, file_name)
    options = RequestOptions(base_url)
    context = ClientContext(base_url, ctx_auth)
    context.request_form_digest()

    options.set_header('Accept', 'application/json; odata=verbose')
    options.set_header('Content-Type', 'application/octet-stream')
    options.set_header('Content-Length', str(os.path.getsize(file_path)))
    options.set_header('X-RequestDigest', context.contextWebInformation.form_digest_value)
    options.method = 'POST'

    with open(file_path, 'rb') as outfile:

        # instead of executing the query directly, we'll try to go around
        # and set the json data explicitly

        context.authenticate_request(options)

        data = requests.post(url=full_url, data=outfile, headers=options.headers, auth=options.auth)

        if data.status_code == 200:
            # our file has uploaded successfully
            # let's return the URL
            base_site = data.json()['d']['Properties']['__deferred']['uri'].split("/sites")[0]
            link = data.json()['d']['LinkingUrl']
            relative_url = data.json()['d']['ServerRelativeUrl'].replace(' ', '%20')
            #LinkingUri
            #return base_site + relative_url
            return link
        else:
            return data.json()['error']

def addApColumn():
    with UseDatabase(app.config['dbconfig1']) as cursor:
        #alter table user MODIFY new1 VARCHAR(1) -->modify field type
        _SQL = "alter table jirarcatable1 ADD column CustomerName  VARCHAR(128)"
        cursor.execute(_SQL)

if __name__ == '__main__':
    #modifyColumn()
    # addApColumn()
    app.run(debug=True,host='0.0.0.0',port=8899)
