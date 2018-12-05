#coding:utf-8
#!/usr/bin/env python
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

from datetime import datetime
from flask import Flask,session, request, flash, url_for, redirect, render_template, abort ,g,send_from_directory,jsonify
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from wtforms import StringField, PasswordField, BooleanField, SubmitField
import mysql.connector
from email.mime.text import MIMEText 
from email.header import Header
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.utils import parseaddr,formataddr
import smtplib,time,os

from httplib2 import socks
import socket
from jira.client import JIRA
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

from office365.sharepoint.file_creation_information import FileCreationInformation
import codecs
from os.path import basename
import shutil

from rcatrackingconfig import addr_dict,addr_dict1,getjira,JiraRequest

"""
PROXY_TYPE_HTTP=3
socks.setdefaultproxy(3,"10.144.1.10",8080)
socket.socket = socks.socksocket
"""

basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root@localhost:3306/fddrca?charset=utf8'
app.config['SQLALCHEMY_COMMIT_ON_TEARDOWN'] = True
app.config['SQLALCHEMY_ECHO'] = False
app.config['SECRET_KEY'] = 'secret_key'
app.config['DEBUG'] = True
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False


db = SQLAlchemy(app)


class User(db.Model):
    __tablename__ = "users"
    id = db.Column('user_id',db.Integer , primary_key=True)
    username = db.Column('username', db.String(20), unique=True , index=True)
    password = db.Column('password' , db.String(250))
    email = db.Column('email',db.String(50),unique=True , index=True)
    registered_on = db.Column('registered_on' , db.DateTime)
    todos = db.relationship('Todo' , backref='user',lazy='select')
    todolongcycletimercas = db.relationship('TodoLongCycleTimeRCA' , backref='user',lazy='select')
    def __init__(self , username ,password , email):
        self.username = username
        self.set_password(password)
        self.email = email
        self.registered_on = datetime.utcnow()

    def set_password(self , password):
        self.password = generate_password_hash(password)

    def check_password(self , password):
        return check_password_hash(self.password , password)

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

class TodoMember(db.Model):
    __tablename__ = 'teammembers'
    ID = db.Column('ID', db.Integer, primary_key=True)
    emailName = db.Column(db.String(64))
    lineManager = db.Column(db.String(32))

    def __init__(self, ID,emailName,lineManager):
        self.ID = ID
        self.emailName = emailName
        self.lineManager = lineManager
    """    
    conn=mysql.connector.connect(host='localhost',user='root',passwd='',port=3306)
    cur=conn.cursor()
    cur.execute('create database if not exists fddrca')
    conn.commit()
    cur.close()
    conn.close()
    """
db.create_all()

class Todo(db.Model):
    __tablename__ = 'rcastatus'
    PRID = db.Column('PRID', db.String(64), primary_key=True)
    PRTitle = db.Column(db.String(1024))
    PRReportedDate = db.Column(db.String(64))
    PRClosedDate=db.Column(db.String(64))
    PROpenDays=db.Column(db.Integer)
    PRRcaCompleteDate = db.Column(db.String(64))
    PRRelease = db.Column(db.String(128))
    PRAttached = db.Column(db.String(128))
    IsLongCycleTime = db.Column(db.String(32))
    IsCatM = db.Column(db.String(32))
    IsRcaCompleted = db.Column(db.String(32)) 
    NoNeedDoRCAReason = db.Column(db.String(64))
    RootCauseCategory=db.Column(db.String(1024))
    FunctionArea = db.Column(db.String(1024))
    CodeDeficiencyDescription = db.Column(db.String(1024))
    CorrectionDescription=db.Column(db.String(1024))
    RootCause = db.Column(db.String(1024))
    IntroducedBy = db.Column(db.String(128))
    Handler = db.Column(db.String(64))

    # New added field for JIRA deployment
    LteCategory=db.Column(db.String(32))
    CustomerOrInternal = db.Column(db.String(32))
    JiraFunctionArea=db.Column(db.String(32))
    TriggerScenarioCategory = db.Column(db.String(128))
    FirstFaultEcapePhase=db.Column(db.String(32))
    FaultIntroducedRelease = db.Column(db.String(256))
    TechnicalRootCause = db.Column(db.String(1024))
    TeamAssessor = db.Column(db.String(64))
    EdaCause = db.Column(db.String(1024))
    RcaRootCause5WhyAnalysis = db.Column(db.String(2048))
    JiraRcaBeReqested = db.Column(db.String(32))
    JiraIssueStatus = db.Column(db.String(32))
    JiraIssueAssignee = db.Column(db.String(128))
    JiraRcaPreparedQualityRating = db.Column(db.Integer)
    JiraRcaDeliveryOnTimeRating = db.Column(db.Integer)
    RcaSubtaskJiraId = db.Column(db.String(32))
    # End of new added field for JIRA

    #rca5whys = db.relationship('Rca5Why' , backref='todo',lazy='select')
    user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'))

    def __init__(self, PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,PRRelease,PRAttached,IsLongCycleTime,\
                 IsCatM,IsRcaCompleted,NoNeedDoRCAReason,RootCauseCategory,FunctionArea,CodeDeficiencyDescription,\
		 CorrectionDescription,RootCause,IntroducedBy,Handler,LteCategory,CustomerOrInternal,JiraFunctionArea,TriggerScenarioCategory, \
                 FirstFaultEcapePhase,FaultIntroducedRelease,TechnicalRootCause,TeamAssessor,EdaCause,RcaRootCause5WhyAnalysis, \
                 JiraRcaBeReqested,JiraIssueStatus,JiraIssueAssignee,JiraRcaPreparedQualityRating,JiraRcaDeliveryOnTimeRating,RcaSubtaskJiraId):
        self.PRID = PRID
        self.PRTitle = PRTitle
        self.PRReportedDate = PRReportedDate
        self.PRClosedDate = PRClosedDate
        self.PROpenDays = PROpenDays
        self.PRRcaCompleteDate = PRRcaCompleteDate
        self.PRRelease = PRRelease
        self.PRAttached = PRAttached
        self.IsLongCycleTime = IsLongCycleTime
        self.IsCatM = IsCatM
        self.IsRcaCompleted = IsRcaCompleted
        self.NoNeedDoRCAReason = NoNeedDoRCAReason
        self.RootCauseCategory = RootCauseCategory
        self.FunctionArea = FunctionArea
        self.CodeDeficiencyDescription = CodeDeficiencyDescription
        self.CorrectionDescription = CorrectionDescription
        self.RootCause = RootCause        
        self.IntroducedBy = IntroducedBy
        self.Handler = Handler

        # New added for JIRA
        self.LteCategory = LteCategory
        self.CustomerOrInternal = CustomerOrInternal
        self.JiraFunctionArea = JiraFunctionArea
        self.TriggerScenarioCategory = TriggerScenarioCategory
        self.FirstFaultEcapePhase = FirstFaultEcapePhase
        self.FaultIntroducedRelease = FaultIntroducedRelease
        self.TechnicalRootCause = TechnicalRootCause
        self.TeamAssessor = TeamAssessor
        self.EdaCause = EdaCause
        self.RcaRootCause5WhyAnalysis = RcaRootCause5WhyAnalysis
        self.JiraRcaBeReqested = JiraRcaBeReqested
        self.JiraIssueStatus = JiraIssueStatus
        self.JiraIssueAssignee = JiraIssueAssignee
        self.JiraRcaPreparedQualityRating = JiraRcaPreparedQualityRating
        self.JiraRcaDeliveryOnTimeRating = JiraRcaDeliveryOnTimeRating
        self.RcaSubtaskJiraId = RcaSubtaskJiraId
        #End of JIRA

class TodoLongCycleTimeRCA(db.Model):
    __tablename__ = 'longcycletimercastatus'
    PRID = db.Column('PRID', db.String(64), primary_key=True)
    PRTitle = db.Column(db.String(1024))
    PRReportedDate = db.Column(db.String(64))
    PRClosedDate=db.Column(db.String(64))
    PROpenDays=db.Column(db.Integer)
    PRRcaCompleteDate = db.Column(db.String(64))
    IsLongCycleTime = db.Column(db.String(32))
    IsCatM = db.Column(db.String(32))
    LongCycleTimeRcaIsCompleted = db.Column(db.String(32)) 
    LongCycleTimeRootCause = db.Column(db.String(1024))
    NoNeedDoRCAReason = db.Column(db.String(64))
    Handler = db.Column(db.String(64))
    user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'))

    def __init__(self, PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,IsLongCycleTime,\
                 IsCatM,LongCycleTimeRcaIsCompleted,LongCycleTimeRootCause,NoNeedDoRCAReason,Handler):
        self.PRID = PRID
        self.PRTitle = PRTitle
        self.PRReportedDate = PRReportedDate
        self.PRClosedDate = PRClosedDate
        self.PRRcaCompleteDate = PRRcaCompleteDate
        self.PROpenDays = PROpenDays
        self.IsLongCycleTime = IsLongCycleTime
        self.IsCatM = IsCatM
        self.LongCycleTimeRcaIsCompleted = LongCycleTimeRcaIsCompleted
        self.LongCycleTimeRootCause = LongCycleTimeRootCause
        self.NoNeedDoRCAReason = NoNeedDoRCAReason
        self.Handler = Handler


class TodoJiraRcaPreparedQualityRating(db.Model):
    __tablename__ = 'jirarcaqualityrating'
    id = db.Column('rating_id', db.Integer, primary_key=True)
    PRID = db.Column(db.String(64))
    RatingValue = db.Column(db.Integer)
    RatingComments = db.Column(db.String(128))
    """
    Facts collection prepared is not good
    Technical Analysis prepared is not good
    RCA analysis prepared is not good
    EDA analysis prepared is not good
    Two or more item prepared is not good
    Three or more item prepared is not good
    All of the analysis prepared is not good


    """
        
db.create_all()
"""
app.config['dbconfig'] = {'host': '127.0.0.1',
                          'user': 'root',
                          'password': '',
                          'database': 'fddrca', }
						  

mysql://10.68.184.123:8080/jupiter4?autoReconnect=true

Database name: jupiter4, Table names: t_boxi_closed_pronto_daily AND t_boxi_new_pronto_daily

User/pwd: root/jupiter111
"""		  
app.config['dbconfig'] = {'host': '10.68.184.123',
                          'port':8080,
                          'user': 'root',
                          'password': 'jupiter111',
                          'database': 'jupiter4', }						  

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


def insert_item(team,PRID,PRTitle,PRReportedDate,PRClosedDate,PRRelease,PRAttached,IntroducedBy,TeamAssessor,jira):
    PROpenDays=daysBetweenDate(PRReportedDate,PRClosedDate)
    PRRcaCompleteDate =''
    if daysBetweenDate(PRReportedDate,PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'
    
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

    IntroducedBy = IntroducedBy
    Handler = team
    #todo_item = Todo.query.get(PRID)

    LteCategory = 'FDD'
    JiraFunctionArea = ''
    TriggerScenarioCategory = ''
    FirstFaultEcapePhase = ''
    FaultIntroducedRelease = ''
    TechnicalRootCause = ''
    TeamAssessor = TeamAssessor
    EdaCause = ''
    RcaRootCause5WhyAnalysis = ''
    status,assignee,RcaSubtaskJiraId = JiraRequest(jira, PRID)
    if status is False:
        CustomerOrInternal = 'No'
        JiraRcaBeReqested = 'No'
        JiraIssueStatus = ''
        JiraIssueAssignee =''
        JiraRcaPreparedQualityRating = 10
        JiraRcaDeliveryOnTimeRating = 10
        RcaSubtaskJiraId =''
    else:
        CustomerOrInternal = 'Yes'
        JiraRcaBeReqested = 'Yes'
        JiraIssueAssignee = assignee
        JiraIssueStatus = status
        JiraRcaPreparedQualityRating = 10
        JiraRcaDeliveryOnTimeRating = 10
        RcaSubtaskJiraId =RcaSubtaskJiraId
    registered_user = Todo.query.filter_by(PRID=PRID).all()
    if len(registered_user) ==0:
        todo = Todo(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,PRRelease,PRAttached,IsLongCycleTime,\
                     IsCatM,IsRcaCompleted,NoNeedDoRCAReason,RootCauseCategory,FunctionArea,CodeDeficiencyDescription,\
             CorrectionDescription,RootCause,IntroducedBy,Handler,LteCategory,CustomerOrInternal,JiraFunctionArea,TriggerScenarioCategory, \
                 FirstFaultEcapePhase,FaultIntroducedRelease,TechnicalRootCause,TeamAssessor,EdaCause,RcaRootCause5WhyAnalysis, \
                 JiraRcaBeReqested,JiraIssueStatus,JiraIssueAssignee,JiraRcaPreparedQualityRating,JiraRcaDeliveryOnTimeRating, \
                    RcaSubtaskJiraId)

        hello = User.query.filter_by(username=team).first()
        todo.user_id=hello.id
        print("todo.user_id=hello.user_id=%s"%hello.id)
        db.session.add(todo)
        db.session.commit()
    else:
        #todo_item = Todo.query.get(PRID)
        todo_item.IntroducedBy = IntroducedBy
        todo_item.Handler = Handler
        todo_item.CustomerOrInternal = CustomerOrInternal
        todo_item.JiraRcaBeReqested = JiraRcaBeReqested
        todo_item.LteCategory = LteCategory
        todo_item.JiraRcaBeReqested = JiraRcaBeReqested
        todo_item.TeamAssessor = TeamAssessor
        todo_item.JiraIssueStatus =JiraIssueStatus
        todo_item.JiraIssueAssignee = JiraIssueAssignee
        #todo_item.JiraIssueAssignee = JiraIssueAssignee
        todo_item.RcaSubtaskJiraId = RcaSubtaskJiraId

        count = TodoJiraRcaPreparedQualityRating.query.filter_by(PRID=PRID).count()
        values = TodoJiraRcaPreparedQualityRating.query.filter_by(PRID=PRID).all()
        sum = 0

        if count != 0:
            for item in values:
                sum = sum + item.RatingValue
            todo_item.JiraRcaPreparedQualityRating = sum/count
        else:
            todo_item.JiraRcaPreparedQualityRating = 10

        current_time = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        dayspast = daysBetweenDate(PRClosedDate, current_time)

        with UseDatabaseDict(app.config['dbconfig']) as cursor:
            _SQL = "select * from user_info where displayName = '"+JiraIssueAssignee+"'"
            cursor.execute(_SQL)
            items = cursor.fetchall()
        if len(items) !=0:
            if JiraIssueStatus in ['Open', 'Reopened', 'In Progress'] and JiraRcaBeReqested == 'Yes':
                todo_item.IsRcaCompleted = 'No'
                if dayspast <= 14:
                    todo_item.JiraRcaDeliveryOnTimeRating = 10
                else:
                    todo_item.JiraRcaDeliveryOnTimeRating = 24 - dayspast
            elif JiraIssueStatus in ['Closed', 'Resolved']:
                todo_item.IsRcaCompleted = 'Yes'
        db.session.commit()
        print ("registered_user.PRTitle=%s"%PRID)

    registered_user = TodoLongCycleTimeRCA.query.filter_by(PRID=PRID).all()
    if len(registered_user) ==0 and IsLongCycleTime=='Yes':
        print 'OK#################################################'
        todo = TodoLongCycleTimeRCA(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,IsLongCycleTime,\
                     IsCatM,LongCycleTimeRcaIsCompleted,LongCycleTimeRootCause,NoNeedDoRCAReason,Handler)

        hello = User.query.filter_by(username=team).first()
        todo.user_id=hello.id
        print("todo.user_id=hello.user_id=%s"%hello.id)
        db.session.add(todo)
        db.session.commit()
    elif IsLongCycleTime == 'Yes':
        todo_item = TodoLongCycleTimeRCA.query.get(PRID)
        todo_item.Handler = Handler
        hello = User.query.filter_by(username=team).first()
        todo_item.user_id=hello.id
        db.session.commit()
        print ("registered_user.PRTitle=%s"%PRID)

def reConnectProntoDb():
    _conn_status = True
    _conn_retry_count = 0
    while _conn_status:
        try:
            print 'Connecting Pronto FDD Db...'
            with UseDatabaseDict(app.config['dbconfig']) as cursor:
                _conn_status = False
                # _SQL = "select * from t_boxi_closed_pronto_daily where PRGroupIC='LTE_DEVCAHZ_CHZ_UPMACPS'"
                _SQL = "SELECT DISTINCT prcls.PRID,prcls.PRTitle, prcls.PRGroupIC, prcls.PRRelease, prcls.PRReportedDate, prcls.PRState, \
                    date_format(prcls.ClosedEnter, '%Y-%m-%d') as PRClosedDate,prcls.PRAttached,\
                    prnew.RespPerson FROM t_boxi_closed_pronto_daily as prcls LEFT JOIN t_boxi_new_pronto_daily as prnew\
                    ON prcls.PRID = prnew.PRID WHERE prcls.PRGroupIC in ('LTE_DEVCAHZ_CHZ_UPMACPS','LTE_DEVPFHZ_CHZ_UPMACPS','NIOYSR2','LTE_DEVHZ1_CHZ_SPEC_UP')\
                     and prcls.PRState='Closed'and prcls.ClosedEnter >='2018-01-01' and prcls.identicalFlag='1' and prcls.isFNR='0'"
                cursor.execute(_SQL)
                contents = cursor.fetchall()
                return contents
        except:
            _conn_retry_count += 1
            print ("_conn_retry_count=%d"%_conn_retry_count)
        print 'ProntoDb connecting Error'
        time.sleep(10)
        continue

def reConnectProntoDbTdd():
    _conn_status = True
    _conn_retry_count = 0
    while _conn_status:
        try:
            print 'Connecting Pronto TDD Db...'
            with UseDatabaseDict(app.config['dbconfig']) as cursor:
                _conn_status = False
                # _SQL = "select * from t_boxi_closed_pronto_daily where PRGroupIC='LTE_DEVCAHZ_CHZ_UPMACPS'"
                _SQL = "SELECT DISTINCT prcls.PRID,prcls.PRTitle, prcls.PRGroupIC, prcls.PRRelease, prcls.PRReportedDate, prcls.PRState, \
                    date_format(prcls.ClosedEnter, '%Y-%m-%d') as PRClosedDate,prcls.PRAttached,\
                    prnew.RespPerson FROM t_boxi_closed_pronto_daily as prcls LEFT JOIN t_boxi_new_pronto_daily as prnew\
                    ON prcls.PRID = prnew.PRID WHERE prcls.PRGroupIC in ('NIHZSMAC','NIHZYSUP','NIHZSMAC_RCK')\
                     and prcls.PRState='Closed'and prcls.ClosedEnter >='2018-01-01' and prcls.identicalFlag='1' and prcls.isFNR='0'"
                cursor.execute(_SQL)
                contents = cursor.fetchall()
                return contents
        except:
            _conn_retry_count += 1
            print ("_conn_retry_count=%d"%_conn_retry_count)
        print 'ProntoDb connecting Error'
        time.sleep(10)
        continue

teams = ['chenlong','xiezhen','yangjinyong','zhangyijie','lanshenghai','liumingjing','lizhongyuan','caizhichao','hujun','wangli']

def update_from_tbox(jira):
    #teams=['chenlong','xiezhen','yangjinyong','zhangyijie','lanshenghai','liumingjing','lizhongyuan','caizhichao','hujun','wuyuanxing']
    contents = reConnectProntoDb()
    count=len(contents)
    member_map = {}
    for linename in teams:
        memberlist = []
        members = TodoMember.query.filter_by(lineManager=linename).all()
        for member in members:
            email=member.emailName
            email=email.encode('utf-8').strip()
            memberlist.append(email)
        member_map.setdefault(linename,memberlist)
        print 'before hello'
        for content in contents:
            PRID = content['PRID'].encode('utf-8').strip()
            todo_item = Todo.query.get(PRID)
            """
            if todo_item:
                continue
            """
            RespPerson = []
            resp= content['RespPerson'].encode('utf-8').strip()
            RespPerson=resp_members(resp)
            #RespPerson.append(resp)
            a=list(set(member_map[linename]).intersection(set(RespPerson)))
            memberline=set(RespPerson)
            retA = [val for val in RespPerson if val in member_map[linename]]
            #RespPerson = set(RespPerson)
            #retA = [i for i in listA if i in listB]
            MemberOfLine= set(member_map[linename])
            s=MemberOfLine.intersection(memberline)
            if MemberOfLine.intersection(memberline):
                IntroducedBy = MemberOfLine.intersection(memberline)
                print 'hello'
                PRID = content['PRID'].encode('utf-8').strip()
                print PRID
                PRTitle = content['PRTitle'].encode('utf-8').strip()
                PRReportedDate = str(content['PRReportedDate'])
                PRClosedDate = content['PRClosedDate'].encode('utf-8').strip()
                PRRelease = content['PRRelease'].encode('utf-8').strip()
                PRAttached = content['PRAttached'].encode('utf-8').strip()
                IntroducedBy = IntroducedBy

                if content['PRGroupIC'] == 'NIOYSR2' or content['PRGroupIC'] == 'LTE_DEVHZ1_CHZ_SPEC_UP':
                    TeamAssessor = 'jun-julian.hu@nokia-sbell.com'
                    insert_item('wangli', PRID, PRTitle, PRReportedDate, PRClosedDate, PRRelease, PRAttached,
                                IntroducedBy,TeamAssessor,jira)
                else:
                    TeamAssessor = addr_dict[linename]['fc']
                    insert_item(linename, PRID, PRTitle, PRReportedDate, PRClosedDate, PRRelease, PRAttached,IntroducedBy,\
                                TeamAssessor, jira)

def insert_item_jira(team,PRID, PRTitle, PRReportedDate, PRClosedDate, PRRelease, PRAttached, IntroducedBy,\
                     LteCategory,JiraRcaBeReqested,TeamAssessor,JiraIssueStatus,JiraIssueAssignee,RcaSubtaskJiraId):

    PROpenDays = daysBetweenDate(PRReportedDate, PRClosedDate)
    PRRcaCompleteDate = ''
    if daysBetweenDate(PRReportedDate, PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'

    IsCatM = ''
    IsRcaCompleted = 'No'
    LongCycleTimeRcaIsCompleted = 'No'
    NoNeedDoRCAReason = ''
    RootCauseCategory = ''
    FunctionArea = ''

    CodeDeficiencyDescription = ''
    CorrectionDescription = ''
    RootCause = ''
    LongCycleTimeRootCause = ''

    IntroducedBy = IntroducedBy
    Handler = team

    LteCategory = LteCategory
    CustomerOrInternal = 'Yes'
    JiraFunctionArea = ''
    TriggerScenarioCategory = ''
    FirstFaultEcapePhase = ''
    FaultIntroducedRelease = ''
    TechnicalRootCause = ''
    TeamAssessor = TeamAssessor
    EdaCause = ''
    RcaRootCause5WhyAnalysis = ''
    JiraRcaBeReqested = JiraRcaBeReqested
    JiraIssueStatus = JiraIssueStatus
    JiraIssueAssignee = JiraIssueAssignee
    JiraRcaPreparedQualityRating = 10
    JiraRcaDeliveryOnTimeRating = 10
    RcaSubtaskJiraId = RcaSubtaskJiraId
    with UseDatabaseDict(app.config['dbconfig']) as cursor:
        #_SQL = "select * from user_info where displayName = '" + JiraIssueAssignee + "'"
        _SQL = "select * from user_info where email = '" + TeamAssessor + "'"
        cursor.execute(_SQL)
        items = cursor.fetchall()
    if len(items) != 0:
        emailname = items[0]['email'].encode('utf-8')
        lmEmail = items[0]['lmEmail'].encode('utf-8')
        if lmEmail in addr_dict1.keys():
            team = addr_dict1[lmEmail]
            #TeamAssessor = emailname
            Handler = team
    #hello = User.query.filter_by(username=Handler).first()
    registered_user = Todo.query.filter_by(PRID=PRID).all()
    if len(registered_user) == 0:
        todo = Todo(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,PRRelease,PRAttached,IsLongCycleTime,\
                 IsCatM,IsRcaCompleted,NoNeedDoRCAReason,RootCauseCategory,FunctionArea,CodeDeficiencyDescription,\
		 CorrectionDescription,RootCause,IntroducedBy,Handler,LteCategory,CustomerOrInternal,JiraFunctionArea,TriggerScenarioCategory, \
                 FirstFaultEcapePhase,FaultIntroducedRelease,TechnicalRootCause,TeamAssessor,EdaCause,RcaRootCause5WhyAnalysis, \
                 JiraRcaBeReqested,JiraIssueStatus,JiraIssueAssignee,JiraRcaPreparedQualityRating,JiraRcaDeliveryOnTimeRating, \
                    RcaSubtaskJiraId)
        # g.user=Todo.query.get(team)
        # todo.user = g.user
        # print("todo.user = g.user=%s"%todo.user)
        hello = User.query.filter_by(username=Handler).first()
        todo.user_id = hello.id
        print("todo.user_id=hello.user_id=%s" % hello.id)
        db.session.add(todo)
        db.session.commit()
    else:
        #todo_item = Todo.query.get(PRID)
        todo_item.IntroducedBy = IntroducedBy
        todo_item.Handler = Handler
        todo_item.LteCategory = LteCategory
        todo_item.CustomerOrInternal = CustomerOrInternal
        todo_item.JiraRcaBeReqested = JiraRcaBeReqested
        todo_item.JiraIssueStatus = JiraIssueStatus
        todo_item.TeamAssessor = TeamAssessor
        todo_item.JiraIssueAssignee = JiraIssueAssignee

        count = TodoJiraRcaPreparedQualityRating.query.filter_by(PRID=PRID).count()
        values = TodoJiraRcaPreparedQualityRating.query.filter_by(PRID=PRID).all()
        sum = 0

        if count != 0:
            for item in values:
                sum = sum + item.RatingValue
            todo_item.JiraRcaPreparedQualityRating = sum/count
        else:
            todo_item.JiraRcaPreparedQualityRating = 10

        current_time = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        dayspast = daysBetweenDate(PRClosedDate, current_time)
        with UseDatabaseDict(app.config['dbconfig']) as cursor:
            SQL = "select * from user_info where displayName = '"+JiraIssueAssignee+"'"
            #_SQL = "select * from user_info where email = '" + TeamAssessor + "'"
            cursor.execute(_SQL)
            items = cursor.fetchall()
            if len(items) !=0:
                if JiraIssueStatus in ['Open', 'Reopened', 'In Progress'] and JiraRcaBeReqested == 'Yes':
                    todo_item.IsRcaCompleted = 'No'
                    if dayspast <= 14:
                        todo_item.JiraRcaDeliveryOnTimeRating = 10
                    else:
                        todo_item.JiraRcaDeliveryOnTimeRating = 24 - dayspast
                elif JiraIssueStatus in ['Closed', 'Resolved']:
                    todo_item.IsRcaCompleted = 'Yes'
        """
            emailname = items[0]['email'].encode('utf-8')
            lmEmail = items[0]['lmEmail'].encode('utf-8')
            if lmEmail in addr_dict1.keys():
                team = addr_dict1[lmEmail]
                #todo_item.TeamAssessor = emailname
                todo_item.Handler = team
        #else:
            #todo_item.IsRcaCompleted = 'Yes'
        hello = User.query.filter_by(username=team).first()
        todo_item.user_id = hello.id
        """
        db.session.commit()
        print ("registered_user.PRTitle=%s" % PRID)

    registered_user = TodoLongCycleTimeRCA.query.filter_by(PRID=PRID).all()
    if len(registered_user) == 0 and IsLongCycleTime == 'Yes':
        print 'OK#################################################'
        todo = TodoLongCycleTimeRCA(PRID, PRTitle, PRReportedDate, PRClosedDate, PROpenDays, PRRcaCompleteDate,
                                    IsLongCycleTime, \
                                    IsCatM, LongCycleTimeRcaIsCompleted, LongCycleTimeRootCause, NoNeedDoRCAReason,
                                    Handler)

        hello = User.query.filter_by(username=team).first()
        todo.user_id = hello.id
        print("todo.user_id=hello.user_id=%s" % hello.id)
        db.session.add(todo)
        db.session.commit()
    elif IsLongCycleTime == 'Yes':
        todo_item = TodoLongCycleTimeRCA.query.get(PRID)
        todo_item.Handler = Handler
        hello = User.query.filter_by(username=team).first()
        todo_item.user_id = hello.id
        db.session.commit()
        print ("registered_user.PRTitle=%s" % PRID)

def update_from_tbox_fdd(jira):
    member_map={}
    contents = reConnectProntoDb()
    for content in contents:
        PRID = content['PRID'].encode('utf-8').strip()
        todo_item = Todo.query.get(PRID)
        # If already in WebApp table, then no need update anymore, keep the manual change valid.
        if todo_item:
            continue
        """
        status,assignee,RcaSubtaskJiraId = JiraRequest(jira,PRID)
        if status is False:
            continue
        JiraIssueStatus = status
        JiraIssueAssignee =assignee
        RcaSubtaskJiraId = RcaSubtaskJiraId
        """
        #JiraIssueStatus = GetJiraIssueStatus(issue)
        for linename in teams:
            email_name = addr_dict[linename]['email']
            with UseDatabaseDict(app.config['dbconfig']) as cursor:
                _SQL = "select * from user_info where lmEmail= '"+email_name+"'"
                cursor.execute(_SQL)
                members = cursor.fetchall()
            memberlist = []
            #members = TodoMember.query.filter_by(lineManager=linename).all()
            for member in members:
                displayname=member['displayName']
                email=displayname.encode('utf-8').strip()
                memberlist.append(email)
            member_map.setdefault(linename,memberlist)
            print 'before hello'
            resp= content['RespPerson'].encode('utf-8').strip()
            RespPerson=resp_members(resp)
            a=list(set(member_map[linename]).intersection(set(RespPerson)))
            memberline=set(RespPerson)
            retA = [val for val in RespPerson if val in member_map[linename]]
            MemberOfLine= set(member_map[linename])
            s=MemberOfLine.intersection(memberline)
            if MemberOfLine.intersection(memberline):
                IntroducedBy = MemberOfLine.intersection(memberline)
                print 'hello'
                PRID = content['PRID'].encode('utf-8').strip()
                print PRID
                PRTitle = content['PRTitle'].encode('utf-8').strip()
                PRReportedDate = str(content['PRReportedDate'])
                PRClosedDate = content['PRClosedDate'].encode('utf-8').strip()
                PRRelease = content['PRRelease'].encode('utf-8').strip()
                PRAttached = content['PRAttached'].encode('utf-8').strip()
                IntroducedBy = IntroducedBy
                #JiraRcaBeReqested = 'Yes'
                #LteCategory = 'TDD'

                if content['PRGroupIC'] == 'NIOYSR2' or content['PRGroupIC'] == 'LTE_DEVHZ1_CHZ_SPEC_UP':
                    TeamAssessor = 'jun-julian.hu@nokia-sbell.com'
                    insert_item('wangli', PRID, PRTitle, PRReportedDate, PRClosedDate, PRRelease, PRAttached,
                                IntroducedBy,TeamAssessor,jira)
                else:
                    TeamAssessor = addr_dict[linename]['fc']
                    insert_item(linename, PRID, PRTitle, PRReportedDate, PRClosedDate, PRRelease, PRAttached,IntroducedBy,\
                                TeamAssessor, jira)
									 
def update_from_tbox_tdd(jira):
    member_map={}
    contents = reConnectProntoDbTdd()
    for content in contents:
        PRID = content['PRID'].encode('utf-8').strip()
        Jira5WhyRcaRequestOnly(jira, PRID)
        if PRID == 'PR353980':
            print 'why'
            continue
        todo_item = Todo.query.get(PRID)
        # If already in WebApp table, then no need update anymore, keep the manual change valid.
        if todo_item:
            continue
        status,assignee,RcaSubtaskJiraId = JiraRequest(jira,PRID)
        #TDD only check Customer PR,Status==False means, this is not a PR from JIRA Customer
        if status is False:
            continue
        else:
            JiraIssueStatus = status
            JiraIssueAssignee =assignee
            RcaSubtaskJiraId = RcaSubtaskJiraId
            #JiraIssueStatus = GetJiraIssueStatus(issue)

        # indicate whether exact line can be found, t-box bug workaround.
        ownerFlag = False

        for linename in teams:
            email_name = addr_dict[linename]['email']
            with UseDatabaseDict(app.config['dbconfig']) as cursor:
                _SQL = "select * from user_info where lmEmail= '"+email_name+"'"
                cursor.execute(_SQL)
                members = cursor.fetchall()
            memberlist = []
            #members = TodoMember.query.filter_by(lineManager=linename).all()
            for member in members:
                displayname=member['displayName']
                email=displayname.encode('utf-8').strip()
                memberlist.append(email)
            member_map.setdefault(linename,memberlist)
            print 'before hello'
            resp= content['RespPerson'].encode('utf-8').strip()
            RespPerson=resp_members(resp)
            a=list(set(member_map[linename]).intersection(set(RespPerson)))
            memberline=set(RespPerson)
            retA = [val for val in RespPerson if val in member_map[linename]]
            MemberOfLine= set(member_map[linename])
            s=MemberOfLine.intersection(memberline)
            if MemberOfLine.intersection(memberline):
                ownerFlag = True
                IntroducedBy = MemberOfLine.intersection(memberline)
                TeamAssessor = addr_dict[linename]['fc']
                break
        if ownerFlag == False:
            IntroducedBy = resp
            linename = 'liumingjing'
            TeamAssessor = addr_dict[linename]['fc']
        print 'hello'
        PRID = content['PRID'].encode('utf-8').strip()
        print PRID
        PRTitle = content['PRTitle'].encode('utf-8').strip()
        PRReportedDate = str(content['PRReportedDate'])
        PRClosedDate = content['PRClosedDate'].encode('utf-8').strip()
        PRRelease = content['PRRelease'].encode('utf-8').strip()
        PRAttached = content['PRAttached'].encode('utf-8').strip()
        IntroducedBy = IntroducedBy
        JiraRcaBeReqested = 'Yes'
        LteCategory = 'TDD'

        if content['PRGroupIC'] == 'NIHZYSUP':
            TeamAssessor = 'yuan_xing.wu@nokia-sbell.com'
            insert_item_jira('wangli', PRID, PRTitle, PRReportedDate, PRClosedDate, PRRelease, PRAttached,
                        IntroducedBy,LteCategory, JiraRcaBeReqested,TeamAssessor,JiraIssueStatus,JiraIssueAssignee, \
                             RcaSubtaskJiraId)
        else:
            #TeamAssessor = addr_dict[linename]['fc']
            insert_item_jira(linename, PRID, PRTitle, PRReportedDate, PRClosedDate, PRRelease, PRAttached,IntroducedBy, \
                             LteCategory, JiraRcaBeReqested,TeamAssessor,JiraIssueStatus,JiraIssueAssignee, \
                             RcaSubtaskJiraId)


def insert_item_common(team,PRID,PRTitle,PRReportedDate,PRClosedDate,PRRelease,PRAttached,IntroducedBy,\
                     LteCategory,JiraRcaBeReqested,TeamAssessor,JiraIssueStatus,JiraIssueAssignee,RcaSubtaskJiraId):

    PROpenDays = daysBetweenDate(PRReportedDate, PRClosedDate)
    PRRcaCompleteDate = ''
    if daysBetweenDate(PRReportedDate, PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'

    IsCatM = ''
    IsRcaCompleted = 'No'
    LongCycleTimeRcaIsCompleted = 'No'
    NoNeedDoRCAReason = ''
    RootCauseCategory = ''
    FunctionArea = ''

    CodeDeficiencyDescription = ''
    CorrectionDescription = ''
    RootCause = ''
    LongCycleTimeRootCause = ''

    IntroducedBy = IntroducedBy
    Handler = team

    LteCategory = LteCategory
    CustomerOrInternal = 'Yes'
    JiraFunctionArea = ''
    TriggerScenarioCategory = ''
    FirstFaultEcapePhase = ''
    FaultIntroducedRelease = ''
    TechnicalRootCause = ''
    TeamAssessor = TeamAssessor
    EdaCause = ''
    RcaRootCause5WhyAnalysis = ''
    JiraRcaBeReqested = JiraRcaBeReqested
    JiraIssueStatus = JiraIssueStatus
    JiraIssueAssignee = JiraIssueAssignee
    JiraRcaPreparedQualityRating = 10
    JiraRcaDeliveryOnTimeRating = 10
    RcaSubtaskJiraId = RcaSubtaskJiraId

    todo = Todo(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,PRRelease,PRAttached,IsLongCycleTime,\
             IsCatM,IsRcaCompleted,NoNeedDoRCAReason,RootCauseCategory,FunctionArea,CodeDeficiencyDescription,\
     CorrectionDescription,RootCause,IntroducedBy,Handler,LteCategory,CustomerOrInternal,JiraFunctionArea,TriggerScenarioCategory, \
             FirstFaultEcapePhase,FaultIntroducedRelease,TechnicalRootCause,TeamAssessor,EdaCause,RcaRootCause5WhyAnalysis, \
             JiraRcaBeReqested,JiraIssueStatus,JiraIssueAssignee,JiraRcaPreparedQualityRating,JiraRcaDeliveryOnTimeRating, \
                RcaSubtaskJiraId)

    hello = User.query.filter_by(username=Handler).first()
    todo.user_id = hello.id
    print("todo.user_id=hello.user_id=%s" % hello.id)
    db.session.add(todo)
    db.session.commit()

    registered_user = TodoLongCycleTimeRCA.query.filter_by(PRID=PRID).all()
    if len(registered_user) == 0 and IsLongCycleTime == 'Yes':
        print 'OK#################################################'
        todo = TodoLongCycleTimeRCA(PRID, PRTitle, PRReportedDate, PRClosedDate, PROpenDays, PRRcaCompleteDate,
                                    IsLongCycleTime, \
                                    IsCatM, LongCycleTimeRcaIsCompleted, LongCycleTimeRootCause, NoNeedDoRCAReason,
                                    Handler)

        hello = User.query.filter_by(username=team).first()
        todo.user_id = hello.id
        print("todo.user_id=hello.user_id=%s" % hello.id)
        db.session.add(todo)
        db.session.commit()


def check_assign_issue(jira):
    print 'Check and assign issue...'
    issues = jira.search_issues('project = MNRCA and assignee = qmxh38 and type = "Analysis subtask" and \
                                status not in (Resolved, Closed)')
    if len(issues):
        for issue in issues:
            summary = issue.fields.summary
            PRID =summary.split(' ',2)[2]
            parentissues = jira.search_issues(jql_str='project = MNPRCA AND summary~' + PRID + ' AND type = "5WhyRCA"',maxResults=5)
            for parentissue in parentissues:
                # Upload Template and get the SharePoint Online excel link
                link = upload_sharepoint(PRID)
                prontolink = "https://pronto.inside.nsn.com/pronto/problemReportSearch.html?freeTextdropDownID=prId&searchTopText={0}".format(PRID)
                dict1 = {'customfield_37015': prontolink}
                parentissue.update(dict1)
                analysislink = str(link)
                dict2 = {'customfield_37064': analysislink}
                parentissue.update(dict2)

            """
            if PRID != 'CAS-154221-D7L9':
                print PRID
                continue
            """
            todo_item = Todo.query.get(PRID)
            if todo_item: # PR in the rca table,just do the assignment and update assignee.
                TeamAssessor = todo_item.TeamAssessor
                with UseDatabaseDict(app.config['dbconfig']) as cursor:
                    _SQL = "select * from user_info where email = '" + TeamAssessor + "'"
                    cursor.execute(_SQL)
                    items = cursor.fetchall()
                if len(items) != 0:
                    emailname = items[0]['accountId'].encode('utf-8')
                    jira.assign_issue(issue, emailname)

                    todo_item.JiraIssueAssignee = items[0]['displayName'].encode('utf-8')
                    #issue.update(assignee={'name': emailname})
                    todo_item.JiraRcaBeReqested = 'Yes'
                    todo_item.RcaSubtaskJiraId = issue.key
                    todo_item.JiraIssueStatus = issue.fields.status
                #jira_assign()
                db.session.commit()
            else: # PR not in the RCA table, first add it to RCA table and do the assignment.
                with UseDatabaseDict(app.config['dbconfig']) as cursor:
                    _conn_status = False
                    #PRID ='PR326676'
                    # _SQL = "select * from t_boxi_closed_pronto_daily where PRID='"+PRID+"'"
                    _SQL = "SELECT DISTINCT prcls.PRID,prcls.PRTitle, prcls.Platform,prcls.PRGroupIC, prcls.PRRelease, prcls.PRReportedDate,\
                     prcls.PRState, date_format(prcls.ClosedEnter, '%Y-%m-%d') as PRClosedDate,prcls.PRAttached,\
                    prnew.RespPerson FROM t_boxi_closed_pronto_daily as prcls LEFT JOIN t_boxi_new_pronto_daily as prnew\
                    ON prcls.PRID = prnew.PRID WHERE prcls.PRID='"+PRID+"'"
                    #_SQL = "select * from user_info where email = '" + JiraIssueAssignee + "'"
                    cursor.execute(_SQL)
                    contents = cursor.fetchall()
                    """
                    if not contents:
                        _SQL = "select * from t_boxi_new_pronto_daily where PRID='" + PRID + "'"
                        cursor.execute(_SQL)
                        contents = cursor.fetchall()
                    """
                #if len(contents):
                member_map = {}
                for content in contents:
                    PRID = content['PRID'].encode('utf-8').strip()
                    ownerFlag = False
                    for linename in teams:
                        email_name = addr_dict[linename]['email']
                        with UseDatabaseDict(app.config['dbconfig']) as cursor:
                            _SQL = "select * from user_info where lmEmail= '" + email_name + "'"
                            cursor.execute(_SQL)
                            members = cursor.fetchall()
                        memberlist = []
                        # members = TodoMember.query.filter_by(lineManager=linename).all()
                        for member in members:
                            displayname = member['displayName']
                            email = displayname.encode('utf-8').strip()
                            memberlist.append(email)
                        member_map.setdefault(linename, memberlist)
                        print 'before hello'
                        resp = content['RespPerson'].encode('utf-8').strip()
                        RespPerson = resp_members(resp)
                        a = list(set(member_map[linename]).intersection(set(RespPerson)))
                        memberline = set(RespPerson)
                        retA = [val for val in RespPerson if val in member_map[linename]]
                        MemberOfLine = set(member_map[linename])
                        s = MemberOfLine.intersection(memberline)
                        if MemberOfLine.intersection(memberline):
                            ownerFlag = True
                            IntroducedBy = MemberOfLine.intersection(memberline)
                            TeamAssessor = addr_dict[linename]['fc']
                            team = linename
                            break
                    if ownerFlag == False:
                        IntroducedBy = resp
                        team = 'liumingjing'
                        TeamAssessor = addr_dict[linename]['fc']

                    print 'hello'
                    PRID = content['PRID'].encode('utf-8').strip()
                    print PRID
                    PRTitle = content['PRTitle'].encode('utf-8').strip()
                    PRReportedDate = str(content['PRReportedDate'])
                    PRClosedDate = content['PRClosedDate'].encode('utf-8').strip()
                    #PRClosedDate = '2018-10-11'
                    PRRelease = content['PRRelease'].encode('utf-8').strip()
                    PRAttached = content['PRAttached'].encode('utf-8').strip()
                    IntroducedBy = IntroducedBy
                    LteCategory = content['Platform']
                    if LteCategory == 'TDLTE':
                        LteCategory = 'TDD'
                    if LteCategory =='FDDMACPS':
                        LteCategory ='FDD'
                    JiraRcaBeReqested = 'Yes'
                    JiraIssueStatus = issue.fields.status
                    RcaSubtaskJiraId = issue.key
                    # LteCategory = 'TDD'
                    if content['PRGroupIC'] == 'NIOYSR2' or content['PRGroupIC'] == 'LTE_DEVHZ1_CHZ_SPEC_UP':
                        TeamAssessor = 'jun-julian.hu@nokia-sbell.com'
                        team = 'wangli'
                    elif content['PRGroupIC'] == 'NIHZYSUP':
                        TeamAssessor = 'yuan_xing.wu@nokia-sbell.com'
                        team = 'wangli'
                    with UseDatabaseDict(app.config['dbconfig']) as cursor:
                        _SQL = "select * from user_info where email = '" + TeamAssessor + "'"
                        cursor.execute(_SQL)
                        items = cursor.fetchall()
                    if len(items) != 0:
                        emailname = items[0]['accountId'].encode('utf-8')
                        JiraIssueAssignee = items[0]['displayName'].encode('utf-8')
                        jira.assign_issue(issue,emailname)
                        #issue.update(assignee={'name': emailname})
                        JiraRcaBeReqested = 'Yes'
                        RcaSubtaskJiraId = issue.key
                        JiraIssueStatus = issue.fields.status
                    insert_item_common(team,PRID,PRTitle,PRReportedDate,PRClosedDate,PRRelease,
                                       PRAttached,IntroducedBy,LteCategory,JiraRcaBeReqested,TeamAssessor,\
                                       JiraIssueStatus,JiraIssueAssignee,RcaSubtaskJiraId)


def quality_rating(jira):
    #todos = Todo.query.filter(Todo.JiraRcaBeReqested == 'Yes', ~Todo.JiraIssueStatus.in_(['Closed', 'Resolved'])).order_by(Todo.PRClosedDate.asc()).all()
    todos = Todo.query.filter(Todo.JiraRcaBeReqested == 'Yes',Todo.JiraIssueStatus.in_(['Open','Closed','Resolved','In Progress','Reopened'])).all()
    #todos = Todo.query.filter( ~Todo.JiraIssueStatus.in_(['Closed', 'Resolved'])).order_by(Todo.PRClosedDate.asc()).all()
    for todo in todos:
        PRID = todo.PRID
        todo_item = Todo.query.get(PRID)
        JiraIssueStatus = todo_item.JiraIssueStatus
        JiraIssueId = todo_item.RcaSubtaskJiraId
        if JiraIssueId:
            issue = jira.issue(JiraIssueId)
            JiraIssueStatus = issue.fields.status
            JiraIssueAssignee = str(issue.fields.assignee)
        status, assignee, jiraissuekey = JiraRequest(jira, PRID)
        if status is False:
            continue
        else:
            JiraIssueStatus = status
            JiraIssueAssignee = assignee
            JiraIssueId = jiraissuekey
        todo_item.JiraIssueStatus = JiraIssueStatus
        todo_item.JiraIssueAssignee = JiraIssueAssignee
        todo_item.JiraRcaBeReqested = 'Yes'
        todo_item.RcaSubtaskJiraId = JiraIssueId
        PRClosedDate = todo_item.PRClosedDate
        #JiraIssueAssignee = todo_item.JiraIssueAssignee

        count = TodoJiraRcaPreparedQualityRating.query.filter_by(PRID=PRID).count()
        values = TodoJiraRcaPreparedQualityRating.query.filter_by(PRID=PRID).all()
        sum = 0
        if count != 0:
            for item in values:
                sum = sum + item.RatingValue
            todo_item.JiraRcaPreparedQualityRating = sum / count
        else:
            todo_item.JiraRcaPreparedQualityRating = 10

        current_time = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        dayspast = daysBetweenDate(PRClosedDate, current_time)
        with UseDatabaseDict(app.config['dbconfig']) as cursor:
            _SQL = "select * from user_info where displayName = '"+JiraIssueAssignee+"'"
            cursor.execute(_SQL)
            items = cursor.fetchall()
            if len(items) !=0:
                if JiraIssueStatus in ['Open', 'Reopened', 'In Progress']:
                    todo_item.IsRcaCompleted = 'No'
                    if dayspast <= 14:
                        todo_item.JiraRcaDeliveryOnTimeRating = 10
                    else:
                        todo_item.JiraRcaDeliveryOnTimeRating = 24 - dayspast
                elif JiraIssueStatus in ['Closed', 'Resolved']:
                    todo_item.IsRcaCompleted = 'Yes'
        db.session.commit()

def TodoAP_Synch_up_with_jira(jira):
    todos = TodoAP.query.filter(TodoAP.CustomerAp == 'Yes',TodoAP.IsApCompleted == 'No').all()
    #.in_(['Open','Closed','Resolved','In Progress','Reopened']
    for todo in todos:
        JiraIssueId = todo.ApJiraId
        APID = todo.APID
        todo_item = TodoAP.query.get(APID)
        if JiraIssueId:
            issue = jira.issue(JiraIssueId)
            JiraIssueStatus = str(issue.fields.status)
            #JiraIssueAssignee = str(issue.fields.assignee)
            if JiraIssueStatus in ['Closed', 'Resolved']:
                todo_item.IsApCompleted = 'Yes'
                todo_item.APCompletedOn = time.strftime('%Y-%m-%d', time.localtime(time.time()))
                db.session.commit()


def view_the_pr():
    """Display the contents of the log file as a HTML table."""
    with UseDatabaseDict(app.config['dbconfig']) as cursor:
        _SQL = "select * from t_boxi_closed_pronto_daily where PRGroupIC='LTE_DEVCAHZ_CHZ_UPMACPS'"
        _SQL="SELECT DISTINCT prcls.PRID,prcls.PRTitle, prcls.PRGroupIC, prcls.PRRelease, prcls.PRReportedDate, prcls.PRState, \
                date_format(prcls.ClosedEnter, '%Y-%m-%d') as PRClosedDate,prcls.PRAttached,\
                prnew.RespPerson FROM t_boxi_closed_pronto_daily as prcls LEFT JOIN t_boxi_new_pronto_daily as prnew\
                ON prcls.PRID = prnew.PRID WHERE prcls.PRGroupIC in ('LTE_DEVCAHZ_CHZ_UPMACPS','LTE_DEVPFHZ_CHZ_UPMACPS')\
                 and prcls.PRState='Closed'and prcls.ClosedEnter >='2018-01-01'"
        cursor.execute(_SQL)
        contents = cursor.fetchall() 
        count=len(contents)
        print count 

def resp_members(b):
    members = []
    #b = 'Ma, Cong 2. (NSB - CN/Hangzhou),Ma, Cong 3. (NSB - CN/Hangzhou),'
    a=b.split('),')
    #c=a[0]+')'
    n=len(a)
    if n>1:
        for i in range(n-1):
            c = a[i] + ')'
            members.append(c.encode('utf-8').strip())
        members.append(a[n-1].encode('utf-8').strip())
    else:
        members.append(b.encode('utf-8').strip())
    return members
    print a

def Jira5WhyRcaRequestOnly(jira, PRID):
    JIRA_STATUS = {'Closed', 'Resolved'}
    try:
        issues = jira.search_issues(jql_str='project = MNPRCA AND summary~' + PRID + ' AND type = "5WhyRCA" and assignee in(qmxh38, mingjliu)', maxResults=5)
    except:
        issues = []
    #issues = jira.search_issues(jql_str='project = MNPRCA AND type = "5WhyRCA" and assignee= mingjliu', maxResults=100)
    for issue in issues:
        #issue = jira.issue('MNRCA-14709')
        analysislink = issue.fields.customfield_37064
        if not analysislink:
            link = upload_sharepoint(PRID)
            prontolink="https://pronto.inside.nsn.com/pronto/problemReportSearch.html?freeTextdropDownID=prId&searchTopText={0}".format(PRID)
            dict2 = {'customfield_37015': prontolink}
            issue.update(dict2)
            analysislink = str(link)
            dict1 = {'customfield_37064': analysislink}
            # issue.fields.customfield_37064 = analysislink # Invalid, verified.
            issue.update(dict1)

        subtasks =issue.fields.subtasks
        if not subtasks:
            RcaAnalysisSubtaskKey = issue.key
            # Jira RCA Analysis Subtask
            summaryinfo = 'RCA for '+ PRID
            issueaddforRCAsubtask = {
                'project': {'id': u'41675'},
                'parent': {"key": RcaAnalysisSubtaskKey},
                'issuetype': {'id': u'19800'},
                'summary': summaryinfo,
                'customfield_10464': {'id': u'219575'},  # Case Type
                'assignee': {'name':'qmxh38'}
            }
            newissue = jira.create_issue(issueaddforRCAsubtask)

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
    #fullpatholdtemplatefilename = 'D://01-TMT/00-Formal RCA/'+ templatefilename
    fullpatholdtemplatefilename = os.path.join(file_dir, templatefilename)
    #fullpatholdtemplatefilename = 'D:\\01-TMT\\00-Formal RCA' + templatefilename
    filename = PRID + '.xlsm'
    fullpathnewtemplatefilename = os.path.join(file_dir, filename)
    shutil.copyfile(fullpatholdtemplatefilename, fullpathnewtemplatefilename)
    return fullpathnewtemplatefilename


def upload_binary_file(file_path, ctx_auth):
    """Attempt to upload a binary file to SharePoint"""
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

def sleeptime(day,hour,min,sec):
    return day*24*3600+hour*3600 + min*60 + sec

counting=0
def update():
    global counting
    s1 = '09:00'
    s2 = datetime.now().strftime('%H:%M')
    if s2 == s2:
        counting=counting+1
        print 'Action now!'
        print counting
        jira = getjira()
        update_from_tbox_fdd(jira)
        update_from_tbox_tdd(jira)
        check_assign_issue(jira)
        quality_rating(jira)
        TodoAP_Synch_up_with_jira(jira)
        print 'update completed!!'
        second = sleeptime(0,0,1,10)
        time.sleep(second)

def JIRAtest(jira):
    JiraRequest(jira,'CAS-120133-Q1T3')

if __name__ == "__main__":
    jira = getjira()
    check_assign_issue(jira)
    #quality_rating(jira)
    #update_from_tbox_fdd(jira)
    update_from_tbox_tdd(jira)
    print 'Update from T-Boxi is Ready now!!!!'
    #while 1==1:
        #update()



