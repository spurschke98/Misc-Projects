# Automation of Mail Chimp Data 
## Import Packages
# Standard Libraries
import streamlit as st
import os
import sys
import math
import json
import random
import pickle
import re
import base64
import warnings
import time
from datetime import datetime, timedelta
import pytz

# Third-Party Libraries
import numpy as np
import pandas as pd
import pytz  # For timezone awareness
from tqdm import tqdm
import seaborn as sns
import matplotlib.pyplot as plt
import networkx as nx
import requests
from fuzzywuzzy import fuzz
from requests.auth import HTTPBasicAuth  # Importing HTTPBasicAuth

# Data Manipulation and Analysis
from sklearn.preprocessing import StandardScaler, MinMaxScaler
from collections import defaultdict, Counter
from itertools import repeat, combinations, permutations, combinations_with_replacement

# Excel and Data Handling
import xlsxwriter
from openpyxl.utils import get_column_letter

# Database Connectivity
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String

# Linear Programming
import pulp
from pulp import lpSum, LpMaximize
import gurobipy as grb
from gurobipy import Model, GRB

# Functional Programming
from functools import reduce

# Copying and Randomization
import copy

# Email Handling
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# API Requests
import mailchimp_marketing as MailchimpMarketing
from mailchimp_marketing.api_client import ApiClientError

# Set random seed for reproducibility
np.random.seed(24)

# Ignore Warnings
warnings.filterwarnings("ignore")



















# Initialize session state variables if they don't exist yet
if 'first_button_clicked' not in st.session_state:
    st.session_state.first_button_clicked = False
if 'second_button_clicked' not in st.session_state:
    st.session_state.second_button_clicked = False
if 'third_button_clicked' not in st.session_state:
    st.session_state.third_button_clicked = False
if 'email_address' not in st.session_state:
    st.session_state.email_address = ''  # Default email state

st.write('Hey there! This is the application to run the reporting for Lead Reports.')


## ADD an input here for the date!! 

st.write('Please click button to pull campaigns from the last two weeks :)')

# Create the big green button
if st.button('Retrieve Campaigns'):
    st.session_state.first_button_clicked = True  # Update session state when button is clicked

# Check if the big green button has been clicked
if st.session_state.first_button_clicked:
    st.write("Pulling Campaigns...")






### \/\/\/\/\/\/\/\/\/\/\/\/\/ PYTHON TO PULL CAMPAIGNS \/\/\/\/\/\/\/\/\/\/\/\/\/ 





    # Define file paths
    api_key_file_path = r"/Users/summerpurschke/Desktop/Keys:Passwords/mailchimp_API_KEY.txt"
    server_prefix_file_path = r"/Users/summerpurschke/Desktop/Keys:Passwords/mailchimp_SERVER_PREFIX.txt"

    # Try opening the file with a different encoding
    with open(api_key_file_path, 'r', encoding='utf-16') as file:
        API_KEY = file.read()

    # Try opening the file with a different encoding
    with open(server_prefix_file_path, 'r', encoding='utf-16') as file:
        SERVER_PREFIX = file.read()

        # Base URL for Mailchimp API requests
    BASE_URL = f'https://{SERVER_PREFIX}.api.mailchimp.com/3.0/'

    # Headers for authentication
    HEADERS = {
        'Authorization': f'Bearer {API_KEY}'
    }

    try:
        client = MailchimpMarketing.Client()
        client.set_config({
            "api_key": API_KEY,
            "server": SERVER_PREFIX
        })
        response = client.ping.get()
        print(response)
    except ApiClientError as error:
        print(error)
    ## Dataframe of all campaigns (emails) in the past 14 days
    # Mailchimp campaigns endpoint
    # url = f'https://{SERVER_PREFIX}.api.mailchimp.com/3.0/campaigns'
    url = f'{BASE_URL}campaigns'


    # Parameters for pagination (adjust count if needed)
    params = {
        'count': 200,  # Set the number of campaigns to retrieve per request (max is 1000)
        'offset': 0  # Start from the first campaign
    }

    all_campaigns = []

    while True:
        # Make the request
        response = requests.get(url, auth=HTTPBasicAuth('anystring', API_KEY), params=params)

        if response.status_code == 200:
            data = response.json()
            campaigns = data['campaigns']
            all_campaigns.extend(campaigns)  # Add current page campaigns to the list

            # Check if there are more campaigns to retrieve
            if len(campaigns) == 0:
                break  # No more campaigns, exit the loop

            # Increment the offset for the next page
            params['offset'] += len(campaigns)
        else:
            print(f"Failed to retrieve campaigns. Status Code: {response.status_code}, Response: {response.text}")
            break

    # Convert to Pandas DataFrame
    # Extract the campaign ID, title, and send date for each campaign
    campaigns_data = [
        {
            'Campaign ID': campaign['id'],
            'Title': campaign['settings']['title'],
            'Send Date': campaign.get('send_time')  # 'send_time' may be None if the campaign hasn't been sent
        }
        for campaign in all_campaigns
    ]
    ## Filter to only those campaigns that are within the last two weeks
    campaigns_df = pd.DataFrame(campaigns_data)

    # Convert 'Send Date' to datetime format (some may be None, so we'll handle that)
    campaigns_df['Send Date'] = pd.to_datetime(campaigns_df['Send Date'], errors='coerce')
    # Convert 'Send Date' to just the date (ignoring time and timezone)
    campaigns_df['Send Date'] = campaigns_df['Send Date'].dt.date
    # Sort by 'Send Date'
    campaigns_df = campaigns_df.sort_values(by='Send Date', ascending=False)

    # Get today's date as a naive date (ignoring time)
    today = datetime.now().date()

    # Define the start of the two-week window as a naive date
    two_weeks_ago = today - timedelta(days=14)

    # # Filter campaigns where 'Send Date' is within the past two weeks
    campaigns_df = campaigns_df[(campaigns_df['Send Date'] >= two_weeks_ago) & (campaigns_df['Send Date'] <= today)]

    # Tempory format fixing until everything is standardized
    campaigns_df['Title'] = campaigns_df['Title'].str.replace(' ','')
    campaigns_df['Title'] = campaigns_df['Title'].str.replace('-HVA','_HVA').str.replace('-CAMP', '_CAMP').str.replace('HUBTech', 'HUBTech_OG')
    campaigns_df['Title'] =campaigns_df['Title'].str.replace('DOTW-','DOTW_').str.replace('-CAMP','_CAMP').str.replace('-PROD','_PROD').str.replace('-HVA','_HVA')
    campaigns_df['Title'] = campaigns_df['Title'].str.replace('OG_OG_', 'OG_')
    ## Open and Click Reports
    ### Create Functions
    #### Function to retrieve Open Details per campaign
    def get_campaign_open_details(campaign_id):
        all_open_details = []
        offset = 0
        count = 1000  # Number of records per page (adjustable, max 1000)

        while True:
            try:
                # Fetch open details with pagination parameters
                response = client.reports.get_campaign_open_details(campaign_id, count=count, offset=offset)
                
                # Check if 'members' exists in the response
                if 'members' not in response:
                    print(f"Error: 'members' field not found in response for campaign {campaign_id}")
                    break
                
                # Append the open details to the list
                all_open_details.extend(response['members'])

                # Check if the number of results is less than the count (end of pages)
                if len(response['members']) < count:
                    break  # We've retrieved all pages
                
                # Increment the offset to get the next page of results
                offset += count

            except ApiClientError as error:
                raise ValueError(f"Open Function: Error fetching campaign open details: {error.text}")
                return None  # Handle API client errors

            except KeyError as key_error:
                raise ValueError(f"Open Function: KeyError: {key_error} in response for campaign {campaign_id}")
                break  # Exit loop on key error, preventing further requests

            except Exception as e:
                raise ValueError(f"Open Function: Unexpected error: {e}")
                break  # Exit loop if any other unexpected error occurs

        return all_open_details

    #### Function to retrieve Click and Open details per campaign (in json format)
    ## Function to retrieve Click reports 
    def get_campaign_activity_details(campaign_id):
        all_activity_details = []
        offset = 0 # start at record 0 (first) 
        count = 1000  # Number of records per page (adjustable, max 1000)

        while True:
            try:
                # Fetch open details with pagination parameters
                response = client.reports.get_email_activity_for_campaign(campaign_id, count=count, offset=offset)
                
                # Append the open details to the list
                all_activity_details.extend(response['emails'])

                # Check if the number of results is less than the count (end of pages)
                if len(response['emails']) < count:
                    break  # We've retrieved all pages
                
                # Increment the offset to get the next page of results
                offset += count

            except ApiClientError as error:
                # Need a value error here so that it adds this campaign to the rerun category
                raise ValueError(f"Click Function: Error fetching campaign click details: {error.text}")
                return None

        return all_activity_details
    #### Function to create a dataframe for all activity for ONE email
    def create_click_dataframe(campaign_email_json): 
        activity_df  = pd.DataFrame()

        for data in campaign_email_json:

            # Step 1: Flatten 'activity' to separate rows
            expanded_data = []
            email_address = data['email_address']
            campaign_id = data['campaign_id']

            # Check for each activity
            for act in data['activity']:
                # is this a click, open or other?
                activity_type = act.get('action')
                url = act.get('url')  # 'click' actions have a URL; 'open' actions do not

                # print(email_address, campaign_id, activity_type, url)
                expanded_data.append({
                    'email_address': email_address,
                    'campaign_id': campaign_id,
                    'action': activity_type,
                    'url': url
                })

        # if there is no expanded data its becasue there is no activity - we dont care about the emails with no activity! 
            if len(expanded_data):
                # Convert to DataFrame
                expanded_data_df = pd.DataFrame(expanded_data)

                # Step 2: Aggregate clicks and open actions
                # Create 'open' column (if 'open' action exists for that email)
                expanded_data_df['open'] = expanded_data_df['action'].apply(lambda x: 1 if x == 'open' else 0)

                # Filter out 'click' actions and group by email and URL to count clicks
                click_df = expanded_data_df[expanded_data_df['action'] == 'click'].groupby(['email_address', 'campaign_id', 'url']).size().reset_index(name='clicks')

                # Merge the click and open data
                # Get max open value per email (since it will be the same across all URLs for the same email)
                open_df = expanded_data_df.groupby('email_address')['open'].max().reset_index()

                # Merge the open info with click data
                merged_df = pd.merge(click_df, open_df, on='email_address')

                activity_df = pd.concat([activity_df, merged_df])
        return activity_df
    #### Function to Flatten a dictionary 
    def flatten_dict(d, parent_key='', sep='_'):
        """
        Flatten a nested dictionary.
        
        Parameters:
        - d (dict): The dictionary to flatten.
        - parent_key (str): The base key that is passed recursively.
        - sep (str): Separator between parent and child keys.
        
        Returns:
        - flat_dict (dict): Flattened dictionary.
        """
        items = []
        for k, v in d.items():
            new_key = parent_key + sep + k if parent_key else k
            if isinstance(v, dict):  # If value is a dict, recursively flatten it
                items.extend(flatten_dict(v, new_key, sep=sep).items())
            elif isinstance(v, list):  # If value is a list, flatten it
                for i, item in enumerate(v):
                    if isinstance(item, dict):  # Handle list of dicts
                        items.extend(flatten_dict(item, f"{new_key}_{i}", sep=sep).items())
                    else:
                        items.append((f"{new_key}_{i}", item))
            else:
                items.append((new_key, v))
        return dict(items)

    ### Pull records from API 
    campaigns_df = campaigns_df.sample(n = 5)
    # Remove any that are drafts

    campaigns_df = campaigns_df[~campaigns_df['Title'].str.contains('DRAFT')]
    campaigns_df = campaigns_df.sort_values(by = 'Send Date')






### ^^^^^^^^^^^^^^^^ PYTHON TO PULL CAMPAIGNS ^^^^^^^^^^^^^^^^^^^^^^^^^^






    st.write('Here are the Campaigns! Please look them over and confirm they are correct, then click next.')
    st.dataframe(campaigns_df[['Title', 'Send Date']])

    # Create the next step button
    if st.button('Process Campaign Data'):
        st.session_state.second_button_clicked = True  # Update session state when next button is clicked


# Check if the next step button has been clicked
if st.session_state.second_button_clicked:
    st.write("Processing campaign data, this may take up to 45 minutes...")





### \/\/\/\/\/\/\/\/\/\/\/\/\/ PYTHON TO CREATE LEAD REPORTS \/\/\/\/\/\/\/\/\/\/\/\/\/ 




    dfs_dict = {}
    open_details_dict = {}
    email_activity_dict = {}
    not_process_campaigns = []

    for index, row  in campaigns_df.iterrows():
        campaign_id = row['Campaign ID']
        title = row['Title']
        send_date = row['Send Date']

        print(f'        working on campaign {title}')

        # Pull details from the title
        company_name = title.split('_')[0]
        email_name  = title.split('_')[1]

    # If it doesn't work try it again later
        try: 
        ## OPEN DATAFRAMES - these are more for the information and details 
            ## Pull Open data for each and add to a dictionary
            open_details = get_campaign_open_details(campaign_id)
            open_details_dict[campaign_id] =open_details

        ## CLICK DATAFRAMES - these are more for the actual data of clicks and opens
            # Pull JSON of all email activity for this campaign_id - raw data not a df 
            email_activity = get_campaign_activity_details(campaign_id)
            email_activity_dict[campaign_id] = email_activity

            # Creates a dataframe for this specific email of activity with URL, clicks, and opens
            dfs_dict[f'click_df_{title}'] = create_click_dataframe(email_activity)
        except:
            not_process_campaigns.append(campaign_id)
    #### For those campaigns that timed out, rerun them
    for campaign_id in not_process_campaigns: 
        title =campaigns_df[campaigns_df['Campaign ID'] == campaign_id]['Title'].values[0]
        send_date = campaigns_df[campaigns_df['Campaign ID'] == campaign_id]['Send Date'].values[0]

        print(f'        working on campaign {title}')

        # Pull details from the title
        company_name = title.split('_')[0]
        email_name  = title.split('_')[1]

        ## OPEN DATAFRAMES - these are more for the information and details 
        ## Pull Open data for each and add to a dictionary
        open_details = get_campaign_open_details(campaign_id)
        open_details_dict[campaign_id] =open_details

        ## CLICK DATAFRAMES - these are more for the actual data of clicks and opens
        # Pull JSON of all email activity for this campaign_id - raw data not a df 
        email_activity = get_campaign_activity_details(campaign_id)
        email_activity_dict[campaign_id] = email_activity

        # # Creates a dataframe for this specific email of activity with URL, clicks, and opens
        dfs_dict[f'click_df_{title}'] = create_click_dataframe(email_activity)
    #### Flatten the open data and create a dataframe from it
    keys_to_keep = ['campaign_id', 'email_address', 'merge_fields', 'opens_count']

    for index, row in campaigns_df.iterrows():
        campaign_id = row['Campaign ID']
        # print(campaign_id)
        title = row['Title']
        st.write(f'     Pulling data for {title} from MailChimp')

        # Filter the dictionary down to the specified keys
        open_details_dict[campaign_id] = [
            {key: d[key] for key in keys_to_keep if key in d}  # Apply the filtering logic to each dict `d` in the list
            for d in open_details_dict[campaign_id]
        ]
        # Assuming open_details_dict[campaign_id] is a list of dictionaries
        open_details_dict[campaign_id] = [
            flatten_dict(d) for d in open_details_dict[campaign_id]  # Apply flatten_dict to each dict in the list
        ]
        # open_details_df = pd.json_normalize(open_details)
        dfs_dict[f'open_df_{title}'] = pd.DataFrame(open_details_dict[campaign_id])



    ## Process Data
    #### Pull list of companies
    companies = []
    for key in list(dfs_dict.keys()):
        company = (key.replace('open_','').replace('click_','').replace('df_','').split('_'))[0]
        companies.append(company)
    # Filter to only unique 
    companies = list(set(companies))
    #### Dictionary of campaigns per company
    campaigns_dict = {}
    for company in companies: 
        campaigns_dict[company] = {}
        campaigns = list(set([key.split("_")[3] for key in dfs_dict.keys() if company in key]))
        for campaign in campaigns:   
            campaigns_dict[company][campaign] = None
    #### Format the names correctly in the dataframes 
    col_mapping_dict =  {'email_address':'Email Address',
        'FNAME':'First Name', 
        'LNAME':'Last Name', 
        'COMPANY':'Company Name', 
        'TITLE':'Job Title',
        'ADDRESS':'Full Address',
        'MMERGE22':'Address - Street',
        'CITY':'Address - City', 
        'STATE':'Address - State',
        'ZIP':'Address - Zip', 
        'COUNTRY':'Address - Country', 
        'PHONE':'Phone Number', 
        'PLINKEDIN':'LinkedIn',
        'DOMAIN':'Company Domain', 
        'SUBSECTOR':'Sub-Sector', 
        'INDUSTRY':'Company Primary Industry',
        'EMPLOYEES':'Employee Count'}

    main_cols = ['Email Address',
        'First Name', 'Last Name', 'Company Name', 'Job Title', 'Full Address',
        'Address - Street', 'Address - City', 'Address - State',
        'Address - Zip', 'Address - Country', 'Phone Number', 'LinkedIn',
        'Company Domain', 'Sub-Sector', 'Company Primary Industry', 'Employee Count']

    # remove the merge_fields text from them 
    for item in [key for key in dfs_dict if 'open' in key]:
            dfs_dict[item].columns = dfs_dict[item].columns.str.replace('merge_fields_', '')
    # Look at all open dataframes that we pulled - open dataframes have a lot more information! 
    for item in [key for key in dfs_dict if 'open' in key]:
        # If all of the columns that we are trying to rename are in this dataframe, go ahead
        if all(key in dfs_dict[item].columns for key in col_mapping_dict.keys()):

            # Rename the columns of the specific DataFrame based on the column_mapping
            dfs_dict[item] = dfs_dict[item].rename(columns=col_mapping_dict)

            # if all of the main_cols are in the columns, subset this dataframe only to the columsn that we need! 
            if all(col in dfs_dict[item].columns for col in main_cols):
                dfs_dict[item] = dfs_dict[item][main_cols]
        # If all the columns we are trying to rename are not in here, we need to know about it
        else: 
            print(f'{item} does not have appropriate columns to rename - check this one')
    ## Merge into one Dataframe per company
    for company in companies: 
        for campaign in campaigns_dict[company].keys():

            # List of DataFrame keys to check
            current_working_keys = [key for key in  dfs_dict.keys() if f'{company}_' in key and f'_{campaign}_' in key and 'open' in key]
            # print(company, campaign, current_working_keys)
            # Check if all DataFrames have the same columns (check against first df)
            first_df_columns = dfs_dict[current_working_keys[0]].columns
            same_columns = all(dfs_dict[key].columns.equals(first_df_columns) for key in current_working_keys)

            if same_columns:
                # If they all have the same columns, concatenate them by stacking on top of one another
                df = pd.concat([dfs_dict[key] for key in current_working_keys], ignore_index=True)
                df = df.drop_duplicates()
                # Quality Check before assignment
                if df['Email Address'].value_counts().unique()[0] != 1:
                    # raise ValueError
                    print('!!!!!!!   SOME EMAILS STILL DUPLICATED')
                else:
                    campaigns_dict[company][campaign]= df
                    print(f'sucessfully combined for {company} campaign: {campaign}')

            else:
                # raise ValueError(
                print(f'!!!!!!!   COLUMNS ARE NOT THE SAME FOR ALL OPEN DFS FOR {company}')

    st.write('      Merging API data for each company... ')

    #### Add columns for each email (opens, clicks, urls)
    for company in companies: 
        # print('company: ', company)
        for campaign in campaigns_dict[company].keys():

            ### ADD CLICK AND OPEN DATA FROM ACTIVITY API REQUEST
            # print(' campaign: ',campaign)
            click_df_keys = [key for key in  dfs_dict.keys() if f'{company}_' in key and f'_{campaign}_' in key and 'click' in key]
            # print(click_df_keys)
            # for each click_dataframe 
            for key in click_df_keys:
                title = key.replace('click_df_','')
                send_date = campaigns_df[campaigns_df['Title'] == title]['Send Date'].values[0]
                send_date = send_date.strftime('%m.%d.%y')
                # send_date = send_date.astype('datetime64[s]').astype(datetime).strftime('%m.%d.%y')
                # print('     email: ',key,f'(sent {send_date})')
                df = dfs_dict[key]
                # display(df)

                # Aggregate the dataframe
                condensed_df = df.groupby('email_address').agg({
                    'clicks': 'sum',                      # Sum of clicks
                    'open': 'sum',                       # Mean of open (sum of values / number of values)
                    'url': lambda x: ', '.join(x.unique())  # Concatenate unique URLs into a string
                }).reset_index()

                condensed_df = condensed_df.rename(columns = {'email_address':'Email Address','url':f'{send_date} URL', 'clicks':f'{send_date} Clicks', 'open':f'{send_date} Opens'})

                ## Merge the dataframe with the main one  
                df =pd.merge(campaigns_dict[company][campaign], condensed_df, on='Email Address', how='left')



                ## ADD VALUES FOR OPENS 
                opened_emails = dfs_dict[f'open_df_{title}']['Email Address'].unique()


                campaigns_dict[company][campaign]  = df

    st.write('      Aggregating MailChimp Data for each email...')

    #### The above adds open values ONLY for those that were also clicked, this looks back and sees which were opened but not clicked
    for company in companies: 
        # print('company: ', company)
        for campaign in campaigns_dict[company].keys():
            for title in campaigns_df[campaigns_df['Title'].str.contains(f'{company}_{campaign}')]['Title'].unique():

                send_date = campaigns_df[campaigns_df['Title'] == title]['Send Date'].values[0]
                send_date = send_date.strftime('%m.%d.%y')

                open_emails = dfs_dict[f'open_df_{title}']['Email Address'].unique()

                # Set opens == 1 for this email, these are the people that opened but did not click anything
                campaigns_dict[company][campaign].loc[
                    # Where the opens doesnt have a value assigned to it yet
                    (campaigns_dict[company][campaign][f'{send_date} Opens'].isna()) & 
                    # Where the email is in the open dataframe
                    (campaigns_dict[company][campaign]['Email Address'].isin(open_emails)),
                    f'{send_date} Opens'] = 1
                
    st.write('      Calculating Unique Clicks, Unique Opens and Lead Scores...')
    ## Calculate Other Metrics
    #### Unique Clicks - How many emails did this person click on a URL in?
    for company in companies: 
        # print('company: ', company)
        for campaign in campaigns_dict[company].keys():

            click_cols = [col for col in campaigns_dict[company][campaign].columns if 'Clicks' in col]

            # Access the relevant DataFrame
            campaigns_dict[company][campaign]['Unique Clicks'] = (campaigns_dict[company][campaign][click_cols] >= 1).sum(axis=1)
    #### Unique Opens - How many emails did this person open? 
    for company in companies: 
        # print('company: ', company)
        for campaign in campaigns_dict[company].keys():

            open_cols = [col for col in campaigns_dict[company][campaign].columns if 'Open' in col]

            # Access the relevant DataFrame
            campaigns_dict[company][campaign]['Unique Opens'] = (campaigns_dict[company][campaign][open_cols] >= 1).sum(axis=1)
    #### Lead Score - P1 if they opened all emails and clicked on one link per email for all emails in this time frame
    for company in companies: 
        # print('company: ', company)
        for campaign in campaigns_dict[company].keys():

            # Count the number of emails
            open_cols = [col for col in campaigns_dict[company][campaign].columns if 'Open' in col if col != 'Unique Opens']

            no_emails = len(open_cols)
            campaigns_dict[company][campaign].loc[
                (campaigns_dict[company][campaign]['Unique Clicks'] ==  no_emails) & 
                (campaigns_dict[company][campaign]['Unique Opens'] ==  no_emails), 'Lead Tier'] = 'P1'
        






### ^^^^^^^^^^^^^^^^ PYTHON TO CREATE LEAD REPORTS ^^^^^^^^^^^^^^^^^^^^^^^^^^











    # Create the next step button
    st.write("Aggregations are complete! ")

    def create_excel_files(companies, campaigns_dict):
        excel_files = []
        today_str = datetime.today().strftime('%m.%d.%y')
        
        for company in companies:
            file_name = f"{today_str} {company} Customer Action Report.xlsx"
            # Specify a directory to save the files
            file_path = os.path.join('temp_files', file_name)  # Create a temp directory
            os.makedirs(os.path.dirname(file_path), exist_ok=True)  # Create directory if it doesn't exist
            
            with pd.ExcelWriter(file_path) as writer:
                for campaign, df in campaigns_dict[company].items():
                    df.to_excel(writer, sheet_name=f"{campaign} Contacts", index=False)
            
            excel_files.append(file_path)  # Append full path to list

        return excel_files

    # Create the files
    if st.button('Generate Reports'):
        excel_files = create_excel_files(companies, campaigns_dict)
        
        # Add download buttons for each file created
        for file_path in excel_files:
            with open(file_path, 'rb') as f:
                st.download_button(
                    label=f"Download {os.path.basename(file_path)}",
                    data=f,
                    file_name=os.path.basename(file_path),
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
    

    # # if the buttons are clicked, delete the files

    # # Clean up: remove the created files if needed
    # for file_name in excel_files:
    #     os.remove(file_name)















