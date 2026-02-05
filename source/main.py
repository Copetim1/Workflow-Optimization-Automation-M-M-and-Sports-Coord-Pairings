
import pandas as pd
import os
from user import User
from mentor import Mentor
from mentee import Mentee
from collections import defaultdict
from typing import DefaultDict
 
#Get the directory containing this script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Build path to the Excel file (in the same directory structure)
file_path = os.path.join(script_dir, 'files', 'Spring_2026_M&M_Responses.xlsx')

print(f"Looking for file at: {file_path}")

df = pd.read_excel(file_path)
print(df.head())

#Creating an individual user from the first row of the dataframe
person = df.head(1)
name = person.iat[0,2]
email = person.iat[0,1]
year = person.iat[0,3]
number = person.iat[0,4]
discord = person.iat[0,5]
major = person.iat[0,6] if person.iat[0,6] is not None else ""
preferences = False if person.iat[0,12] == "No Preference" else True
answers = person.iloc[0, [13,34]]
user = User(name, email, year, number, discord, major, False, preferences, answers)

# if "Mentor" in person.iat[0,13]:
#     Mentor(name, email, year, 0,"", "", False, "", [])
'''
Create two hashmap of hashmaps.

Outer map:
key: names of mentors/mentees
value: Innermap
Inner map:
key: name of mentee/mentor
value: percentage for compatibility

'''
def addMentees(mentee) -> DefaultDict[str,int]:
    '''
    name = 
    email =
    year =
    number =
    discord =
    major =
    sameMajor =
    preferences = 

    #if statements and stuff to parse
    answers =
    #add answers to list
    '''


def addMentors(mentors)-> DefaultDict[str,int]:
   '''
    name = 
    email =
    year =
    number =
    discord =
    major =
    sameMajor =
    preferences = 
    answers =
    '''

def calcCompatibilityForMentee(mentee, mentor):
    x = 6

def calcCompatibilityForMentor(mentor, mentee):
    x = 5

mentees = defaultdict(dict)

mentors = defaultdict(dict)