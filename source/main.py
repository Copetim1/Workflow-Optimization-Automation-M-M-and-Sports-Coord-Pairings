
import pandas as pd
import os

from collections import defaultdict
from typing import DefaultDict



#Get the directory containing this script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Build path to the Excel file (in the same directory structure)
file_path = os.path.join(script_dir, 'files', 'Spring_2026_M&M_Responses.xlsx')

print(f"Looking for file at: {file_path}")

df = pd.read_excel(file_path)
print(df.head())

# df = pd.read_excel('source/files/Spring_2026_M&M_Responses.xlsx')
# df.head()
# print(df)
# print(df.head())

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