import pandas as pd
from collections import defaultdict
from typing import DefaultDict

#df = pd.read_excel('')
#df.head()
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
    

def calcCompatibilityForMentor(mentor, mentee):

mentees = defaultdict(dict)
mentors = defaultdict(dict)