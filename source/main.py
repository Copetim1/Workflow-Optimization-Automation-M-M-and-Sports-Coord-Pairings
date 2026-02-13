import webbrowser
import pandas as pd
import csv
import os
import heapq
from user import User
from mentor import Mentor
from mentee import Mentee
from collections import defaultdict
from typing import DefaultDict
from typing import List


#Get the directory containing this script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Build path to the Excel file (in the same directory structure)
file_path = os.path.join(script_dir, 'files', 'Spring_2026_Responses.xlsx')

print(f"Looking for file at: {file_path}")

df = pd.read_excel(file_path)
#print(df.head())

#Creating an individual user from the first row of the dataframe
# person = df.head(1)
# name = person.iat[0,2]
# email = person.iat[0,1]
# year = person.iat[0,3]
# number = person.iat[0,4]
# discord = person.iat[0,5]
# major = person.iat[0,6] if person.iat[0,6] is not None else ""
# preferences = False if person.iat[0,12] == "No Preference" else True
#answers = person.iloc[0, [13,34]]

#print(df.columns.tolist()) #prints the columns of the first row 1

#Assigning and creating users to mentor or mentees dictionaries

mentees = []
mentors = []
errorUsers = []

completePairings = {}

finalPairings = {}



def assignAnswers(row)-> dict:
    answers = {}
    #If not found in mentee column, then go to mentor column, 
    #if not found in both, output = "No User Input"

    orgsInvolved = row['What student organizations are you currently involved/want to be more involved with? (Mentee)']
    if pd.isna(orgsInvolved):
        orgsInvolved = row['What student organizations are you currently involved/want to be more involved with? (Mentor)']
        if pd.isna(orgsInvolved):
            orgsInvolved = "No User Input"

    answers["Orgs Involved"] = orgsInvolved

    industryInterest = row['What industries/companies are you interested in working in and, if applicable, which industries/companies have you worked for? (Mentee)']
    if pd.isna(industryInterest): 
        industryInterest = row['What industries/companies are you interested in working in and, if applicable, which industries/companies have you worked for? (Mentor)']
        if pd.isna(industryInterest): 
            industryInterest = "No User Input"

    answers["IndustryInterest"] = industryInterest

    professionalHelp = row['What type of professional help do you want to receive? ']
    if pd.isna(professionalHelp):
        professionalHelp = row['What type of professional help do you want to provide?']
        if pd.isna(professionalHelp):
            professionalHelp = "No User Input"
    
    answers["ProfessionalHelp"] = professionalHelp

    personalityType = row['What is your MBTI Personality type? (Mentee)']
    if pd.isna(personalityType): 
        personalityType = row['What is your MBTI Personality type? (Mentor)']
        if pd.isna(personalityType): 
            personalityType = "No User Input"

    answers["PersonalityType"] = personalityType

    perfectRelationship = row['Describe your perfect relationship with your mentor.']
    if pd.isna(perfectRelationship):
        perfectRelationship = row['Describe your perfect relationship with your mentee.']
        if pd.isna(perfectRelationship):
            perfectRelationship = "No User Input"
    
    answers["PerfectRelationship"] = perfectRelationship

    groupInvolvement = row['How involved do you want to be with your M&M group? (Mentee)']
    if pd.isna(groupInvolvement):
        groupInvolvement = row ['How involved do you want to be with your M&M group? (Mentor)']
        if pd.isna(groupInvolvement):
            groupInvolvement = "No User Input"
        
    answers["GroupInvolvement"] = groupInvolvement

    cupInvolvement = row['How much do you want to win the M&M Cup? (Mentee)']
    if pd.isna(cupInvolvement):
        cupInvolvement = row['How much do you want to win the M&M Cup? (Mentor)']
        if pd.isna(cupInvolvement):
            cupInvolvement = "No User Input"
    
    answers["CupInvolvement"] = cupInvolvement

    userRequest = row['Do you have a specific mentor you would like to request? If yes, please list their FULL NAME separated by commas. (ex: Bryant Cao, Leann Tang)']
    if pd.isna(userRequest):
        userRequest = row['Do you have a specific mentee you would like to request? If yes, please list their FULL NAME separated by a comma and space. (ex: Bryant Cao, Leann Tang)']
        if pd.isna(userRequest):
            userRequest = "No User Input"
        
        answers["UserRequest"] = userRequest

    return answers
    
def initializeUsers():

    for i, row in df.iterrows():
        name = row['Name (First Last)']
        email = row['Email address']
        year = row['Year']
        number = row['Phone Number']
        discord = row['Discord (e.g. jellyduck224)']
        answers = {}

        sameMajorOrNah = row['Do you want your mentor/mentee to be the same major?']
        if  sameMajorOrNah == "Same major":
            preferences = True
        else:
            preferences = False
        
        academicFocus = row['What best describes your academic focus?']
        if academicFocus == "Pre-health" or academicFocus == "Pre-law":
            major = row['What is your major? (Pre-health)']
        elif academicFocus == "College of Engineering":
            major = row['What is your major? (Engineering)']
        elif academicFocus == "Other":
            major = row['What is your major? (Other)']
        elif academicFocus == "Postgrad":
            major = row['What is your major? (Postgrad)']

        answers = assignAnswers(row)
        
        mentorOrMentee = row['I want to be ...']

        if mentorOrMentee == "a Mentee!":
            user = Mentee(name, email, year, number, discord, academicFocus, major, preferences, answers)
            mentees.append(user)

        elif mentorOrMentee == "a Mentor!":

            numMentees = row['Do you have a preference on how many mentees you receive?']
            if pd.isna(numMentees):
                numMentees = 0

            particpateInMMCup = row['To encourage interaction and assist with bonding between Mentors and Mentees, SASE hosts a semester-long mentorship cup, where you will participate in biweekly challenges with your mentees such as eating food together. Do you understand that as a mentor, your mentees may want you to participate in these challenges?']
            if particpateInMMCup == "Yes, I understand":
                answers["participateInMMCup"] = "Participating"
            else:
                answers["participateInMMCup"] = "Not Participating"

            user = Mentor(name, email, year, number, discord, academicFocus, major, preferences, answers, numMentees)
            mentors.append(user)

        else:
            mentorOrMentee == "Something went wrong"
            errorUsers.append(user)
        
    
#df_head = df.head(2)

#print(df_head.iloc[0,1])

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


#Could def add more majors here
def isSimilarMajor(major1, major2):
    # Define groups of similar majors

    similarMajors = {
        'engineering': ['Mechanical Engineering','Aerospace Engineering', 'Electrical Engineering', 'Civil Engineering', 
                       'Chemical Engineering', 'Computer Engineering','Nuclear Engineering', 'Industrial & Systems Engineering',
                         'Industrial Systems Engineering', 'Biomedical Engineering', 'Digital Arts & Sciences'],

        'computer_science': ['Computer Science', 'Information Systems', 
                             'Data Science', 'Digital Arts & Sciences'],

        'business': ['Business Administration', 'Finance', 'Accounting', 'Marketing', 
                    'Economics'],

        'pre_health': ['Health Science', 'Biomedical Engineering', 'Microbiology and Cell Science',
                        'Chemistry', 'Biochemistry', 'Biology', 'Public Health', 'Microbiology'] 
    }
    
    # Check if both majors are in the same group
    for group, majors in similarMajors.items():
        if major1 in majors and major2 in majors:
            return True

    return False


#Assigns score based on how close major relates to each other
#70 pts max
def assignScoreMajor(mentee, mentor):

    score = 0
    #= major 70pts
    if mentee.getMajor == mentor.getMajor:
        
        score = 70
        return score
    
    #= academic focus 20pts
    if mentee.getAcademicFocus == mentor.getAcademicFocus:
        score += 20
    
    #similar major 25pts
    if isSimilarMajor(mentee.getMajor, mentor.getMajor) == True:
        score += 25

    return score

#Assigns score based on how many similar professional and
#perfect relationship answers are in common
#55 pts Max
def assignScoreProfHelpPerfRelationship(mentee, mentor):
    score = 0
    menteeAnswers = mentee.getAnswers
    mentorAnswers = mentor.getAnswers

    check = menteeAnswers["ProfessionalHelp"]
    if check == "No User Input":
        return 0
    check = mentorAnswers["ProfessionalHelp"]
    if check == "No User Input":
        return 0
    
    menteeHelp = menteeAnswers["ProfessionalHelp"].split(", ")
    mentorHelp = mentorAnswers["ProfessionalHelp"].split(", ")
    
    common = set(menteeHelp)& set(mentorHelp)
    count = len(common)

    menteeHelp = menteeAnswers["PerfectRelationship"].split(", ")
    mentorHelp = mentorAnswers["PerfectRelationship"].split(", ")

    common = set(menteeHelp)& set(mentorHelp)
    count += len(common)
    score += count * 5

    return score

#MBTI Compatibility
#30pts Max
def assignScoreMBTI(mentee, mentor):
    score = 0
    menteeAnswers = mentee.getAnswers
    mentorAnswers = mentor.getAnswers

    menteeMBTI = menteeAnswers["PersonalityType"]
    mentorMBTI = mentorAnswers["PersonalityType"]

    if menteeMBTI == "No User Input" or mentorMBTI == "No User Input":
        return 0
    
    #Define MBTI compatibility scores
    #Based on Myer Briggs Type compatibility: Best match, Good match, Okay match, Challenging match
    compatibility = {
        # ISTJ compatibilities
        'ISTJ': {
            'ESTP': 30, 'ESFP': 30,  # Best matches
            'ISTJ': 25, 'ISFJ': 25, 'ESTJ': 25, 'ESFJ': 25,  # Good matches
            'ISTP': 15, 'ISFP': 15, 'INTJ': 15, 'INFJ': 15,  # Okay matches
            'INTP': 10, 'INFP': 10, 'ENTP': 10, 'ENFP': 10, 'ENTJ': 10, 'ENFJ': 10  # Challenging
        },
        
        # ISFJ compatibilities
        'ISFJ': {
            'ESFP': 30, 'ESTP': 30,
            'ISFJ': 25, 'ISTJ': 25, 'ESFJ': 25, 'ESTJ': 25,
            'ISFP': 15, 'ISTP': 15, 'INFJ': 15, 'INTJ': 15,
            'INFP': 10, 'INTP': 10, 'ENFP': 10, 'ENTP': 10, 'ENFJ': 10, 'ENTJ': 10
        },
        
        # INFJ compatibilities
        'INFJ': {
            'ENFP': 30, 'ENTP': 30,
            'INFJ': 25, 'INTJ': 25, 'ENFJ': 25, 'ENTJ': 25,
            'INFP': 15, 'INTP': 15, 'ISFJ': 15, 'ISTJ': 15,
            'ISFP': 10, 'ISTP': 10, 'ESFP': 10, 'ESTP': 10, 'ESFJ': 10, 'ESTJ': 10
        },
        
        # INTJ compatibilities
        'INTJ': {
            'ENFP': 30, 'ENTP': 30,
            'INTJ': 25, 'INFJ': 25, 'ENTJ': 25, 'ENFJ': 25,
            'INTP': 15, 'INFP': 15, 'ISTJ': 15, 'ISFJ': 15,
            'ISTP': 10, 'ISFP': 10, 'ESTP': 10, 'ESFP': 10, 'ESTJ': 10, 'ESFJ': 10
        },
        
        # ISTP compatibilities
        'ISTP': {
            'ESFJ': 30, 'ESTJ': 30,
            'ISTP': 25, 'ISFP': 25, 'ESTP': 25, 'ESFP': 25,
            'ISTJ': 15, 'ISFJ': 15, 'INTJ': 15, 'INFJ': 15,
            'INTP': 10, 'INFP': 10, 'ENTP': 10, 'ENFP': 10, 'ENTJ': 10, 'ENFJ': 10
        },
        
        # ISFP compatibilities
        'ISFP': {
            'ESFJ': 30, 'ESTJ': 30,
            'ISFP': 25, 'ISTP': 25, 'ESFP': 25, 'ESTP': 25,
            'ISFJ': 15, 'ISTJ': 15, 'INFJ': 15, 'INTJ': 15,
            'INFP': 10, 'INTP': 10, 'ENFP': 10, 'ENTP': 10, 'ENFJ': 10, 'ENTJ': 10
        },
        
        # INFP compatibilities
        'INFP': {
            'ENFJ': 30, 'ENTJ': 30,
            'INFP': 25, 'INTP': 25, 'ENFP': 25, 'ENTP': 25,
            'INFJ': 15, 'INTJ': 15, 'ISFP': 15, 'ISTP': 15,
            'ISFJ': 10, 'ISTJ': 10, 'ESFJ': 10, 'ESTJ': 10, 'ESFP': 10, 'ESTP': 10
        },
        
        # INTP compatibilities
        'INTP': {
            'ENTJ': 30, 'ENFJ': 30,
            'INTP': 25, 'INFP': 25, 'ENTP': 25, 'ENFP': 25,
            'INTJ': 15, 'INFJ': 15, 'ISTP': 15, 'ISFP': 15,
            'ISTJ': 10, 'ISFJ': 10, 'ESTJ': 10, 'ESFJ': 10, 'ESTP': 10, 'ESFP': 10
        },
        
        # ESTP compatibilities
        'ESTP': {
            'ISFJ': 30, 'ISTJ': 30,
            'ESTP': 25, 'ESFP': 25, 'ISTP': 25, 'ISFP': 25,
            'ESTJ': 15, 'ESFJ': 15, 'ENTP': 15, 'ENFP': 15,
            'ENTJ': 10, 'ENFJ': 10, 'INTJ': 10, 'INFJ': 10, 'INTP': 10, 'INFP': 10
        },
        
        # ESFP compatibilities
        'ESFP': {
            'ISTJ': 30, 'ISFJ': 30,
            'ESFP': 25, 'ESTP': 25, 'ISFP': 25, 'ISTP': 25,
            'ESFJ': 15, 'ESTJ': 15, 'ENFP': 15, 'ENTP': 15,
            'ENFJ': 10, 'ENTJ': 10, 'INFJ': 10, 'INTJ': 10, 'INFP': 10, 'INTP': 10
        },
        
        # ENFP compatibilities
        'ENFP': {
            'INTJ': 30, 'INFJ': 30,
            'ENFP': 25, 'ENTP': 25, 'INFP': 25, 'INTP': 25,
            'ENFJ': 15, 'ENTJ': 15, 'ESFP': 15, 'ESTP': 15,
            'ESFJ': 10, 'ESTJ': 10, 'ISFJ': 10, 'ISTJ': 10, 'ISFP': 10, 'ISTP': 10
        },
        
        # ENTP compatibilities
        'ENTP': {
            'INFJ': 30, 'INTJ': 30,
            'ENTP': 25, 'ENFP': 25, 'INTP': 25, 'INFP': 25,
            'ENTJ': 15, 'ENFJ': 15, 'ESTP': 15, 'ESFP': 15,
            'ESTJ': 10, 'ESFJ': 10, 'ISTJ': 10, 'ISFJ': 10, 'ISTP': 10, 'ISFP': 10
        },
        
        # ESTJ compatibilities
        'ESTJ': {
            'ISFP': 30, 'ISTP': 30,
            'ESTJ': 25, 'ESFJ': 25, 'ISTJ': 25, 'ISFJ': 25,
            'ESTP': 15, 'ESFP': 15, 'ENTJ': 15, 'ENFJ': 15,
            'ENTP': 10, 'ENFP': 10, 'INTJ': 10, 'INFJ': 10, 'INTP': 10, 'INFP': 10
        },
        
        # ESFJ compatibilities
        'ESFJ': {
            'ISTP': 30, 'ISFP': 30,
            'ESFJ': 25, 'ESTJ': 25, 'ISFJ': 25, 'ISTJ': 25,
            'ESFP': 15, 'ESTP': 15, 'ENFJ': 15, 'ENTJ': 15,
            'ENFP': 10, 'ENTP': 10, 'INFJ': 10, 'INTJ': 10, 'INFP': 10, 'INTP': 10
        },
        
        # ENFJ compatibilities
        'ENFJ': {
            'INFP': 30, 'INTP': 30,
            'ENFJ': 25, 'ENTJ': 25, 'INFJ': 25, 'INTJ': 25,
            'ENFP': 15, 'ENTP': 15, 'ESFJ': 15, 'ESTJ': 15,
            'ESFP': 10, 'ESTP': 10, 'ISFJ': 10, 'ISTJ': 10, 'ISFP': 10, 'ISTP': 10
        },
        
        # ENTJ compatibilities
        'ENTJ': {
            'INTP': 30, 'INFP': 30,
            'ENTJ': 25, 'ENFJ': 25, 'INTJ': 25, 'INFJ': 25,
            'ENTP': 15, 'ENFP': 15, 'ESTJ': 15, 'ESFJ': 15,
            'ESTP': 10, 'ESFP': 10, 'ISTJ': 10, 'ISFJ': 10, 'ISTP': 10, 'ISFP': 10
        }
    }
    
    if mentorMBTI in compatibility and menteeMBTI in compatibility[mentorMBTI]:
        score = compatibility[mentorMBTI][menteeMBTI]
    
    return score
    

#Assign score based on involvment with group and for the M&M cup   
#40 pts max
def assignScoreInvolvement(mentee, mentor):
    score = 0
    menteeAnswers = mentee.getAnswers
    mentorAnswers = mentor.getAnswers

    menteeT = menteeAnswers["GroupInvolvement"]
    mentorT = mentorAnswers["GroupInvolvement"]

    if menteeT == mentorT:
        score += 20
    elif abs(menteeT - mentorT) <= 2:
        score += 15
    elif abs(menteeT - mentorT) < 5:
        score+= 10
    
    menteeT = menteeAnswers["CupInvolvement"]
    mentorT = mentorAnswers["CupInvolvement"]

    if menteeT == mentorT:
        score += 20
    elif abs(menteeT - mentorT) <= 2:
        score += 15
    elif abs(menteeT - mentorT) < 5:
        score+= 10

    return score


#Calculates compatibility between a mentor and all the mentees
#195 Max Score
def calcScore_Mentor(mentor, mentees):
    menteeScores = {}

    for mentee in mentees:

        if mentor.getPreferences == True and mentee.getMajor != mentor.getMajor:
            continue  # Skip this mentee

        if mentee.getYear >= mentor.getYear:
            continue
        
        if mentee.getName == mentor.getName:
            continue

        # Filter: If mentee wants same major, skip mentors with different majors  
        if mentee.getPreferences == True and mentee.getMajor != mentor.getMajor:
            continue  # Skip this mentee
        score = 0
        score += assignScoreMajor(mentee, mentor)
        score += assignScoreProfHelpPerfRelationship(mentee, mentor)
        score += assignScoreMBTI(mentee, mentor)
        score += assignScoreInvolvement(mentee, mentor)
        mentee.setScore(score)

        menteeScores[mentee.getName] = mentee.getScores

    return menteeScores


#Calculates compatibility between a mentee and all the mentors
def calcScore_Mentee(mentee, mentors):
    mentorScores = {}
    for mentor in mentors:
        score = 0
        
        # Avoid self-pairing
        if mentee.getName == mentor.getName:
            continue

        # CHANGE: Instead of skipping, give a large penalty for year mismatch
        if mentee.getYear >= mentor.getYear:
            score -= 100 
        
        # CHANGE: Instead of skipping, give a penalty for major mismatch if preferred
        if mentee.getPreferences == True and mentee.getMajor != mentor.getMajor:
            score -= 50
            
        # Add the rest of your scoring logic
        score += assignScoreMajor(mentee, mentor)
        score += assignScoreProfHelpPerfRelationship(mentee, mentor)
        score += assignScoreMBTI(mentee, mentor)
        score += assignScoreInvolvement(mentee, mentor)
        
        # Ensure score doesn't go below a baseline so they stay in the list
        mentorScores[mentor.getName] = max(1, score) 

    return mentorScores


def initializeFinalMap():
    completePairs = {}

    for mentor in mentors:
        temp = calcScore_Mentor(mentor,mentees)
        completePairs[mentor.getName] = temp

    return completePairs




def gale_shapley_round(mentees_to_match, mentors, existing_pairings, round_number):
    print(f"\n--- Starting Gale-Shapley Round {round_number} ---")
    
    # 1. Setup Data Structures
    mentor_capacity = {}
    mentor_current_matches = {} # {MentorName: [(Score, MenteeName), ...]}
    
    for mentor in mentors:
        # Check how many mentees they ALREADY have from previous rounds
        current_count = len(existing_pairings.get(mentor.getName, []))
        total_limit = mentor.getNumMentees
        
        # Calculate remaining capacity for THIS round [cite: 61, 62]
        if total_limit == 0:
            remaining = 999 - current_count
        else:
            remaining = total_limit - current_count
            
        mentor_capacity[mentor.getName] = max(0, remaining)
        mentor_current_matches[mentor.getName] = []

    # 2. Build Preference Lists and CACHE scores [cite: 63, 68]
    free_mentees = [] 
    mentee_prefs = {} 
    all_mentee_scores = {} 

    for mentee in mentees_to_match:
        free_mentees.append(mentee.getName)
        
        # Get scores for this mentee against ALL mentors
        scores = calcScore_Mentee(mentee, mentors)
        all_mentee_scores[mentee.getName] = scores 
        
        # Find mentors they are already paired with to avoid duplicates [cite: 64, 75]
        already_paired_with = []
        for m_name, paired_list in existing_pairings.items():
            if mentee.getName in paired_list:
                already_paired_with.append(m_name)
        
        # Sort mentors by score (Highest first), excluding current matches [cite: 65]
        sorted_mentors = sorted(
            [m for m in scores.items() if m[0] not in already_paired_with], 
            key=lambda x: x[1], 
            reverse=True
        )
        
        mentee_prefs[mentee.getName] = [m[0] for m in sorted_mentors]

    mentee_proposal_index = {m_name: 0 for m_name in free_mentees}

    # 3. The Algorithm Loop [cite: 66]
    while free_mentees:
        mentee_name = free_mentees.pop(0)
        prefs = mentee_prefs[mentee_name]
        idx = mentee_proposal_index[mentee_name]
        
        if idx >= len(prefs):
            continue
            
        target_mentor_name = prefs[idx]
        mentee_proposal_index[mentee_name] += 1 
        
        # Retrieve the specific score from our cached dictionary [cite: 69]
        score = all_mentee_scores[mentee_name].get(target_mentor_name, 0)
        
        matches = mentor_current_matches[target_mentor_name]
        cap = mentor_capacity[target_mentor_name]
        
        # FIX: Handle mentors who are already full from Round 1
        if cap == 0:
            free_mentees.append(mentee_name)
            continue

        if len(matches) < cap:
            # Mentor has space in THIS round 
            heapq.heappush(matches, (score, mentee_name))
        else:
            # Mentor is full for this round - compare with lowest current match [cite: 71]
            # Safety check: ensure matches is not empty before accessing index 0
            if matches:
                lowest_score, lowest_mentee = matches[0]
                
                if score > lowest_score:
                    # Swap for the better match [cite: 72, 73]
                    heapq.heappop(matches) 
                    heapq.heappush(matches, (score, mentee_name))
                    free_mentees.append(lowest_mentee) 
                else:
                    # New mentee is worse; keep current matches [cite: 74]
                    free_mentees.append(mentee_name)
            else:
                # Fallback if capacity exists but matches list is empty
                free_mentees.append(mentee_name)

    # 4. Update the Master Pairings List [cite: 75]
    new_pairings_count = 0
    for mentor_name, match_list in mentor_current_matches.items():
        if mentor_name not in existing_pairings:
            existing_pairings[mentor_name] = []
        
        for score, mentee_name in match_list:
            existing_pairings[mentor_name].append(mentee_name)
            new_pairings_count += 1
            
    print(f"Round {round_number} Complete. {new_pairings_count} matches formed.")
    return existing_pairings

def pairing(completePairings):
    # This function now orchestrates the Two-Round Gale-Shapley
    
    finalPairings = {} # {MentorName: [MenteeName1, MenteeName2]}
    
    # ----------------------------------------
    # ROUND 1: EVERYONE GETS THEIR FIRST MATCH
    # ----------------------------------------
    # In this round, we want every mentee to secure 1 spot.
    
    finalPairings = gale_shapley_round(mentees, mentors, finalPairings, 1)
    
    # ----------------------------------------
    # ROUND 2: FILLING SECOND SPOTS
    # ----------------------------------------
    # Only mentees who want a second mentor enter this round.
    # (And mentors only participate if they have capacity left)
    
    # Optional: Filter mentees who specifically asked for 2? 
    # For now, we assume everyone is eligible for a second if they want it.
    
    finalPairings = gale_shapley_round(mentees, mentors, finalPairings, 2)
    
    # ----------------------------------------
    # GENERATE STATS
    # ----------------------------------------
    paired_mentees = {}
    for mentee in mentees:
        paired_mentees[mentee.getName] = []
        
    for mentor_name, mentee_list in finalPairings.items():
        for m_name in mentee_list:
            if m_name in paired_mentees:
                paired_mentees[m_name].append(mentor_name)
    
    unpaired_mentees = [m.getName for m in mentees if len(paired_mentees.get(m.getName, [])) == 0]
    partially_paired = [m.getName for m in mentees if len(paired_mentees.get(m.getName, [])) == 1]
    
    return finalPairings, unpaired_mentees, partially_paired, paired_mentees

def export_rankings_to_csv(mentors, completePairings, filename='mentor_rankings.csv'):
    # Get the directory to save the file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, 'files', filename)
    
    # Prepare the data rows
    rows = []
    max_mentees_count = 0

    for mentor in mentors:
        mentor_name = mentor.getName
        
        # Get all compatible mentees and scores
        # completePairings structure: { MentorName: { MenteeName: Score, ... } }
        mentee_dict = completePairings.get(mentor_name, {})
        
        # Sort by score descending (Highest to Lowest)
        sorted_mentees = sorted(mentee_dict.items(), key=lambda item: item[1], reverse=True)
        
        # Track max length for header generation later
        if len(sorted_mentees) > max_mentees_count:
            max_mentees_count = len(sorted_mentees)
            
        # Build the row: [Mentor Name, Mentee1, Score1, Mentee2, Score2, ...]
        row = [mentor_name]
        for mentee_name, score in sorted_mentees:
            row.append(mentee_name)
            row.append(score)
            
        rows.append(row)

    # Create the dynamic header
    # Header: Mentor Name, Rank 1 Mentee, Rank 1 Score, Rank 2 Mentee, Rank 2 Score...
    header = ['Mentor Name']
    for i in range(1, max_mentees_count + 1):
        header.append(f'Rank {i} Mentee')
        header.append(f'Rank {i} Score')

    # Write to CSV
    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(header) # Write header
            writer.writerows(rows)  # Write all data rows
            
        print(f"✓ Rankings exported successfully to: {output_path}")
        # Open the file automatically for you
        os.startfile(output_path) 
    except Exception as e:
        print(f"Error exporting CSV: {e}")




def export_mentee_rankings_to_csv(mentees, mentors, filename='mentee_rankings.csv'):
    # Get the directory to save the file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, 'files', filename)
    
    rows = []
    max_matches_count = 0

    print(f"Generating Mentee Rankings CSV...")

    for mentee in mentees:
        # Calculate compatibility with ALL mentors for this mentee
        # This uses your existing function
        mentor_scores = calcScore_Mentee(mentee, mentors)
        
        # Sort by score descending (Highest to Lowest)
        # mentor_scores is a dict { 'MentorName': Score }
        sorted_mentors = sorted(mentor_scores.items(), key=lambda item: item[1], reverse=True)
        
        # Track max length for header generation
        if len(sorted_mentors) > max_matches_count:
            max_matches_count = len(sorted_mentors)
            
        # Build the row: [Mentee Name, Mentor1, Score1, Mentor2, Score2, ...]
        row = [mentee.getName]
        for mentor_name, score in sorted_mentors:
            row.append(mentor_name)
            row.append(score)
            
        rows.append(row)

    # Create the dynamic header
    # Header: Mentee Name, Rank 1 Mentor, Rank 1 Score, ...
    header = ['Mentee Name']
    for i in range(1, max_matches_count + 1):
        header.append(f'Rank {i} Mentor')
        header.append(f'Rank {i} Score')

    # Write to CSV
    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerows(rows)
            
        print(f"✓ Mentee Rankings exported to: {output_path}")
        # Try to open the file automatically (Windows)
        try:
            os.startfile(output_path)
        except:
            pass # Ignore if on Mac/Linux
            
    except Exception as e:
        print(f"Error exporting CSV: {e}")


def export_compatibility_matrix(mentors, mentees, filename='compatibility_matrix.csv'):
    # Get the directory to save the file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, 'files', filename)
    
    print(f"Generating Compatibility Matrix CSV...")
    
    # 1. Prepare data for a DataFrame
    # We will create a list of rows where each row is a mentor
    matrix_data = []
    
    for mentor in mentors:
        # Create a dictionary for this mentor
        # The key will be the column name (Mentee Name) and the value will be the score
        row = {'Mentor Name': mentor.getName}
        
        for mentee in mentees:
            # We use your existing score calculation functions
            # Note: This ignores the filters (Year/Major) to show the "raw" score 
            # If you want to see 0s for filtered out pairs, you can wrap this in a try/except
            score = (assignScoreMajor(mentee, mentor) + 
                     assignScoreProfHelpPerfRelationship(mentee, mentor) +
                     assignScoreMBTI(mentee, mentor) + 
                     assignScoreInvolvement(mentee, mentor))
            
            row[mentee.getName] = score
            
        matrix_data.append(row)

    # 2. Convert to DataFrame and Export
    try:
        matrix_df = pd.DataFrame(matrix_data)
        
        # Set Mentor Name as the index so it becomes the first column
        matrix_df.set_index('Mentor Name', inplace=True)
        
        # Write to CSV
        matrix_df.to_csv(output_path)
        
        print(f"✓ Compatibility Matrix exported to: {output_path}")
        
        # Try to open the file automatically (Windows)
        try:
            os.startfile(output_path)
        except:
            pass
            
    except Exception as e:
        print(f"Error exporting Matrix CSV: {e}")



def export_to_html(finalPairings, mentors, mentees, paired_mentees, filename='mentorship_report.html'):
    """
    Export to an interactive HTML report.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, 'files', filename)
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Mentorship Pairings Report</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 20px;
                background-color: #f5f5f5;
            }
            h1 {
                color: #333;
                border-bottom: 3px solid #4CAF50;
                padding-bottom: 10px;
            }
            h2 {
                color: #555;
                margin-top: 30px;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                background-color: white;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                margin-bottom: 30px;
            }
            th {
                background-color: #4CAF50;
                color: white;
                padding: 12px;
                text-align: left;
                position: sticky;
                top: 0;
            }
            td {
                padding: 10px;
                border-bottom: 1px solid #ddd;
            }
            tr:hover {
                background-color: #f5f5f5;
            }
            .stats {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 20px;
                margin-bottom: 30px;
            }
            .stat-card {
                background: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            .stat-number {
                font-size: 32px;
                font-weight: bold;
                color: #4CAF50;
            }
            .stat-label {
                color: #666;
                margin-top: 5px;
            }
            .high-score { color: #4CAF50; font-weight: bold; }
            .medium-score { color: #FF9800; }
            .low-score { color: #F44336; }
            .tabs {
                display: flex;
                gap: 10px;
                margin-bottom: 20px;
            }
            .tab {
                padding: 10px 20px;
                background: white;
                border: none;
                cursor: pointer;
                border-radius: 5px 5px 0 0;
            }
            .tab.active {
                background: #4CAF50;
                color: white;
            }
            .tab-content {
                display: none;
            }
            .tab-content.active {
                display: block;
            }
            .candidate-list {
                font-size: 0.9em;
                line-height: 1.6;
            }
            .candidate-item {
                margin: 5px 0;
                padding: 5px;
                background: #f9f9f9;
                border-left: 3px solid #4CAF50;
                padding-left: 10px;
            }
            .candidate-score {
                font-weight: bold;
                color: #4CAF50;
            }
            .available-badge {
                display: inline-block;
                background: #4CAF50;
                color: white;
                padding: 2px 8px;
                border-radius: 3px;
                font-size: 0.8em;
                margin-left: 5px;
            }
            .at-capacity-badge {
                display: inline-block;
                background: #F44336;
                color: white;
                padding: 2px 8px;
                border-radius: 3px;
                font-size: 0.8em;
                margin-left: 5px;
            }
        </style>
        <script>
            function showTab(tabName) {
                // Hide all tab contents
                const contents = document.querySelectorAll('.tab-content');
                contents.forEach(c => c.classList.remove('active'));
                
                // Remove active from all tabs
                const tabs = document.querySelectorAll('.tab');
                tabs.forEach(t => t.classList.remove('active'));
                
                // Show selected tab
                document.getElementById(tabName).classList.add('active');
                event.target.classList.add('active');
            }
            
            function sortTable(tableId, column) {
                const table = document.getElementById(tableId);
                const rows = Array.from(table.querySelectorAll('tbody tr'));
                const isNumeric = !isNaN(rows[0].cells[column].textContent);
                
                rows.sort((a, b) => {
                    const aVal = a.cells[column].textContent;
                    const bVal = b.cells[column].textContent;
                    
                    if (isNumeric) {
                        return parseFloat(bVal) - parseFloat(aVal);
                    }
                    return aVal.localeCompare(bVal);
                });
                
                rows.forEach(row => table.querySelector('tbody').appendChild(row));
            }
        </script>
    </head>
    <body>
        <h1>🎓 Mentorship Pairings Report</h1>
    """
    
    # Calculate statistics
    total_mentees = len(mentees)
    total_mentors = len(mentors)
    fully_paired = sum(1 for v in paired_mentees.values() if len(v) == 2)
    partially_paired = sum(1 for v in paired_mentees.values() if len(v) == 1)
    unpaired = sum(1 for v in paired_mentees.values() if len(v) == 0)
    total_pairings = sum(len(v) for v in finalPairings.values())
    
    html_content += f"""
        <div class="stats">
            <div class="stat-card">
                <div class="stat-number">{total_mentors}</div>
                <div class="stat-label">Total Mentors</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{total_mentees}</div>
                <div class="stat-label">Total Mentees</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{total_pairings}</div>
                <div class="stat-label">Total Pairings</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{fully_paired}</div>
                <div class="stat-label">Fully Paired (2 mentors)</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{partially_paired}</div>
                <div class="stat-label">Partially Paired (1 mentor)</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{unpaired}</div>
                <div class="stat-label">Unpaired</div>
            </div>
        </div>
        
        <div class="tabs">
            <button class="tab active" onclick="showTab('pairings')">All Pairings</button>
            <button class="tab" onclick="showTab('mentors')">By Mentor</button>
            <button class="tab" onclick="showTab('mentees')">By Mentee</button>
            <button class="tab" onclick="showTab('unpaired')">Needs Pairing</button>
        </div>
    """
    
    # Tab 1: All Pairings
    html_content += """
        <div id="pairings" class="tab-content active">
            <h2>All Pairings</h2>
            <table id="pairings-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('pairings-table', 0)">Mentor</th>
                        <th onclick="sortTable('pairings-table', 1)">Mentor Major</th>
                        <th onclick="sortTable('pairings-table', 2)">Mentee</th>
                        <th onclick="sortTable('pairings-table', 3)">Mentee Major</th>
                        <th onclick="sortTable('pairings-table', 4)">Score</th>
                        <th>Mentor Email</th>
                        <th>Mentee Email</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for mentor_name, mentee_list in finalPairings.items():
        mentor = next((m for m in mentors if m.getName == mentor_name), None)
        if not mentor:
            continue
            
        for mentee_name in mentee_list:
            mentee = next((m for m in mentees if m.getName == mentee_name), None)
            if not mentee:
                continue
            
            total_score = (assignScoreMajor(mentee, mentor) + 
                          assignScoreProfHelpPerfRelationship(mentee, mentor) +
                          assignScoreMBTI(mentee, mentor) + 
                          assignScoreInvolvement(mentee, mentor))
            
            score_class = 'high-score' if total_score >= 130 else ('medium-score' if total_score >= 80 else 'low-score')
            
            html_content += f"""
                <tr>
                    <td>{mentor.getName}</td>
                    <td>{mentor.getMajor}</td>
                    <td>{mentee.getName}</td>
                    <td>{mentee.getMajor}</td>
                    <td class="{score_class}">{total_score}/195</td>
                    <td>{mentor.getEmail}</td>
                    <td>{mentee.getEmail}</td>
                </tr>
            """
    
    html_content += """
                </tbody>
            </table>
        </div>
    """
    
    # Tab 2: By Mentor
    html_content += """
        <div id="mentors" class="tab-content">
            <h2>Mentors and Their Mentees</h2>
            <table id="mentors-table">
                <thead>
                    <tr>
                        <th>Mentor Name</th>
                        <th>Email</th>
                        <th>Discord</th>
                        <th>Major</th>
                        <th>Mentees</th>
                        <th>Mentee Names</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for mentor in mentors:
        mentee_list = finalPairings.get(mentor.getName, [])
        mentee_names = '<br>'.join(mentee_list) if mentee_list else 'None'
        
        html_content += f"""
            <tr>
                <td>{mentor.getName}</td>
                <td>{mentor.getEmail}</td>
                <td>{mentor.getDiscord}</td>
                <td>{mentor.getMajor}</td>
                <td>{len(mentee_list)}</td>
                <td>{mentee_names}</td>
            </tr>
        """
    
    html_content += """
                </tbody>
            </table>
        </div>
    """
    
    # Tab 3: By Mentee
    html_content += """
        <div id="mentees" class="tab-content">
            <h2>Mentees and Their Mentors</h2>
            <table id="mentees-table">
                <thead>
                    <tr>
                        <th>Mentee Name</th>
                        <th>Email</th>
                        <th>Discord</th>
                        <th>Major</th>
                        <th># Mentors</th>
                        <th>Mentor Names</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for mentee in mentees:
        mentor_list = paired_mentees.get(mentee.getName, [])
        mentor_names = '<br>'.join(mentor_list) if mentor_list else 'None'
        num_mentors = len(mentor_list)
        status = 'Fully Paired' if num_mentors == 2 else ('Partially Paired' if num_mentors == 1 else 'Unpaired')
        status_class = 'high-score' if num_mentors == 2 else ('medium-score' if num_mentors == 1 else 'low-score')
        
        html_content += f"""
            <tr>
                <td>{mentee.getName}</td>
                <td>{mentee.getEmail}</td>
                <td>{mentee.getDiscord}</td>
                <td>{mentee.getMajor}</td>
                <td>{num_mentors}/2</td>
                <td>{mentor_names}</td>
                <td class="{status_class}">{status}</td>
            </tr>
        """
    
    html_content += """
                </tbody>
            </table>
        </div>
    """
    
    # Tab 4: Unpaired with Top Candidates
    html_content += """
        <div id="unpaired" class="tab-content">
            <h2>Mentees Needing Pairing</h2>
            <table id="unpaired-table">
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>Email</th>
                        <th>Major</th>
                        <th>Current Mentors</th>
                        <th>Top 5 Mentor Candidates</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for mentee in mentees:
        mentor_count = len(paired_mentees.get(mentee.getName, []))
        if mentor_count < 2:
            current_mentors = ', '.join(paired_mentees.get(mentee.getName, [])) or 'None'
            
            # Calculate compatibility scores for all available mentors
            mentor_scores = calcScore_Mentee(mentee, mentors)
            
            # Sort by score and get top 5
            sorted_candidates = sorted(mentor_scores.items(), key=lambda x: x[1], reverse=True)[:5]
            
            # Build HTML for candidates
            candidates_html = '<div class="candidate-list">'
            for i, (mentor_name, score) in enumerate(sorted_candidates, 1):
                # Find the mentor object
                mentor = next((m for m in mentors if m.getName == mentor_name), None)
                if not mentor:
                    continue
                
                # Check if mentor has capacity
                current_mentees = len(finalPairings.get(mentor_name, []))
                max_mentees = mentor.getNumMentees
                has_capacity = (max_mentees == 0) or (current_mentees < max_mentees)
                
                # Determine score color
                score_class = 'high-score' if score >= 130 else ('medium-score' if score >= 80 else 'low-score')
                
                # Add capacity badge
                capacity_badge = ''
                if has_capacity:
                    spots_left = 'unlimited' if max_mentees == 0 else str(max_mentees - current_mentees)
                    capacity_badge = f'<span class="available-badge">✓ Available ({spots_left} spots)</span>'
                else:
                    capacity_badge = '<span class="at-capacity-badge">At Capacity</span>'
                
                candidates_html += f'''
                    <div class="candidate-item">
                        {i}. <strong>{mentor_name}</strong> ({mentor.getMajor})
                        - <span class="{score_class}">Score: {score}/195</span>
                        {capacity_badge}
                    </div>
                '''
            
            candidates_html += '</div>'
            
            html_content += f"""
                <tr>
                    <td><strong>{mentee.getName}</strong></td>
                    <td>{mentee.getEmail}</td>
                    <td>{mentee.getMajor}</td>
                    <td>{mentor_count}/2 ({current_mentors})</td>
                    <td>{candidates_html}</td>
                </tr>
            """
    
    html_content += """
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"\n✓ HTML report exported to: {output_path}")
    print(f"  Open in browser: file://{output_path}")
    return output_path
# ==========================================
# 4. MAIN EXECUTION BLOCK
# ==========================================

# 1. Initialize Users (Only call this ONCE)
initializeUsers()

# 2. Calculate All Scores
completePairings = initializeFinalMap()

# 3. Export the Reports
# Mentor Ranking CSV
export_rankings_to_csv(mentors, completePairings)
# Mentee Ranking CSV
export_mentee_rankings_to_csv(mentees, mentors)
# mentor mentee score matrix
export_compatibility_matrix(mentors, mentees)


# 4. Run the matching logic (The "Fairness" Algorithm)
finalPairings, unpaired, partially_paired, paired_mentees = pairing(completePairings)

# 5. Export to HTML
html_path = export_to_html(finalPairings, mentors, mentees, paired_mentees)

# 6. Open the HTML report
webbrowser.open('file://' + html_path)