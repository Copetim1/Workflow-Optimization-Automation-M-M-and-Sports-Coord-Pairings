#  M&M Pairing System

An mentorship matching algorithm that pairs mentees with mentors based on compatibility scores all around including major, MBTI personality type, professional interests, and involvement preferences.

##  Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Compatibility Scoring](#compatibility-scoring)
- [Output](#output)
- [File Structure](#file-structure)
- [Configuration](#configuration)

##  Overview

This system automates the process of matching mentees with mentors in a student organization (SASE - Society of Asian Scientists and Engineers). It analyzes survey responses and creates optimal pairings based on:

- Academic majors and focus areas
- MBTI personality compatibility
- Professional development interests
- Group involvement preferences
- M&M Cup participation enthusiasm
- User-requested pairings

Each mentee can be paired with up to **2 mentors**.

##  Features

### Intelligent Matching
- **Multi-dimensional compatibility scoring** (max 195 points)
- **Major-based matching** with similar major groupings
- **MBTI personality compatibility** using research-based compatibility matrices
- **Professional interest alignment**
- **Involvement level matching** for group activities and M&M Cup

### Flexible Pairing
- Supports user-requested mentor/mentee pairings
- Respects mentor capacity preferences
- Allows mentees to have up to 2 mentors
- Handles major preference requirements

### Comprehensive Reporting
- **Interactive HTML dashboard** with:
  - Statistics overview
  - Tabbed interface for different views
  - Sortable tables
  - Color-coded compatibility scores
  - Top 5 mentor candidates for unpaired mentees
  - Availability indicators for mentors

### Auto-Opens in Browser
- Automatically launches the HTML report after generation
- Easy sharing and viewing

##  Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager

### Required Packages

```bash
pip install pandas openpyxl --break-system-packages
```

### File Setup

1. Place your Excel file in the `files/` directory
2. Name it `Spring_2026_Responses.xlsx` or update the `file_path` variable
3. Ensure the Excel file has the required columns (see Configuration)

## Usage

### Basic Usage

Run the main script:

```bash
python main.py
```

This will:
1. Load and process the survey data
2. Calculate compatibility scores
3. Create optimal pairings
4. Generate an HTML report
5. Automatically open the report in your browser

### Custom Configuration

You can modify the following in the code:

```python
# Change the input file
file_path = os.path.join(script_dir, 'files', 'YOUR_FILE.xlsx')

# Change the output filename
export_to_html(finalPairings, mentors, mentees, paired_mentees, 
               filename='custom_report.html')
```

## 🔧 How It Works

### Phase 1: Data Processing
1. **Load Excel data** from survey responses
2. **Create User objects** (Mentor or Mentee)
3. **Extract and normalize** survey answers
4. **Categorize majors** into similar groups

### Phase 2: Compatibility Calculation
For each mentor-mentee pair, calculate scores across:
- Major compatibility (70 points max)
- Professional help alignment (55 points max)
- MBTI compatibility (30 points max)
- Involvement preferences (40 points max)

**Total possible score: 195 points**

### Phase 3: Pairing Algorithm
1. **User Requests (Priority 1)**
   - Process mentee-requested mentors
   - Process mentor-requested mentees

2. **Compatibility-Based Pairing (Priority 2)**
   - Sort mentors by available capacity
   - For each mentor, pair with highest-scoring available mentees
   - Respect mentor capacity limits
   - Allow mentees to receive up to 2 mentors

### Phase 4: Report Generation
- Generate interactive HTML dashboard
- Calculate statistics
- Display all pairings with scores
- Show top candidates for unpaired mentees

## Compatibility Scoring

### Major Score (70 points max)
- **Same major**: 70 points
- **Same academic focus**: 20 points
- **Similar major**: 25 points

**Major Groups:**
- Engineering (ME, AE, EE, CE, ChemE, CompE, NucE, ISE, BME, DAS)
- Computer Science (CS, Info Systems, Data Science, DAS)
- Business (BA, Finance, Accounting, Marketing, Economics)
- Pre-Health (Health Science, BME, Microbiology, Chemistry, Biochemistry, Biology, Public Health)

### Professional Help Score (55 points max)
- 5 points per matching professional interest
- 5 points per matching "perfect relationship" descriptor

### MBTI Score (30 points max)
Based on Myers-Briggs compatibility research:
- **Best match**: 30 points (e.g., INTJ ↔ ENFP)
- **Good match**: 25 points
- **Okay match**: 15 points
- **Challenging match**: 10 points

### Involvement Score (40 points max)
- **Group involvement alignment**: up to 20 points
- **M&M Cup enthusiasm alignment**: up to 20 points
- Exact match: 20 points
- Within 2 levels: 15 points
- Within 5 levels: 10 points

## Output

### HTML Report Features

**Statistics Dashboard:**
- Total mentors and mentees
- Total pairings made
- Fully paired mentees (2 mentors)
- Partially paired mentees (1 mentor)
- Unpaired mentees

**Four Interactive Tabs:**

1. **All Pairings**
   - Complete list of mentor-mentee pairs
   - Compatibility scores (color-coded)
   - Contact information

2. **By Mentor**
   - Each mentor with their assigned mentees
   - Mentor capacity tracking

3. **By Mentee**
   - Each mentee with their assigned mentors
   - Pairing status

4. **Needs Pairing**
   - Unpaired/partially paired mentees
   - **Top 5 mentor candidates** for each
   - Availability status (spots remaining)
   - Compatibility scores

### Color Coding
- **Green (High)**: Score ≥ 130/195
- **Orange (Medium)**: Score 80-129/195
- **Red (Low)**: Score < 80/195

## 📁 File Structure

```
mentorship-pairing/
├── main.py                  # Main script with pairing logic
├── user.py                  # Base User class
├── mentor.py                # Mentor class
├── mentee.py                # Mentee class
├── README.md                # This file
└── files/
    ├── Spring_2026_Responses.xlsx    # Input survey data
    └── mentorship_report.html        # Generated report
```

## ⚙️ Configuration

### Required Excel Columns

Your Excel file must include these columns:

**Basic Information:**
- `Name (First Last)`
- `Email address`
- `Year`
- `Phone Number`
- `Discord (e.g. jellyduck224)`

**Academic:**
- `What best describes your academic focus?`
- `What is your major? (Pre-health)`
- `What is your major? (Engineering)`
- `What is your major? (Other)`
- `What is your major? (Postgrad)`
- `Do you want your mentor/mentee to be the same major?`

**Role:**
- `I want to be ...` (values: "a Mentee!" or "a Mentor!")

**For Mentors:**
- `Do you have a preference on how many mentees you receive?`
- `To encourage interaction and assist with bonding between Mentors and Mentees...`

**Survey Questions:**
- `What student organizations are you currently involved/want to be more involved with? (Mentee/Mentor)`
- `What industries/companies are you interested in working in...`
- `What type of professional help do you want to receive/provide?`
- `What is your MBTI Personality type? (Mentee/Mentor)`
- `Describe your perfect relationship with your mentor/mentee.`
- `How involved do you want to be with your M&M group? (Mentee/Mentor)`
- `How much do you want to win the M&M Cup? (Mentee/Mentor)`
- `Do you have a specific mentor/mentee you would like to request?`

### Adding New Major Groups

Edit the `isSimilarMajor()` function in `main.py`:

```python
similarMajors = {
    'your_group': ['Major 1', 'Major 2', 'Major 3'],
    # Add more groups...
}
```

### Adjusting Score Weights

Modify the scoring functions:
- `assignScoreMajor()` - Currently 70 points max
- `assignScoreProfHelpPerfRelationship()` - Currently 55 points max
- `assignScoreMBTI()` - Currently 30 points max
- `assignScoreInvolvement()` - Currently 40 points max

## Contributing

To improve the pairing algorithm:

1. Adjust scoring weights based on pairing success feedback
2. Add additional compatibility dimensions
3. Enhance the MBTI compatibility matrix
4. Expand similar major groupings
5. Improve the HTML report design

## Notes

- The system prioritizes user-requested pairings over algorithmic matches
- Mentors can specify their capacity (0 = unlimited)
- Mentees are filtered out if they're the same year or older than potential mentors
- Major preference settings are respected (both mentor and mentee must agree if either requires same major)
- The algorithm aims to maximize compatibility while respecting constraints

## Troubleshooting

**Issue: "No module named 'pandas'"**
```bash
pip install pandas openpyxl --break-system-packages
```

**Issue: "File not found"**
- Check that your Excel file is in the `files/` directory
- Verify the filename matches `Spring_2026_Responses.xlsx`

**Issue: "No mentees paired"**
- Check that year filtering isn't too restrictive
- Verify mentors have capacity (not all set to 0 and already full)
- Review major preference requirements

**Issue: HTML report won't open**
- Check the console for the file path
- Manually open the file in `files/mentorship_report.html`
- Try a different browser if auto-open fails

## Support

For questions or issues with the pairing system, contact wangt2@ufl.edu

---

**Made for the SASE M&M Program** 🎓