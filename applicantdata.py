# -*- coding: utf-8 -*-
"""ApplicantData.ipynb

This notebook uses python to generate a random list of applicants to assist in generating candidate records for testing an ERP system prior to go live. This process could be improved with additional data instead of hard coded random lists. This data was created to generate word and pdf resumes to be uploaded alongside candidate records. View the ResumeGenerator.ipynb notebook once the test set is generated here to generate your own resumes.
"""

#@title Install required packages
!pip -q install pandas openpyxl

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
from pathlib import Path

#@title Configuration (paths, switches, seeds)

# ---------- OUTPUT ARTIFACTS ----------
ART_DIR = Path("artifacts")
ART_DIR.mkdir(exist_ok=True)

# Iteration artifacts
OUT_POSITIONS_XLSX   = ART_DIR / "01_positions_seed.xlsx"
OUT_APPS_XLSX        = ART_DIR / "02_applications.xlsx"
OUT_WORK_DIVERSE_XLSX= ART_DIR / "03_applications_diverse_workhistory.xlsx"
OUT_ACH_V2_XLSX      = ART_DIR / "04_applications_achievements_v2.xlsx"
OUT_SENIOR_V3_XLSX   = ART_DIR / "05_applications_seniorhistory_v3.xlsx"

# ---------- RANDOM SEED ----------
random.seed(456)  # fixed seed for reproducibility
np.random.seed(456)

# Generate list of open positions scraped from colleges external career site

positions_seed = [
    {"Title":"Head Coach, Field Hockey","Category":"Coaching","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"On-site","Req":"HEADC009823"},
    {"Title":"Associate Director, Upward Bound","Category":"Student Program Administrator","Campus":"FSU - Brighton","Schedule":"Full Time","Type":"On-site","Req":"ASSOC009821"},
    {"Title":"Event Management Adjunct Faculty","Category":"Instructional","Campus":"FSU - Camden","Schedule":"Part Time","Type":"On-site","Req":"AUDIO009820"},
    {"Title":"Part Time Faculty – ELT 3050L Microcontroller Techniques II","Category":"Instructional","Campus":"FSU - Dover","Schedule":"Part Time","Type":"On-site","Req":"PARTT009793"},
    {"Title":"Assistant Professor of Restorative Justice","Category":"Academic Instruction","Campus":"FSU - Camden","Schedule":"Full Time","Type":"On-site","Req":"ASSIS009815"},
    {"Title":"Registration Services Specialist","Category":"Administration","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"On-site","Req":"REGIS009814"},
    {"Title":"Chief Business Officer","Category":"Executive Leadership - Unclassified","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"On-site","Req":"CHIEF009813"},
    {"Title":"Associate Campus Operations Director","Category":"Administration","Campus":"FSU - Brighton","Schedule":"Full Time","Type":"On-site","Req":"ASSOC009812"},
    {"Title":"Director Clinical Education","Category":"Program Administration","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"Hybrid","Req":"DIREC009817"},
    {"Title":"Workday Program Director for Student Implementation","Category":"IT Management","Campus":"FSU - Dover","Schedule":"Full Time","Type":"Remote Eligible","Req":"WORKD009816"},
    {"Title":"Business Advisor","Category":"Economic and Business Development","Campus":"FSU - Elmwood","Schedule":"Full Time","Type":"Hybrid","Req":"BUSIN009807"},
    {"Title":"Clinical Education Coordinator, Associate Degree Nursing","Category":"Program Administration","Campus":"FSU - Elmwood","Schedule":"Full Time","Type":"Hybrid","Req":"CLINI009806"},
    {"Title":"Student Success Assistant","Category":"Student Services","Campus":"FSU - Elmwood","Schedule":"Full Time","Type":"On-site","Req":"STUDE009808"},
    {"Title":"Head Athletic Trainer","Category":"Student Services","Campus":"FSU - Camden","Schedule":"Full Time","Type":"On-site","Req":"HEADA009704"},
    {"Title":"University Counselor","Category":"Administration","Campus":"FSU - Brighton","Schedule":"Part Time","Type":"On-site","Req":"UNIVE009598"},
    {"Title":"Full-time Faculty – PN/AD Nursing Programs (Dover/Hartwell)","Category":"Instructional","Campus":"FSU - Dover/Hartwell","Schedule":"Full Time","Type":"On-site","Req":"FULLT009684"},
    {"Title":"Part-time – Dental Hygiene Program Clinical Instructor (Dover)","Category":"Instructional","Campus":"FSU - Dover","Schedule":"Part Time","Type":"On-site","Req":"ASSIS009668"},
    {"Title":"Radiologic Sciences Adjunct Faculty (Specialized Imaging)","Category":"Instructional","Campus":"FSU - Dover","Schedule":"Part Time","Type":"On-site","Req":"PARTT009640"},
    {"Title":"Enterprise Systems Administrator","Category":"Infrastructure","Campus":"SO - Fairview","Schedule":"Full Time","Type":"On-site","Req":"ENTER009789"},
    {"Title":"Part-Time Faculty – ARE 2052 Architectural Design II","Category":"Academic Instruction","Campus":"FSU - Elmwood","Schedule":"Part Time","Type":"On-site","Req":"PARTT009756"},
    {"Title":"Part-Time Faculty – CIS 1152L Advanced Website Development – Lab","Category":"Academic Instruction","Campus":"FSU - Dover","Schedule":"Part Time","Type":"On-site","Req":"PARTT009752"},
    {"Title":"Part-Time Faculty – ELT 1110 Introduction to Digital Circuits","Category":"Instructional","Campus":"FSU - Elmwood","Schedule":"Part Time","Type":"On-site","Req":"PARTT009786"},
    {"Title":"Part-Time Faculty – ELT 1110L Introduction to Digital Circuits – Lab (Elmwood)","Category":"Instructional","Campus":"FSU - Elmwood","Schedule":"Part Time","Type":"On-site","Req":"PARTT009788"},
    {"Title":"Part-Time Faculty of Theatre Arts (Ashford)","Category":"Instructional","Campus":"FSU - Ashford","Schedule":"Part Time","Type":"On-site","Req":"PARTT009781"},
    {"Title":"Senior Mechanical Systems Technician","Category":"Facilities","Campus":"FSU - Elmwood","Schedule":"Full Time","Type":"On-site","Req":"SENIO009771"},
    {"Title":"Custodian/Housekeeper III (Ashford)","Category":"Facilities","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"On-site","Req":"CUSTO009744"},
    {"Title":"Student Financial Support Specialist","Category":"Student Services","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"On-site","Req":"STUDE009718"},
    {"Title":"Part-Time Faculty – PHY 2010 Electricity and Electronics (Camden)","Category":"Instructional","Campus":"FSU - Camden","Schedule":"Part Time","Type":"On-site","Req":"PARTT009572"},
    {"Title":"Interim Assistant Director of Student Activities","Category":"Student Program Administrator","Campus":"FSU - Elmwood","Schedule":"Full Time","Type":"On-site","Req":"INTER009804"},
    {"Title":"FSU Online Adjunct Instructor: HOH2060 Naturopathic Medicine","Category":"Academic Instruction","Campus":"FSU - Online (Brighton)","Schedule":"Part Time","Type":"Remote","Req":"FSUOL009801"},
    {"Title":"Part-Time – Paramedic Certificate Program (Online)","Category":"Instructional","Campus":"FSU - Online (Dover)","Schedule":"Part Time","Type":"Remote","Req":"PARTT009805"},
    {"Title":"Assistant / Associate Professor in Nursing – Ashford (TT)","Category":"Instructional","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"On-site","Req":"ASSIS009517"},
    {"Title":"Assistant / Associate Professor in Nursing – Ashford (TT)","Category":"Instructional","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"On-site","Req":"ASSIS009516"},
    {"Title":"Full-time Faculty – PN/AD Nursing Programs – Elmwood","Category":"Instructional","Campus":"FSU - Elmwood","Schedule":"Full Time","Type":"On-site","Req":"FULLT009521"},
    {"Title":"Assistant / Associate Professor in Nursing – Ashford","Category":"Instructional","Campus":"FSU - Ashford","Schedule":"Full Time","Type":"On-site","Req":"FULLT009514"},
    {"Title":"Lifeguard","Category":"Student Services","Campus":"FSU - Ashford","Schedule":"Part Time","Type":"On-site","Req":"LIFEG009624"},
]

positions_df = pd.DataFrame(positions_seed)
positions_df.to_excel(OUT_POSITIONS_XLSX, sheet_name="Positions", index=False)
positions_df.head(), OUT_POSITIONS_XLSX

#@title Applicant generation (4–6 per position) with category-specific locality mix & scorecards

# Locality preferences by category (overall target ~60–70% local)
LOCAL_PREF = {
    "Student Services": 0.80, "Administration": 0.80, "Department Administrative Assistants": 0.80,
    "Student Program Administrator": 0.75, "Program Administration": 0.70, "Coaching": 0.75,
    "Facilities": 0.75, "Infrastructure": 0.60, "Executive Leadership - Unclassified": 0.55,
    "IT Management": 0.55, "Academic Instruction": 0.55, "Instructional": 0.55,
    "Economic and Business Development": 0.65,
}
SOURCES = ["University Careers Site","LinkedIn","Indeed","Employee Referral","Professional Association"]

LOCAL_CITIES = [
    ("Bennington","VT"),("Castleton","VT"),("Johnson","VT"),("Lyndonville","VT"),("Randolph Center","VT"),
    ("Williston","VT"),("St. Albans","VT"),("Montpelier","VT"),("Burlington","VT"),("Rutland","VT"),
    ("Concord","NH"),("Keene","NH"),("Lebanon","NH"),("Portland","ME"),("Amherst","MA"),("Northampton","MA"),
    ("Brattleboro","VT"),("Middlebury","VT"),("Stowe","VT"),("Winooski","VT")
]
NONLOCAL_CITIES = [
    ("Albany","NY"),("Syracuse","NY"),("Pittsburgh","PA"),("Rochester","NY"),("Philadelphia","PA"),("Baltimore","MD"),
    ("Columbus","OH"),("Chicago","IL"),("Boulder","CO"),("Madison","WI"),("Atlanta","GA"),("Raleigh","NC"),
    ("Austin","TX"),("Seattle","WA"),("San Diego","CA")
]

CATEGORY_SKILLS = {
    "Coaching": ["Team leadership","Recruiting","Game strategy","NCAA compliance","Athlete development","CPR/AED"],
    "Student Program Administrator": ["Program coordination","Budget tracking","Grant reporting","Student engagement","Event planning"],
    "Instructional": ["Course design","Assessment","Canvas LMS","Student mentoring","Active learning"],
    "Academic Instruction": ["Curriculum design","Scholarly activity","Assessment","Advising","Learning analytics"],
    "Administration": ["Office management","Customer service","Records","Scheduling","Policy"],
    "Executive Leadership - Unclassified": ["Strategic planning","Finance","Operations","Governance","Change management"],
    "Program Administration": ["Project management","Stakeholder coordination","Clinical placement","SOPs","Quality assurance"],
    "Economic and Business Development": ["Business advising","Market analysis","Financial modeling","Outreach","Client services"],
    "Student Services": ["Advising","Retention","Case management","Health & wellness","Athletics"],
    "Infrastructure": ["Linux/Windows admin","VMware","Storage","Backup/DR","Monitoring","Scripting"],
    "IT Management": ["Workday Student","Data governance","Change management","Program leadership","Stakeholder alignment"],
    "Facilities": ["CMMS","Preventive maintenance","Safety compliance","Vendor management","HVAC"]
}
CATEGORY_CERTS = {
    "Coaching": ["CPR/AED"], "Student Program Administrator": ["Mental Health First Aid"],
    "Instructional": ["Train-the-Trainer"], "Academic Instruction": ["IRB Training"],
    "Administration": ["Notary Public"], "Executive Leadership - Unclassified": ["PMP"],
    "Program Administration": ["PMP","CITI - Responsible Conduct of Research"],
    "Economic and Business Development": ["Certified Business Advisor"],
    "Student Services": ["QPR Gatekeeper","Mental Health First Aid"],
    "Infrastructure": ["ITIL Foundation","CompTIA Security+"],
    "IT Management": ["PMP","Prosci Change Practitioner"],
    "Facilities": ["OSHA 30","First Aid/CPR"],
}

def choose_local(category: str) -> bool:
    p = LOCAL_PREF.get(category, 0.66)
    return random.random() < p

def score_for_category(category: str) -> dict:
    base = {
        "Role Fit": float(np.clip(np.random.normal(3.4, 0.7), 1, 5)),
        "Experience": float(np.clip(np.random.normal(3.6, 0.8), 1, 5)),
        "Communication": float(np.clip(np.random.normal(3.8, 0.6), 1, 5)),
        "Culture Add": float(np.clip(np.random.normal(3.7, 0.6), 1, 5)),
    }
    if category in ("Infrastructure","IT Management"):
        base["Technical"] = float(np.clip(np.random.normal(3.6, 0.8), 1, 5))
    if category in ("Academic Instruction","Instructional"):
        base["Teaching Demo"] = float(np.clip(np.random.normal(3.5, 0.7), 1, 5))
    if category in ("Student Services","Student Program Administrator"):
        base["Advising & Retention"] = float(np.clip(np.random.normal(3.7, 0.6), 1, 5))
    if category in ("Executive Leadership - Unclassified",):
        base["Leadership"] = float(np.clip(np.random.normal(3.8, 0.6), 1, 5))
    if category in ("Facilities",):
        base["Safety & Compliance"] = float(np.clip(np.random.normal(3.6, 0.7), 1, 5))
    return base

def stage_and_outcome(avg_score: float, app_date: datetime.date):
    stages = ["Submitted","Phone Screen","Hiring Manager","Campus Panel","Final"]
    if avg_score >= 4.0:
        stage = random.choice(stages[2:])
    elif avg_score >= 3.3:
        stage = random.choice(stages[1:4])
    else:
        stage = random.choice(stages[:3])
    outcome = "Proceed" if avg_score >= 3.6 else ("Hold" if avg_score >= 3.0 else "Decline")
    interview_date = None
    if stage != "Submitted":
        interview_date = (datetime.combine(app_date, datetime.min.time()) + timedelta(days=random.randint(3,21))).date()
    return stage, outcome, interview_date

# Applicant generation
apps_rows = []
app_id_counter = 1
today = datetime(2025,12,12).date()

for _, pos in positions_df.iterrows():
    count = random.randint(4,6)  # 4–6 per posting
    category = pos["Category"]
    skills = CATEGORY_SKILLS.get(category, ["Communication","Teamwork","Problem solving"])
    certs = CATEGORY_CERTS.get(category, [])
    for _ in range(count):
        is_local = choose_local(category)
        city, state = random.choice(LOCAL_CITIES if is_local else NONLOCAL_CITIES)
        fname, lname = random.choice(["Avery","Jordan","Taylor","Morgan","Casey","Riley","Quinn","Emerson","Alex","Cameron","Jamie","Hayden","Rowan","Reese","Parker","Sydney","Logan"]), \
                       random.choice(["Smith","Johnson","Williams","Brown","Jones","Miller","Davis","Garcia","Rodriguez","Martinez","Anderson","Taylor","Thomas","Hernandez","Moore","Martin","Jackson","Thompson"])
        email = f"{fname.lower()}.{lname.lower()}@applicantmail.com"
        phone = f"(802) {random.randint(200,999)}-{random.randint(1000,9999)}"
        app_date = (today - timedelta(days=random.randint(0,60)))
        summary = f"{fname} {lname} brings {random.randint(3,18)} years' experience in {category.lower()} for {pos['Title']}. Strengths: {', '.join(random.sample(skills, k=min(3,len(skills))))}."
        resume_file = f"{lname}_{fname}_Resume.pdf"
        source = random.choices(SOURCES, weights=[0.4,0.25,0.2,0.15,0.1], k=1)[0]
        scores = score_for_category(category)
        avg_score = round(sum(scores.values()) / len(scores), 2)
        stage, outcome, interview_date = stage_and_outcome(avg_score, app_date)

        apps_rows.append({
            "ApplicantID": f"A{app_id_counter:04d}",
            "Full Name": f"{fname} {lname}",
            "Preferred Name": fname,
            "Email": email,
            "Phone": phone,
            "City": city,
            "State": state,
            "Local": "Yes" if is_local else "No",
            "Campus": pos["Campus"],
            "Position Applied For": pos["Title"],
            "Job Category": category,
            "Schedule": pos["Schedule"],
            "Job Location Type": pos["Type"],
            "Requisition Number": pos["Req"],
            "Source": source,
            "Resume FileName": resume_file,
            "Application Date": app_date,
            "Hiring Stage": stage,
            "Interview Date": interview_date,
            "Screening Outcome": outcome,
            "Resume Summary": summary,
            "Key Skills": ", ".join(random.sample(skills, k=min(5, len(skills)))),
            "Certifications": ", ".join(random.sample(certs, k=min(2, len(certs)))),
            **scores,
            "Average Score": avg_score
        })
        app_id_counter += 1

apps_df = pd.DataFrame(apps_rows)
with pd.ExcelWriter(OUT_APPS_XLSX, engine="openpyxl") as w:
    positions_df.to_excel(w, sheet_name="Positions", index=False)
    apps_df.to_excel(w, sheet_name="Applications", index=False)

# Quick summary
local_ratio = (apps_df["Local"] == "Yes").mean()
print(f"Generated {len(apps_df)} applicants across {len(positions_df)} postings. Local ratio ≈ {round(local_ratio,2)}")
OUT_APPS_XLSX

#@title Build diversified work history: USA/Canada universities + New England employers

USA_UNIS = [
    "Harvard University","Massachusetts Institute of Technology","Boston University","Northeastern University",
    "Brown University","Dartmouth College","University of Vermont","University of Massachusetts Amherst",
    "University of Connecticut","Stanford University","University of California, Berkeley","UCLA","UC San Diego",
    "University of Washington","University of Michigan","Ohio State University","Penn State University","UW–Madison",
    "UNC Chapel Hill","Duke University","University of Virginia","Georgia Tech","Purdue University","Indiana University",
    "University of Minnesota","University of Colorado Boulder","Arizona State University","UT Austin","Rice University","UIUC",
    "University of Rochester","RPI","Syracuse University","Carnegie Mellon University","Princeton University","Columbia University",
    "Cornell University","Yale University","University of Pennsylvania"
]
CAN_UNIS = [
    "University of Toronto","McGill University","University of British Columbia","University of Waterloo",
    "Queen’s University","Western University","University of Calgary","University of Alberta","McMaster University",
    "Dalhousie University","University of Ottawa","Simon Fraser University","Concordia University","Université Laval",
    "University of Victoria","University of Manitoba","University of Saskatchewan","York University","Carleton University"
]
NE_EMPLOYERS = [
    "Mass General Brigham","Boston Children’s Hospital","Biogen","Vertex Pharmaceuticals","Fidelity Investments",
    "Liberty Mutual","Wayfair","Bose Corporation","HubSpot","Eversource Energy","National Grid","Keurig Dr Pepper",
    "Pratt & Whitney","Raytheon Technologies","Ben & Jerry’s","GlobalFoundries","Dealer.com",
    "UVM Health Network","Dartmouth Health","MaineHealth","Connecticut Children’s",
    "VT Agency of Human Services","VT Agency of Education","VT Dept. of Labor","NH Dept. of Education",
    "RI Dept. of Health","Yale New Haven Health","UMass Memorial Health"
]

achievements_pool = [
    "Improved student persistence by 4.2pp via targeted outreach",
    "Reduced processing time by 28% through workflow automation",
    "Launched new course sequence with 95% pass rate",
    "Implemented analytics dashboard adopted by 6 departments",
    "Secured $300k grant for student success initiatives",
    "Optimized vendor contracts saving 11% annually",
    "Achieved 99.9% systems uptime across critical services",
    "Led accreditation preparation across multi-campus unit",
    "Expanded community partnerships resulting in 150+ placements",
    "Introduced CMMS preventive program reducing incidents by 18%",
]

work_rows = []
for aid in apps_df["ApplicantID"].unique():
    # Two prior roles per applicant
    app_date = pd.to_datetime(apps_df.loc[apps_df["ApplicantID"] == aid, "Application Date"].iloc[0]).date()
    # University role
    uni_country = random.choice(["USA","Canada"])
    org_uni = random.choice(USA_UNIS if uni_country=="USA" else CAN_UNIS)
    end1 = app_date - timedelta(days=random.randint(120,540))
    start1 = end1 - timedelta(days=random.randint(500,1200))
    work_rows.append({
        "ApplicantID": aid,
        "Organization": org_uni,
        "Organization Type": "University",
        "Country": uni_country,
        "Region": "National" if uni_country=="USA" else "Canada",
        "Role": random.choice(["Adjunct Instructor","Lecturer","Program Coordinator","Advisor","Systems Administrator","Student Success Advisor"]),
        "Start Date": start1,
        "End Date": end1,
        "Achievements": random.choice(achievements_pool)
    })
    # New England employer role
    org_ne = random.choice(NE_EMPLOYERS)
    end2 = start1 - timedelta(days=random.randint(90,360))
    start2 = end2 - timedelta(days=random.randint(500,1200))
    work_rows.append({
        "ApplicantID": aid,
        "Organization": org_ne,
        "Organization Type": "Employer",
        "Country": "USA",
        "Region": "New England",
        "Role": random.choice(["Administrative Specialist","Office Manager","Registration Coordinator","Program Manager","Mechanical Systems Technician","Business Advisor"]),
        "Start Date": start2,
        "End Date": end2,
        "Achievements": random.choice(achievements_pool)
    })

work_df = pd.DataFrame(work_rows).sort_values(["ApplicantID","Start Date"])
with pd.ExcelWriter(OUT_WORK_DIVERSE_XLSX, engine="openpyxl") as w:
    positions_df.to_excel(w, sheet_name="Positions", index=False)
    apps_df.to_excel(w, sheet_name="Applications", index=False)
    work_df.to_excel(w, sheet_name="ApplicantWorkHistory", index=False)

print(f"Diverse work history created: {len(work_df)} rows, unique orgs: {work_df['Organization'].nunique()}")
OUT_WORK_DIVERSE_XLSX

#@title Update only the Achievements column with diverse statements (no reordering)

ACH_NEW_POOL = [
    "Developed cross-campus mentorship program increasing engagement by 22%",
    "Implemented cloud-based scheduling system reducing conflicts by 35%",
    "Negotiated vendor contracts saving $150K annually",
    "Designed and launched diversity training adopted by 500+ staff",
    "Introduced predictive analytics for enrollment boosting yield by 8%",
    "Led accreditation review resulting in zero compliance findings",
    "Built mobile app for student services with 4.8/5 user rating",
    "Secured multi-year research grant of $500K for STEM initiatives",
    "Streamlined onboarding process cutting time-to-productivity by 40%",
    "Piloted hybrid learning model increasing course completion by 15%",
    "Established wellness program reducing absenteeism by 12%",
    "Automated financial aid workflows improving turnaround by 30%",
    "Expanded internship partnerships adding 75 new employer sites",
    "Developed cybersecurity protocols reducing incidents by 90%",
    "Created faculty development series attended by 200+ participants",
    "Implemented energy efficiency upgrades lowering costs by 18%",
    "Launched alumni engagement campaign raising $1.2M in donations",
    "Introduced competency-based assessments improving retention by 10%",
    "Managed relocation project completing ahead of schedule and under budget",
    "Designed virtual advising platform serving 3,000 students annually",
]

work_v2 = work_df.copy()
used = set()
new_vals = []
for i in range(len(work_v2)):
    for _ in range(10):
        s = random.choice(ACH_NEW_POOL)
        if s not in used:
            used.add(s)
            new_vals.append(s)
            break
    else:
        new_vals.append(f"Delivered multi-year initiative; outcomes improved {random.randint(10,40)}% (row {i})")
work_v2["Achievements"] = new_vals

with pd.ExcelWriter(OUT_ACH_V2_XLSX, engine="openpyxl") as w:
    positions_df.to_excel(w, sheet_name="Positions", index=False)
    apps_df.to_excel(w, sheet_name="Applications", index=False)
    work_v2.to_excel(w, sheet_name="ApplicantWorkHistory", index=False)

print(f"Achievements updated (unique): {len(set(work_v2['Achievements']))} / {len(work_v2)}")
OUT_ACH_V2_XLSX

#@title Extend senior applicants so earliest start <= 2005 (append older roles where needed)

def is_senior_applicant(row: pd.Series) -> bool:
    senior_categories = {"Executive Leadership - Unclassified","IT Management","Program Administration","Facilities"}
    senior_keywords = ("chief","director","program director","head","senior","manager")
    cat = str(row.get("Job Category","")).strip()
    title = str(row.get("Position Applied For","")).lower()
    return (cat in senior_categories) or any(k in title for k in senior_keywords)

ROLE_TITLE_MAP = {
    "Executive Leadership - Unclassified": ["Executive Director","Division Director","Chief of Staff"],
    "IT Management": ["Program Director","Project Manager","Change Manager"],
    "Program Administration": ["Program Manager","Quality Assurance Coordinator","Clinical Placement Coordinator"],
    "Facilities": ["Mechanical Systems Technician","Facilities Coordinator","Maintenance Supervisor"]
}

work_v3 = work_v2.copy()
apps_tmp = apps_df.copy()
apps_tmp["is_senior"] = apps_tmp.apply(is_senior_applicant, axis=1)

cutoff = datetime(2005,1,1).date()
appended = 0
used_texts = set()

for aid in apps_tmp.loc[apps_tmp["is_senior"], "ApplicantID"].unique():
    prior = work_v3[work_v3["ApplicantID"] == aid]
    if prior.empty:
        continue
    try:
        earliest = pd.to_datetime(prior["Start Date"], errors="coerce").dt.date.min()
    except Exception:
        earliest = None
    if earliest is None or earliest > cutoff:
        # Append an older role (2001–2004 start, 2005–2007 end)
        start_year = random.randint(2001, 2004)
        end_year = random.randint(2005, 2007)
        start_date = datetime(start_year, random.randint(1,12), random.randint(1,28)).date()
        end_date = datetime(end_year, random.randint(1,12), random.randint(1,28)).date()
        org_type = random.choice(["University","Employer"])
        if org_type == "University":
            country = random.choice(["USA","Canada"])
            org = random.choice(USA_UNIS if country=="USA" else CAN_UNIS)
            region = "National" if country=="USA" else "Canada"
        else:
            country = "USA"
            org = random.choice(NE_EMPLOYERS)
            region = "New England"
        cat = str(apps_tmp.loc[apps_tmp["ApplicantID"]==aid, "Job Category"].iloc[0])
        role = random.choice(ROLE_TITLE_MAP.get(cat, ["Senior Specialist"]))
        # Unique achievement
        for _ in range(10):
            ach = random.choice(ACH_NEW_POOL)
            if ach not in used_texts:
                used_texts.add(ach)
                break
        else:
            ach = f"Standardized reporting; audit findings reduced to {random.randint(10,40)}%"
        work_v3 = pd.concat([work_v3, pd.DataFrame([{
            "ApplicantID": aid, "Organization": org, "Organization Type": org_type,
            "Country": country, "Region": region, "Role": role,
            "Start Date": start_date, "End Date": end_date, "Achievements": ach
        }])], ignore_index=True)
        appended += 1

# Save v3
with pd.ExcelWriter(OUT_SENIOR_V3_XLSX, engine="openpyxl") as w:
    positions_df.to_excel(w, sheet_name="Positions", index=False)
    apps_df.to_excel(w, sheet_name="Applications", index=False)
    work_v3.to_excel(w, sheet_name="ApplicantWorkHistory", index=False)

# Check: seniors earliest start <= 2005
seniors = apps_tmp.loc[apps_tmp["is_senior"], "ApplicantID"].tolist()
earliest_by_senior = (work_v3[work_v3["ApplicantID"].isin(seniors)]
                      .assign(StartDateParsed=pd.to_datetime(work_v3["Start Date"], errors="coerce"))
                      .groupby("ApplicantID")["StartDateParsed"].min()
                      .dropna())
count_pre2005 = (earliest_by_senior <= pd.Timestamp("2005-01-01")).sum()

print(f"Appended older roles for seniors: {appended}")
print(f"Seniors with earliest start ≤ 2005: {count_pre2005} / {len(seniors)}")
OUT_SENIOR_V3_XLSX

#@title Quick reporting and sanity checks

# Local mix by category
local_mix = apps_df.groupby(["Job Category","Local"]).size().reset_index(name="Count")
display(local_mix.sort_values(["Job Category","Local"]))

# Applicants per category
cat_counts = apps_df.groupby("Job Category")["ApplicantID"].count().reset_index(name="Applicants")
display(cat_counts.sort_values("Applicants", ascending=False))

# Organization diversity in final work history
org_diversity = work_v3["Organization"].nunique()
print(f"Unique organizations in final ApplicantWorkHistory: {org_diversity}")

# Save summary CSVs (optional)
local_mix.to_csv(ART_DIR / "local_mix_by_category.csv", index=False)
cat_counts.to_csv(ART_DIR / "category_counts.csv", index=False)