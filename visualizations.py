
import csv
import os
import re
from collections import Counter, defaultdict
from statistics import mean

import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook


BASE = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.dirname(os.path.abspath(__file__))

FILE_2025 = os.path.join(BASE, "cleaned_indian_job_market_2025.xlsx")
FILE_2026 = os.path.join(BASE, "cleaned_indian_tech_jobs_2026.csv")

os.makedirs(OUTPUT, exist_ok=True)


def clean(value):
    if value is None:
        return ""
    return str(value).strip()


def number(value):
    try:
        value = str(value).replace(",", "").strip()
        match = re.search(r"-?\d+(\.\d+)?", value)
        return float(match.group()) if match else None
    except (ValueError, TypeError):
        return None


def normalize_city(value):
    city = clean(value).lower()
    aliases = {
        "bangalore": "Bengaluru",
        "bengaluru": "Bengaluru",
        "gurgaon": "Gurugram",
        "gurugram": "Gurugram",
        "bombay": "Mumbai",
        "mumbai": "Mumbai",
        "madras": "Chennai",
        "new delhi": "Delhi"
    }
    return aliases.get(city, city.title() if city else "Unknown")


def normalize_role(value):
    role = clean(value).lower()

    groups = {
        "Software Developer": ["software developer", "software engineer", "sde"],
        "Data Analyst": ["data analyst", "business analyst"],
        "Data Scientist": ["data scientist"],
        "AI / ML Engineer": ["machine learning", "ml engineer", "ai engineer"],
        "Web Developer": ["web developer", "frontend developer", "backend developer"],
        "DevOps Engineer": ["devops"],
        "Cloud Engineer": ["cloud engineer"],
        "Cybersecurity": ["cybersecurity", "security engineer"],
        "Other": []
    }

    for category, keywords in groups.items():
        if any(keyword in role for keyword in keywords):
            return category

    return "Other"


def read_2025():
    workbook = load_workbook(FILE_2025, read_only=True, data_only=True)
    sheet = workbook.active

    headers = [clean(cell).lower() for cell in next(sheet.iter_rows(values_only=True))]
    rows = []

    for values in sheet.iter_rows(min_row=2, values_only=True):
        row = dict(zip(headers, values))

        rows.append({
            "year": "2025",
            "title": clean(row.get("title")),
            "role": normalize_role(row.get("title")),
            "city": normalize_city(row.get("location")),
            "currency": clean(row.get("currency")).upper() or "Unknown",
            "min_salary": number(row.get("minimumsalary")),
            "max_salary": number(row.get("maximumsalary")),
            "min_exp": number(row.get("minimumexperience")),
            "max_exp": number(row.get("maximumexperience")),
            "skills": clean(row.get("tagsandskills"))
        })

    workbook.close()
    return rows


def read_2026():
    rows = []

    with open(FILE_2026, "r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)

        for row in reader:
            skills = clean(row.get("skills_required"))

            rows.append({
                "year": "2026",
                "title": clean(row.get("job_title")),
                "role": normalize_role(row.get("role_category") or row.get("job_title")),
                "city": normalize_city(row.get("primary_city")),
                "currency": "LPA",
                "min_salary": number(row.get("salary_min_lpa")),
                "max_salary": number(row.get("salary_max_lpa")),
                "mid_salary": number(row.get("salary_midpoint_lpa")),
                "min_exp": number(row.get("experience_min_yrs")),
                "max_exp": number(row.get("experience_max_yrs")),
                "skills": skills,
                "fresher": clean(row.get("is_fresher_friendly")).lower()
            })

    return rows


def average_salary(row):
    low = row.get("min_salary")
    high = row.get("max_salary")

    if low is not None and high is not None:
        return (low + high) / 2

    return low if low is not None else high


def save_chart(fig, filename):
    fig.update_layout(
        template="plotly_white",
        font=dict(family="Arial", size=13),
        title_font=dict(size=21),
        margin=dict(l=65, r=35, t=85, b=65),
        autosize=True,
        legend=dict(
            title_text="Select categories",
            orientation="h",
            yanchor="bottom",
            y=-0.28,
            xanchor="left",
            x=0
        )
    )

    fig.write_html(
        os.path.join(OUTPUT, filename),
        include_plotlyjs=True,
        full_html=True,
        config={
            "responsive": True,
            "displaylogo": False,
            "scrollZoom": True
        }
    )


data25 = read_2025()
data26 = read_2026()

all_data = data25 + data26


# 1. CITY-WISE JOB AVAILABILITY

city_counts = Counter()
for row in all_data:
    if row["city"] != "Unknown":
        city_counts[(row["year"], row["city"])] += 1

top_cities = [
    city for city, count in
    Counter({
        city: sum(v for (year, c), v in city_counts.items() if c == city)
        for year, city in city_counts
    }).most_common(12)
]

city_chart_data = [
    {"Year": year, "City": city, "Jobs": count}
    for (year, city), count in city_counts.items()
    if city in top_cities
]

fig = px.bar(
    city_chart_data,
    x="City",
    y="Jobs",
    color="Year",
    barmode="group",
    title="Job Availability Across Major Indian Cities",
    labels={"Jobs": "Number of job listings"},
    hover_data=["Year", "City", "Jobs"]
)
fig.update_xaxes(tickangle=-35)
save_chart(fig, "01_city_availability.html")


# 2. ROLE-WISE JOB DEMAND

role_counts = Counter()
for row in all_data:
    if row["role"] != "Other":
        role_counts[(row["year"], row["role"])] += 1

role_chart_data = [
    {"Year": year, "Role": role, "Jobs": count}
    for (year, role), count in role_counts.items()
]

fig = px.bar(
    role_chart_data,
    x="Jobs",
    y="Role",
    color="Year",
    barmode="group",
    orientation="h",
    title="Demand for Different Technology Roles",
    labels={"Jobs": "Number of job listings"},
    hover_data=["Year", "Role", "Jobs"]
)
fig.update_layout(yaxis={"categoryorder": "total ascending"})
save_chart(fig, "02_role_demand.html")


# 3. SKILL DEMAND TREEMAP

def extract_skills(text):
    if not text:
        return []

    parts = re.split(r"[,;|/]", text)
    return list({
        part.strip().title()
        for part in parts
        if len(part.strip()) > 1
    })


skill_counts = Counter()

for row in all_data:
    for skill in extract_skills(row["skills"]):
        skill_counts[skill] += 1

top_skills = skill_counts.most_common(25)

if top_skills:
    fig = px.treemap(
        names=[item[0] for item in top_skills],
        parents=[""] * len(top_skills),
        values=[item[1] for item in top_skills],
        title="Most Frequently Mentioned Skills",
        color=[item[1] for item in top_skills],
        color_continuous_scale="Blues"
    )
    fig.update_traces(
        textinfo="label+value+percent entry",
        hovertemplate="<b>%{label}</b><br>Mentions: %{value}<extra></extra>"
    )
    save_chart(fig, "03_skill_treemap.html")


# 4. ROLE DISTRIBUTION DONUT

role_distribution = Counter(
    row["role"] for row in data26 if row["role"] != "Other"
)

fig = px.pie(
    names=list(role_distribution.keys()),
    values=list(role_distribution.values()),
    hole=0.55,
    title="Technology Job Role Distribution – 2026"
)
fig.update_traces(
    textposition="inside",
    textinfo="percent",
    hovertemplate="<b>%{label}</b><br>Jobs: %{value}<br>Share: %{percent}<extra></extra>",
    showlegend=True
)
save_chart(fig, "04_role_distribution.html")


# 5. SALARY DISTRIBUTION PIE – 2026 DATA (LPA)

salary_values = [
    row.get("mid_salary") or average_salary(row)
    for row in data26
]

salary_values = [
    value for value in salary_values
    if value is not None and 0 < value < 200
]

salary_bands = {
    "Below 3 LPA": 0,
    "3–6 LPA": 0,
    "6–10 LPA": 0,
    "10–15 LPA": 0,
    "15–25 LPA": 0,
    "Above 25 LPA": 0
}

for salary in salary_values:
    if salary < 3:
        salary_bands["Below 3 LPA"] += 1
    elif salary < 6:
        salary_bands["3–6 LPA"] += 1
    elif salary < 10:
        salary_bands["6–10 LPA"] += 1
    elif salary < 15:
        salary_bands["10–15 LPA"] += 1
    elif salary < 25:
        salary_bands["15–25 LPA"] += 1
    else:
        salary_bands["Above 25 LPA"] += 1

fig = px.pie(
    names=list(salary_bands.keys()),
    values=list(salary_bands.values()),
    title="Salary Range Distribution – 2026",
    hole=0
)
fig.update_traces(
    textinfo="percent",
    hovertemplate="<b>%{label}</b><br>Listings: %{value}<br>Share: %{percent}<extra></extra>",
    showlegend=True
)
save_chart(fig, "05_salary_distribution.html")


# 6. FRESHER OPPORTUNITY FILTER FUNNEL – 2026

total_listings = len(data26)

fresher_listings = [
    row for row in data26
    if row["fresher"] in ("true", "1", "yes", "y")
]

fresher_with_salary = [
    row for row in fresher_listings
    if average_salary(row) is not None
]

funnel_labels = [
    "All job listings",
    "Fresher-friendly listings",
    "Fresher listings with salary data"
]

funnel_values = [
    total_listings,
    len(fresher_listings),
    len(fresher_with_salary)
]

fig = go.Figure(go.Funnel(
    y=funnel_labels,
    x=funnel_values,
    textinfo="value+percent initial",
    hovertemplate="<b>%{y}</b><br>Listings: %{x}<extra></extra>"
))

fig.update_layout(
    title="Fresher Opportunity Filtering – 2026",
    xaxis_title="Number of job listings"
)
save_chart(fig, "06_fresher_opportunities.html")


# 7. SALARY VS EXPERIENCE LINE GRAPH – 2026

experience_groups = defaultdict(list)

for row in data26:
    salary = row.get("mid_salary") or average_salary(row)
    minimum = row.get("min_exp")
    maximum = row.get("max_exp")

    if salary is None or minimum is None:
        continue

    if salary <= 0 or salary >= 200:
        continue

    if minimum < 0 or minimum > 30:
        continue

    experience_groups[int(minimum)].append(salary)

line_data = [
    {
        "Experience": experience,
        "Average Salary": round(mean(salaries), 2),
        "Job Listings": len(salaries)
    }
    for experience, salaries in sorted(experience_groups.items())
    if len(salaries) >= 10
]

fig = px.line(
    line_data,
    x="Experience",
    y="Average Salary",
    markers=True,
    title="Average Salary by Minimum Required Experience – 2026",
    hover_data=["Job Listings"],
    labels={
        "Experience": "Minimum experience required (years)",
        "Average Salary": "Average salary (LPA)"
    }
)

fig.update_traces(
    line=dict(width=3),
    marker=dict(size=9),
    hovertemplate=(
        "Experience: %{x} years<br>"
        "Average salary: %{y} LPA<br>"
        "Listings: %{customdata[0]}<extra></extra>"
    )
)
save_chart(fig, "07_salary_experience.html")


print("\nVISUALIZATION GENERATION COMPLETE")
print("Total graphs created: 7")
print("Output folder:", OUTPUT)

for filename in sorted(os.listdir(OUTPUT)):
    if filename.endswith(".html"):
        print("Created:", filename)
