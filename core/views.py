from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from onboarding.models import Onboarding, OnboardingTasks
from users.models import Profile
from django.utils import timezone
from django.db.models import Count, Q
from django.db.models.functions import ExtractYear, ExtractMonth
from django.db.models.functions import TruncMonth, ExtractYear, ExtractMonth
from datetime import timedelta
import json
from collections import defaultdict


############################# Hiring Report ##############


from collections import defaultdict
from datetime import datetime


# Helper Functions
def get_gender_data(applicant_profiles):
    return list(applicant_profiles.values('gender').annotate(count=Count('gender')))

def get_age_data(applicant_profiles):
    age_ranges = [
        {"label": "20-30", "min_age": 20, "max_age": 30},
        {"label": "30-40", "min_age": 30, "max_age": 40},
        {"label": "40-50", "min_age": 40, "max_age": 50},
    ]
    age_data = []
    for age_range in age_ranges:
        count = sum(
            1
            for profile in applicant_profiles
            if profile.age and age_range["min_age"] <= profile.age < age_range["max_age"]
        )
        age_data.append({"label": age_range["label"], "count": count})
    return age_data

# def get_attrition_da

from datetime import datetime

def get_attrition_data():
    """
    Get attrition data grouped by year and month.
    """
    attrition_data = Profile.objects.filter(end_date__isnull=False).annotate(
        year=ExtractYear('end_date'),
        month=ExtractMonth('end_date')
    ).values('year', 'month').annotate(count=Count('id')).order_by('year', 'month')

    # Generate a dictionary for attrition counts by year and month
    attrition_dict = defaultdict(lambda: [0] * 12)  # Default is a list of 12 zeros for each year
    for item in attrition_data:
        attrition_dict[item['year']][item['month'] - 1] = item['count']  # Zero-indexed months

    # Ensure years from 2008 to current year are included
    current_year = datetime.now().year
    all_years = list(range(2008, current_year + 1))
    for year in all_years:
        if year not in attrition_dict:
            attrition_dict[year] = [0] * 12

    # Prepare the chart data format
    return {
        "years": all_years,  # All years from 2008 to now
        "data": attrition_dict,
    }




def get_hiring_data():
    """
    Get hiring data grouped by year and month.
    """
    hiring_data = Profile.objects.filter(date_of_joining__isnull=False).annotate(
        year=ExtractYear('date_of_joining'),
        month=ExtractMonth('date_of_joining')
    ).values('year', 'month').annotate(count=Count('id')).order_by('year', 'month')

    # Generate a dictionary for hiring counts by year and month
    hiring_dict = defaultdict(lambda: [0] * 12)  # Default is a list of 12 zeros for each year
    for item in hiring_data:
        hiring_dict[item['year']][item['month'] - 1] = item['count']  # Zero-indexed months

    # Ensure years from 2008 to current year are included
    current_year = datetime.now().year
    all_years = list(range(2008, current_year + 1))
    for year in all_years:
        if year not in hiring_dict:
            hiring_dict[year] = [0] * 12

    # Prepare the chart data format
    return {
        "years": all_years,  # All years from 2008 to now
        "data": hiring_dict,
    }


def get_task_data(today):
    active_onboardings = Onboarding.objects.filter(entry_date__gte=today)
    past_due_tasks = OnboardingTasks.objects.filter(
        Q(state='ST') | Q(state='PR'), date_due__lt=today
    )[:10]
    tasks_due_week = OnboardingTasks.objects.filter(
        date_due__lt=today + timedelta(days=7), date_due__gte=today
    )[:10]
    return active_onboardings, past_due_tasks, tasks_due_week





@login_required
def main_page(request):
    if not request.user.is_staff:
        return main_page_user(request)

    today = timezone.now()
    applicant_profiles = Profile.objects.filter(applicant__isnull=False)

    # Get chart data
    gender_data = get_gender_data(applicant_profiles)
    age_data = get_age_data(applicant_profiles)
    attrition_data = get_attrition_data()
    hiring_data = get_hiring_data()

    # Get task data
    active_onboardings, past_due_tasks, tasks_due_week = get_task_data(today)

    # Context for rendering
    context = {
    'gender_data': json.dumps(gender_data),
    'age_data': json.dumps(age_data),
    'attrition_data': json.dumps(attrition_data),
    'attrition_years': attrition_data['years'],  # For dropdown
    'hiring_data': json.dumps(hiring_data),  # Hiring data for chart
    'hiring_years': hiring_data['years'],  # For dropdown
    # 'hiring_years' : [2020, 2021, 2022],  # Replace with your logic
    #  attrition_years = [2019, 2020, 2021]  # Replace with your logic
    'active_onboardings': active_onboardings,
    'past_due_tasks': past_due_tasks,
    'tasks_due_week': tasks_due_week,
    }
 
    return render(request, 'core/main.html', context)

def main_page_user(request):
    today = timezone.now()
    tasks_due = OnboardingTasks.objects.filter(
        Q(date_due__gte=today) | Q(date_due=None),
        assigned_to=request.user
    ).exclude(state=OnboardingTasks.COMPLETED)
    tasks_completed = OnboardingTasks.objects.filter(
        state=OnboardingTasks.COMPLETED, assigned_to=request.user
    ).order_by('-date_due')[:10]

    context = {
        'tasks_due': tasks_due,
        'tasks_completed': tasks_completed,
    }
    return render(request, 'core/main_user.html', context)
