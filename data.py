from datetime import datetime, date, timedelta


wilayas = {
    "blida":[],
    
    "bouira":[],

    "medea":[],

    "tiziouzou":[],

    "djelfa":[],

    "tipaza":[],

    "boumerdes":[],

    "aindefla":[],

    "chlef":[],
     
    "tissemsilt":[],
    
}



wilayas_ventes={
    "blida":[],
    
    "bouira":[],

    "medea":[],

    "tiziouzou":[],

    "djelfa":[],

    "tipaza":[],

    "boumerdes":[],

    "aindefla":[],

    "chlef":[],
     
    "tissemsilt":[],
    
}

wilayas_creance={
    "blida":[],
    
    "bouira":[],

    "medea":[],

    "tiziouzou":[],

    "djelfa":[],

    "tipaza":[],

    "boumerdes":[],

    "aindefla":[],

    "chlef":[],
     
    "tissemsilt":[],
    
}

wilayas_rnc={
    "blida":[],
    
    "bouira":[],

    "medea":[],

    "tiziouzou":[],

    "djelfa":[],

    "tipaza":[],

    "boumerdes":[],

    "aindefla":[],

    "chlef":[],
     
    "tissemsilt":[],
    
}











    
TB = 'tableau_de_bord\TB.xlsx'


# Get the current year and month
current_year = datetime.now().year
current_month = datetime.now().month

# Calculate the last day of the current month
# Handle February with leap years automatically
def last_day_of_month(year, month):
    if month == 12:  # December
        return date(year, 12, 31)
    else:
        return date(year, month + 1, 1) - timedelta(days=1)

# Calculate total days from the start of the year to the last day of the current month
total_days = (last_day_of_month(current_year, current_month) - date(current_year, 1, 1)).days + 1 