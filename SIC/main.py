import function as f
import view as v
from tkinter import messagebox
import pandas as pd
import numpy as np
import datetime as dt

def norm_date_func(date_object):
    norm_date = date_object.strftime("%d %B %Y")
    return norm_date

def main():
    name, birth_date, insurance_start_date, sex = v.get_insured_data()

    birth_date = dt.datetime.strptime(birth_date, "%d-%m-%Y")
    insurance_start_date = dt.datetime.strptime(insurance_start_date, "%d-%m-%Y")

    if birth_date > insurance_start_date :
        messagebox.showwarning('ERROR', 'Birth date cannot be after insurance start date')
        v.get_insured_data()

    age,product,assumption,fractional_insurance,insurance_duration,insurance_deferred,fractional_annuity,annuity_duration,first_benefit,second_benefit = v.insurance_options(insurance_start_date,birth_date)

    columns = ['Attribute', 'Value']

    data = [
    ['Name', str(name).title()],
    ['Date of Birth', norm_date_func(birth_date)],
    ['Effective Date', norm_date_func(insurance_start_date)],
    ['Insurance Product', str(product).title()],
    ['Fractional Assumption', str(assumption).upper()],
    ['Annual Fraction of Insurance Cover Period', fractional_insurance],
    ['Insurance Cover Period', insurance_duration],
    ['Insurance Coverage Delay Period', insurance_deferred],
    ['Annual Fraction of Premium', fractional_annuity],
    ['Duration of Premium Payments in Years', annuity_duration],
    ['First Benefit', first_benefit],
    ['Second Benefit (Only for Endowment Product)', second_benefit]
    ]

    df1 = pd.DataFrame(data, columns=columns)

    annuity_interest_rate = v.create_entries(annuity_duration/fractional_annuity,'Annuity Interest Rate')
    annuity_array = np.arange(0,annuity_duration,fractional_annuity)
    annuity = f.annuity(age,sex,annuity_duration,fractional_annuity,assumption,annuity_interest_rate)

    objek = f.insurance(age,fractional_insurance,sex,assumption)
    df2 = pd.DataFrame()
    df2['Expected Benefit Payment Date'] = None

    if product == 'Wholelife insurance' :
        max_age = f.max_age()
        age_interval = max_age - age
        n_array = np.arange(0,age_interval,fractional_insurance)
        np.append(n_array,age_interval)
        defferred_interval = int(insurance_deferred/fractional_insurance)
        n_array = n_array[defferred_interval:]
        n_array = n_array + fractional_insurance
        insurance_interest_rate = v.create_entries(len(n_array),'Insurance Interest Rate')
        pv_benefit = objek.wholelife(insurance_deferred,insurance_interest_rate,first_benefit)
        premium = pv_benefit/annuity

        for index,value in enumerate(n_array) :
            df2.loc[index,'Expected Benefit Payment Date'] = norm_date_func(insurance_start_date + dt.timedelta(days=round(value * 360)))

    elif product == 'N-years term insurance':
        n_array = np.arange(0,insurance_duration,fractional_insurance)
        n_array = n_array[int(insurance_deferred/fractional_insurance):]
        np.append(n_array,n_array[-1]+fractional_insurance)
        n_array = n_array + fractional_insurance
        insurance_interest_rate = v.create_entries(len(n_array),'Insurance Interest Rate')
        pv_benefit = objek.ntermslife(insurance_duration,insurance_deferred,insurance_interest_rate,first_benefit)
        premium = pv_benefit/annuity

        for index,value in enumerate(n_array) :
            df2.loc[index,'Expected Benefit Payment Date'] = norm_date_func(insurance_start_date + dt.timedelta(days=round(value * 360)))

    elif product == 'Pure endowment insurance' :
        insurance_interest_rate = v.create_entries(1,'Insurance Interest Rate')
        pv_benefit = objek.pure_endowment(insurance_duration,insurance_interest_rate,first_benefit)
        premium = pv_benefit/annuity

        df2.loc[0,'Expected Benefit Payment Date'] = norm_date_func(insurance_start_date + dt.timedelta(days=round(insurance_duration*360)))
        
    elif product == 'Endowment insurance' :
        n_array = np.arange(0,insurance_duration,fractional_insurance)
        n_array = n_array + fractional_insurance
        insurance_interest_rate = v.create_entries(len(n_array)+1,'Insurance Interest Rate')
        pv_benefit = objek.endowment(insurance_interest_rate[:-1],insurance_interest_rate[-1],insurance_duration,first_benefit,second_benefit)
        premium = pv_benefit / annuity

        for index,value in enumerate(n_array) :
            df2.loc[index,'Expected Benefit Payment Date'] = norm_date_func(insurance_start_date + dt.timedelta(days=round(value * 360)))


    df1.loc[-1,'Attribute'] = 'Premium'
    df1.loc[-1,'Value'] = premium


    df2['Premium Payment Date'] = None
    for index,value in enumerate(annuity_array) :
        df2.loc[index,'Premium Payment Date'] = norm_date_func(insurance_start_date + dt.timedelta(days=round(value * 360)))

    folder_path = v.last_option()
    folder_path = f'{folder_path}/{name}.xlsx'

    df3 = pd.DataFrame()
    df3['Insurance'] = None
    df3['Annuity'] = None

    for index,value in enumerate(insurance_interest_rate):
        df3.loc[index,'Insurance'] = value
    
    for index,value in enumerate(annuity_interest_rate):
        df3.loc[index,'Annuity'] = value

    with pd.ExcelWriter(folder_path) as writer:
        df1.to_excel(writer, sheet_name='Summary', index=False)
        df2.to_excel(writer, sheet_name='Dates', index=False)
        df3.to_excel(writer,sheet_name='Interest Rates',index=False)