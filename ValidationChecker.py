import pandas as pd
import numpy as np
from datetime import datetime
import xlsxwriter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Color
from openpyxl import Workbook
import re
from dateutil.parser import parse
import time



def check_work_permit_exp_date_foreigner(df):
    date_regex = r'^\d{2}/\d{2}/\d{4}$'
    date_format = '%d/%m/%Y'

    date_valid = df['DATE'].str.contains(date_regex) & df['WORKPERMIT_EXPDATE'].str.contains(date_regex)

    permit_exp_dates = pd.to_datetime(df.loc[date_valid, 'WORKPERMIT_EXPDATE'], format=date_format)
    dates = pd.to_datetime(df.loc[date_valid, 'DATE'], format=date_format)

    result = pd.Series(['OK'] * len(df))
    result[(df['NATIONALITY_SENDER'] != 'MALDIVES') & df['WORKPERMIT_NUMBER'].isna()] = 'EMPTY'
    result[date_valid & (permit_exp_dates < dates)] = 'EXPIRED'
    result[date_valid & (permit_exp_dates >= dates)] = 'OK'

    return result.tolist()


def check_passport_expDate(df):
    result = pd.Series('OK', index=df.index)

    empty_mask = df['PASSPORT_EXPDATE'].isna()
    result.loc[empty_mask] = 'EMPTY'

    valid_mask = df['PASSPORT_EXPDATE'].notna() & df['DATE'].notna()
    passport_exp_date = pd.to_datetime(df.loc[valid_mask, 'PASSPORT_EXPDATE'], format='%d/%m/%Y', dayfirst=True,
                                       errors='coerce')
    date = pd.to_datetime(df.loc[valid_mask, 'DATE'], dayfirst=True, errors='coerce')

    result.loc[valid_mask & passport_exp_date.isna()] = 'INVALID FORMAT'
    result.loc[valid_mask & (passport_exp_date < date)] = 'EXPIRED'

    return result


def check_sending_receiving_country_same(df):
    list = np.where(df['SENDING_COUNTRY'] == df['RECEIVING_COUNTRY'], 'DUPLICATE', 'OK')
    return list


def countries_as_per_list_Of_countries(df, dfCountries):
    df['RECEIVING_COUNTRY'] = df['RECEIVING_COUNTRY'].astype(str).str.strip()
    df['SENDING_COUNTRY'] = df['SENDING_COUNTRY'].astype(str).str.strip()
    df['NATIONALITY_RECEIVER'] = df['NATIONALITY_RECEIVER'].astype(str).str.strip()
    df['NATIONALITY_SENDER'] = df['NATIONALITY_SENDER'].astype(str).str.strip()

    annex_list = set(dfCountries["ANNEX 1 : LIST OF COUNTRIES (REVISED ON 23 APRIL 2018)"])
    mask = (df['RECEIVING_COUNTRY'].isin(annex_list)) & \
           (df['SENDING_COUNTRY'].isin(annex_list)) & \
           (df['NATIONALITY_RECEIVER'].isin(annex_list)) & \
           (df['NATIONALITY_SENDER'].isin(annex_list))
    result = np.where(mask, 'OK', 'INVALID')
    return result


def ref_number_duplicate(df):
    is_duplicate = df.duplicated(subset='REF_NUMBER', keep=False)
    result = ['DUPLICATE' if d else 'OK' for d in is_duplicate]
    # result = np.where(is_duplicate,'DUPLICATE','OK')
    return result


def passport_validity(df):
    mask = ((df['PASSPORT_NUMBER'] == df['PASSPORT_EXPDATE']) | (df['PASSPORT_NUMBER'] == df['WORKPERMIT_EXPDATE']))
    result = np.where(mask, 'INVALID', 'OK')

    return result


def work_permit_uniqe_check(df, columnName):
    results = []
    df['NAME_SENDER_TEMP'] = df['NAME_SENDER'].str.upper().str.replace(' ', '')

    for index, row in df.iterrows():
        work_permit_or_passport_number = row[columnName]
        sender_nationality = row['NATIONALITY_SENDER']

        if columnName == 'WORKPERMIT_NUMBER' and sender_nationality == 'MALDIVES':
            results.append('OK')
            continue

        if pd.isna(work_permit_or_passport_number) or work_permit_or_passport_number == '':
            if columnName == 'WORKPERMIT_NUMBER' and sender_nationality == 'MALDIVES':
                results.append('OK')
            else:
                results.append('EMPTY')
        else:
            # Check if multiple sender_names have the same workpermit_number
            sender_names = df[df[columnName] == work_permit_or_passport_number]['NAME_SENDER_TEMP']
            if sender_names.nunique() > 1:
                results.append('INVALID')
            else:
                results.append('OK')

    df.drop(['NAME_SENDER_TEMP'], axis=1, inplace=True)

    return results


def extra_space_checker(df):
    df['NAME_SENDER'] = df['NAME_SENDER'].astype(str)
    df['NAME_RECEIVER'] = df['NAME_RECEIVER'].astype(str)

    regex = '\s{2,}' # to check if there are two or more consecutive whitespace characters
    extra_space_sender = df['NAME_SENDER'].str.contains(regex)
    extra_space_receiver = df['NAME_RECEIVER'].str.contains(regex)

    result = np.where(extra_space_sender | extra_space_receiver, 'EXTRA SPACE', 'OK')

    return result.tolist()




def seconds_converter(seconds):
    if seconds>=60:
        minutes = seconds // 60
        remaining_seconds = seconds % 60
        return f"{minutes:0f}m {remaining_seconds:0.2f}s"
    else:
        return f"{seconds:0.2f}s"



start_time = time.time()
df = pd.read_csv(r'C:\Users\raamee.ahmed\Desktop\ONE_FILE_UPLOAD\CombinedRawData.csv', )



dfCountries = pd.read_excel(
    r'M:\FileServer\PSD\POS\Payments Systems Oversight\Remittance Service\Supervision\WEEKLY REPORT FORMAT AND INSTRUCTIONS\List of countries\Annex 1 List of countries (23 April 2018).xlsx')

func1_start_time = time.time()
df['WORKPERMIT EXP VALIDITY'] = check_work_permit_exp_date_foreigner(df)
func1_end_time = time.time()
func1_time = func1_end_time - func1_start_time
fun1 = seconds_converter(func1_time)
print(f"Function 1 execution time: {fun1}")


func2_start_time = time.time()
df['PASSPORT EXP VALIDITY'] = check_passport_expDate(df)
func2_end_time = time.time()
func2_time = func2_end_time - func2_start_time
fun2 = seconds_converter(func2_time)
print(f"Function 2 execution time: {fun2}")


func3_start_time = time.time()
df['PASSPORT NO VALIDITY'] = passport_validity(df)
func3_end_time = time.time()
func3_time = func3_end_time - func3_start_time
fun3 = seconds_converter(func3_time)
print(f"Function 3 execution time: {fun3}")


func4_start_time = time.time()
df['RECEIVING_COUNTRY != SENDING_COUNTRY'] = check_sending_receiving_country_same(df)
func4_end_time = time.time()
func4_time = func4_end_time - func4_start_time
fun4 = seconds_converter(func4_time)
print(f"Function 4 execution time: {fun4}")


func5_start_time = time.time()
df['COUNTRY NAMES == list of countries'] = countries_as_per_list_Of_countries(df, dfCountries)
func5_end_time = time.time()
func5_time = func5_end_time - func5_start_time
fun5 = seconds_converter(func5_time)
print(f"Function 5 execution time: {fun5}")


func6_start_time = time.time()
df['DUPLICATE REF_NUMBER'] = ref_number_duplicate(df)
func6_end_time = time.time()
func6_time = func6_end_time - func6_start_time
fun6 = seconds_converter(func6_time)
print(f"Function 6 execution time: {fun6}")


func7_start_time = time.time()
df['DUPLICATE WORKPERMIT NUMBER'] = work_permit_uniqe_check(df, 'WORKPERMIT_NUMBER')
func7_end_time = time.time()
func7_time = func7_end_time - func7_start_time
fun7 = seconds_converter(func7_time)
print(f"Function 7 execution time: {fun7}")


func8_start_time = time.time()
df['DUPLICATE PASSPORT NUMBER'] = work_permit_uniqe_check(df, 'PASSPORT_NUMBER')
func8_end_time = time.time()
func8_time = func8_end_time - func8_start_time
fun8 = seconds_converter(func8_time)
print(f"Function 8 execution time: {fun8}")


func9_start_time = time.time()
df['SENDER NAME EXTRA SPACE'] = extra_space_checker(df)
func9_end_time = time.time()
func9_time = func9_end_time - func9_start_time
fun9 = seconds_converter(func9_time)
print(f"Function 9 execution time: {fun9}")

df.to_excel(r'C:\Users\raamee.ahmed\Desktop\ONE_FILE_UPLOAD\validationChecker.xlsx', index=False)


end_time = time.time()
total_time = end_time - start_time
totalTime = seconds_converter(total_time)
print(f"Total execution time: {totalTime}")
