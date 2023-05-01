from openpyxl import load_workbook
from openpyxl import Workbook
import datetime
import numpy as np
import json

import pandas as pd
import mal_to_excel


def to_excel(list, username):
    file = "user_stats_template.xlsx"
    workbook = load_workbook(filename=file)


    userlist = list.json()
    df = pd.DataFrame(columns=["ID", "title", "media type","source","mean score", "date_start", "date_finish", "my score", "year",
                               "season", "episodes", "episode duration"])
    dfgenres = pd.DataFrame(columns=["ID","genres"])
    oldest = datetime.datetime.now().year
    seen = []
    for list in userlist["data"]["MediaListCollection"]["lists"]:


            for entry in list["entries"]:
                if entry["media"]["id"] not in seen:
                    seen.append(entry["media"]["id"])
                    if entry['startedAt']['year'] and entry['startedAt']['year']<oldest:
                        oldest=entry['startedAt']['year']


                    # get season by start date if needed
                    if entry["media"]["season"]:
                        season = entry["media"]["seasonYear"]
                    elif entry["media"]["startDate"]["month"] in [1,2,3]:
                        season = "WINTER"
                    elif entry["media"]["startDate"]["month"] in [4, 5, 6]:
                        season = "SPRING"
                    elif entry["media"]["startDate"]["month"] in [7, 8, 9]:
                        season = "SUMMER"
                    elif entry["media"]["startDate"]["month"] in [10, 11, 12]:
                        season = "SUMMER"
                    else:
                        season = "season no defined"
                    # get year by start date if needed
                    if entry["media"]["seasonYear"]:
                        seasonyear = entry["media"]["seasonYear"]
                    elif entry["media"]["startDate"]["year"]:
                        seasonyear = entry["media"]["startDate"]["year"]
                    else: "year not specified"

                    if entry["media"]["duration"]:
                        duration = entry["media"]["duration"]/60
                    else:
                        duration = "none"
                    if None not in (entry['startedAt']['year'], entry['startedAt']['month'], entry['startedAt']['day']):
                        startdate = datetime.datetime(entry['startedAt']['year'], entry['startedAt']['month'], entry['startedAt']['day'])
                    else:
                        startdate = "Date not valid"
                    if None not in (entry['completedAt']['year'], entry['completedAt']['month'], entry['completedAt']['day']):
                        completdate = datetime.datetime(entry['completedAt']['year'], entry['completedAt']['month'], entry['completedAt']['day'])
                    else:
                        completdate = "Date no valid"
                    try:
                        if completdate< startdate:
                            completdate = "complet  date is before start date"
                    except Exception:
                        pass




                    df.loc[len(df.index)] = [
                        entry["media"]["id"],
                        entry["media"]["title"]["romaji"],
                        entry["media"]["format"],
                        entry["media"]["source"],
                        entry["media"]["meanScore"]/10,
                        startdate,
                        completdate,
                        entry["score"],
                        seasonyear,
                        season,
                        entry["media"]["episodes"],
                        duration
                    ]

                    for genre in entry["media"]["genres"]:
                        dfgenres.loc[len(dfgenres.index)] = [entry["media"]["id"],genre]
                    for tag in entry["media"]["tags"]:
                        dfgenres.loc[len(dfgenres.index)] = [entry["media"]["id"],tag["name"]]



    df["show duration"] = df["episode duration"]*df["episodes"]
    #in case the dates are not valid
    df["days watching"] = (pd.to_datetime(df["date_finish"], errors= "coerce")-
                           pd.to_datetime(df["date_start"], errors="coerce"))/ np.timedelta64(1, 'D')+1
    df["hours a day"] = pd.to_numeric(df["show duration"], errors= "coerce")/df["days watching"]
    df["episodes a day"] = df["episodes"] / df["days watching"]





    #resize tables on progresssion table
    workbook = mal_to_excel.resize_table(workbook, oldest)
    genreslist = dfgenres["genres"].unique().tolist()
    workbook = mal_to_excel.resize_genres_tables(workbook, genreslist)
    workbook=mal_to_excel.insert_dates(oldest,workbook)
    #write list to excel
    mal_to_excel.insert_table(df, dfgenres, workbook, username,"anilist")
