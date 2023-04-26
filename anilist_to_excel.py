from openpyxl import load_workbook
from openpyxl import Workbook
import datetime
import numpy as np
import json
from dateutil.parser import parse
from calendar import monthrange
import pandas as pd
import mal_to_excel


def to_excel(list, username):

    userlist = list.json()
    print(userlist)
    df = pd.DataFrame(columns=[ "animeid", "title", "media type","mean score", "date_start", "date_finish", "my score", "year",
                               "season", "episodes", "episode duration"])
    dfgenres = pd.DataFrame(columns=["animeid","genres"])
    oldest = datetime.datetime.now().year
    seen = []
    for list in userlist["data"]["MediaListCollection"]["lists"]:


            for entry in list["entries"]:
                print (entry)
                if entry["media"]["id"] not in seen:
                    seen.append(entry["media"]["id"])
                    if entry['startedAt']['year'] and entry['startedAt']['year']<oldest:
                        oldest=entry['startedAt']['year']


                    # get season by start date find needed
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


    print(df.dtypes)
    df["show duration"] = df["episode duration"]*df["episodes"]
    print(df.dtypes)
    #in case the dates are not valid
    df["days watching"] = (pd.to_datetime(df["date_finish"], errors= "coerce")-
                           pd.to_datetime(df["date_start"], errors="coerce"))/ np.timedelta64(1, 'D')+1
    df["hours a day"] = pd.to_numeric(df["show duration"], errors= "coerce")/df["days watching"]
    df["episodes a day"] = df["episodes"] / df["days watching"]

    dfanimegenres = pd.merge(df, dfgenres, how="inner", on="animeid")
    file = "python1.xlsx"
    workbook = load_workbook(filename=file)

    workbook = mal_to_excel.insert_dates(oldest, workbook)
    workbook = mal_to_excel.resize_table(workbook, oldest)

    genreslist = dfgenres["genres"].unique().tolist()
    workbook = mal_to_excel.resize_genres_tables(workbook, genreslist)

    with pd.ExcelWriter(f"userlist/{username}-anilist.xlsx", engine='openpyxl') as writer:


        # adds workbook and sheets to writer
        writer.book = workbook
        print(workbook['user_list'])
        writer.sheets = dict([('user_list', workbook['user_list']),("genres_table",workbook["genres_table"])])

        df.to_excel(writer, sheet_name='user_list', header=False, startrow=1, index = False)
        dfanimegenres.to_excel(writer, sheet_name='genres_table', header=False, startrow=1)


        writer.sheets["user_list"].tables['tb_list'].ref = 'A1:O' + str(len(df) + 1)
        writer.sheets["genres_table"].tables['tb_anime_genres'].ref = 'A1:Q' + str(len(dfgenres) + 1)