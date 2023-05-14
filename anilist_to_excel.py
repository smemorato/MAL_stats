from openpyxl import load_workbook
from openpyxl import Workbook
import datetime
import numpy as np
import json

import pandas as pd
import mal_to_excel

# todo make a dataframe with pd.json_normalize like in mal module
def to_excel(userlist, username):
    file = "user_stats_template.xlsx"
    workbook = load_workbook(filename=file)
    #print(json.dumps(userlist.json(), indent=0))
    userlist = userlist.json()
    df=pd.DataFrame()
    for lis in userlist["data"]["MediaListCollection"]["lists"]:

        #df = pd.concat(pd.DataFrame.from_dict(lis))

        df = pd.concat([df,pd.json_normalize(lis,"entries")])


    df["startedAt"]=pd.to_datetime(dict(year=df['startedAt.year'], month=df['startedAt.month'], day=df['startedAt.day']))
    df["completedAt"] = pd.to_datetime(dict(year=df['completedAt.year'], month=df['completedAt.month'],
                                            day=df['completedAt.day']))
    df["mediastart"] = pd.to_datetime(dict(year=df['media.startDate.year'], month=df['media.startDate.month'],
                                            day=df['media.startDate.day']))
    df["aux"] = pd.DatetimeIndex(df['mediastart']).month
    df["aux"] = df["aux"].replace([[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]], ["winter", "spring", "summer", "fall"])
    df["media.season"].fillna(df['aux'], inplace=True)
    df['aux'] = pd.DatetimeIndex(df['mediastart']).year
    df["media.seasonYear"].fillna(df['aux'], inplace=True)



    df_genreslist = df["media.genres"].explode().apply(pd.Series)
    df_genreslist.fillna("no Genre", inplace=True)
    df_tagslist = df["media.tags"].explode().apply(pd.Series)
    df_genreslist = pd.concat([df_genreslist, df_tagslist["name"]])
    df_genreslist = df_genreslist[0].unique()


    df_genrestable = df[["media.id", "media.genres"]].copy()
    df_genrestable = df_genrestable.explode("media.genres")
    df_genrestable["media.genres"].fillna("no Genre", inplace=True)

    df_tags = df[["media.id", "media.tags"]].copy()
    df_tags = df_tags.explode("media.tags")
    df_tags = pd.concat([df_tags.drop(["media.tags"], axis=1), df_tags["media.tags"].apply(pd.Series)], axis=1)
    df_tags.drop(columns=[0], inplace=True)

    df_genrestable = pd.concat([df_genrestable,df_tags.rename(columns={"name": "media.genres"})])
    df_genrestable.columns = ["ID", "genres"]



    df.drop(columns=['startedAt.year', 'startedAt.month', "startedAt.day","completedAt.year","completedAt.month",
                     "completedAt.day","status","media.tags","media.genres","media.startDate.year",
                     "media.startDate.month","media.startDate.day","mediastart","aux"], inplace=True)




    df.columns=["score","ID", "title","episodes", "format","seasonYear", "season", "source", "ep duration","meanScore",
                "startedAt", "completedAt"]



    df=df[["ID", "title", "format", "source", "meanScore", "startedAt", "completedAt", "score", "seasonYear",
           "season", "episodes", "ep duration"]]


    df["meanScore"] = df["meanScore"].astype(int)/10
    df["ep duration"] = df["ep duration"]/60


    oldest = pd.to_datetime(df['startedAt'], errors="coerce").min().year


    df["show duration"] = df["ep duration"]*df["episodes"]
    # in case the dates are not valid
    df["days watching"] = (pd.to_datetime(df["completedAt"], errors="coerce") -
                           pd.to_datetime(df["startedAt"], errors="coerce")) / np.timedelta64(1, 'D')+1
    df["hours a day"] = pd.to_numeric(df["show duration"], errors="coerce")/df["days watching"]
    df["episodes a day"] = df["episodes"] / df["days watching"]

    # resize tables on progression table
    workbook = mal_to_excel.resize_table_progression(workbook, oldest)
    workbook = mal_to_excel.resize_challenge_tables(workbook, df_genreslist)
    workbook = mal_to_excel.resize_days_table(oldest, workbook)
    # write list to excel
    mal_to_excel.insert_table(df, df_genrestable, workbook, username, "anilist")
