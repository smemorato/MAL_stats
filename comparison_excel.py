import pandas as pd
import datetime
import numpy as np

def to_excel(workbook, lists):
    df = pd.DataFrame()
    pd.set_option("display.max_columns", None)

    for page in lists[0][1]:
        page_df = pd.json_normalize(page['data'])

        if "list_status.start_date" not in page_df.columns:
            page_df["list_status.start_date"] = ""
        if "list_status.finish_date" not in page_df.columns:
            page_df["list_status.finis_date"] = ""
    print(lists[0][0])
    page_df["username"] = lists[0][0]

    df = pd.concat([df, page_df])

    for users in lists[1:]:

        for page in users[1]:
            page_df = pd.json_normalize(page['data'])

            print(users[0])
            page_df["username"] = users[0]
            #print(page_df)
            df =pd.concat([df,page_df])
    print(df)
    df.columns = ["ID", "Title", "pic big", "pic medium", "start Date", "End Date", "mean",
                  "type", "genres", "num ep", "season year", "season", "source", "ep duration", "studios", "status"
        , "score", "ep watched", "rewatched", "update", "start date", "finish date","username"]




    #pd.set_option("display.max_rows", None, "display.max_columns", None)
    print(df)
    df_genreslist = df["genres"].explode().apply(pd.Series)["name"].unique()



    df_genrestable = df[["ID","genres"]].copy()
    df_genrestable = df_genrestable.explode("genres").reset_index(drop=True)
    print(df_genrestable)
    print(pd.json_normalize(df_genrestable["genres"]))
    df_genrestable= pd.concat([df_genrestable.drop(["genres"], axis=1), pd.json_normalize(df_genrestable["genres"])], axis=1)
    print(df_genrestable)
    df_genrestable.drop(columns = ["id"],inplace=True)
    df_genrestable.drop_duplicates(['ID','name'],inplace=True)


    #df["season year"].fillna(df['Title'], inplace=True)
    df['start Date']= pd.to_datetime(df['start Date'], errors="coerce")
    df['start date'] = pd.to_datetime(df['start date'], errors="coerce")
    df['finish date'] = pd.to_datetime(df['finish date'], errors="coerce")
    df["ep duration"] = df["ep duration"]/3600
    df["aux"] = pd.to_numeric(pd.DatetimeIndex(df['start Date']).month, errors="coerce")
    df["aux"] = df["aux"].replace([[1,2,3],[4,5,6],[7,8,9],[10,11,12]], ["winter","spring","summer","fall"])
    df['start Date'] = pd.to_numeric(pd.DatetimeIndex(df['start Date']).year, errors="coerce")
    df.drop(columns = ["pic big","pic medium","start Date","End Date"
                ,"genres","num ep","studios","status",
                "rewatched","update","aux"])

    df["season year"].fillna(df['start Date'], inplace=True)
    df["season"].fillna(df['aux'], inplace=True)

    df = df[["username","ID", "Title","type","source","mean","start date","finish date","score","season year", "season","ep watched",
             "ep duration"]]

    df["show duration"] = df["ep duration"]*df["ep watched"]

    df.replace({'score': 0}, np.nan, inplace= True)

    print(df.dtypes)
    #in case the dates are not valid
    df["days watching"] = (pd.to_datetime(df["finish date"], errors= "coerce")-
                           pd.to_datetime(df["start date"], errors="coerce"))/ np.timedelta64(1, 'D')+1
    df["hours a day"] = pd.to_numeric(df["show duration"], errors= "coerce")/df["days watching"]
    df["episodes a day"] = df["ep watched"] / df["days watching"]

    df_average= df.groupby("username")["score"].mean()
    print(df_average)
    print(df_average)
    df=pd.merge(df, df_average, how="inner", on="username")



    df_users_genres = pd.merge(df, df_genrestable, how="inner", on="ID")
    df_anime = pd.DataFrame()
    df_anime["id"] = pd.unique(df["ID"])
    print(df_anime)


    with pd.ExcelWriter(f"comparison/ test.xlsx", engine='openpyxl') as writer:


        # adds workbook and sheets to writer
        writer.book = workbook
        writer.sheets = dict((ws.title, ws) for ws in workbook.worksheets)

        df.to_excel(writer, sheet_name='users list', header=False, startrow=1,na_rep='NA')
        df_users_genres.to_excel(writer, sheet_name='users genres', header=False, startrow=1,na_rep='NA')
        df_anime.to_excel(writer,sheet_name="anime stats", header=False, startrow=4,na_rep='NA')

        writer.sheets["users list"].tables['tb_userslist'].ref = 'A1:S' + str(len(df) + 1)
        writer.sheets["users genres"].tables['table_genres'].ref = 'A1:T' + str(len(df_users_genres) + 1)
        #writer.sheets["anime stats"].tables['tb_anime'].ref = 'A4:I' + str(len(df_anime) + 1)
    #print((lists[0][1][0]["data"]))
    #df.columns =["username", "aniemid", "title", "episode", "date_start", "date_finish", "score", "year", "season"]


    #     "episode duration": 1,
    #     "show duration": 1,
    #     "hours a day": 1,
    #     "episodes a day": 1,
    #     "meanscore": 1,
    #
    #
    #
    # df2 = pd.json_normalize(lists[0][1][0]["data"])
    # pd.set_option("display.max_rows", None, "display.max_columns", None)
    # print("asdfg")
    # print(df2)
