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
    page_df["username"] = lists[0][0]

    df = pd.concat([df, page_df])

    for users in lists[1:]:

        for page in users[1]:
            page_df = pd.json_normalize(page['data'])

            page_df["username"] = users[0]
            #print(page_df)
            df =pd.concat([df,page_df])

    df.columns = ["ID", "Title", "pic big", "pic medium", "start Date", "End Date", "mean",
                  "type", "genres", "num ep", "season year", "season", "source", "ep duration", "studios", "status"
        , "score", "ep watched", "rewatched", "update", "start date", "finish date","username"]




    #pd.set_option("display.max_rows", None, "display.max_columns", None)
    #print(df)
    df_genreslist = df["genres"].explode().apply(pd.Series)["name"].unique()



    df_genrestable = df[["ID","genres"]].copy()
    df_genrestable = df_genrestable.explode("genres").reset_index(drop=True)

    df_genrestable= pd.concat([df_genrestable.drop(["genres"], axis=1), pd.json_normalize(df_genrestable["genres"])], axis=1)

    #delete genre id
    df_genrestable.drop(columns = ["id"],inplace=True)
    df_genrestable.drop_duplicates(['ID','name'],inplace=True)


    #df["season year"].fillna(df['Title'], inplace=True)
    #todo get season and year from start date like in user stats
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

    #df_average= df.groupby("username")["score"].mean()
    df_mean = df.groupby("username").agg(user_mean=('score', 'mean'))
    print(df_mean)

    df=pd.merge(df, df_mean, how="inner", on="username")

    #https: // stackoverflow.com / questions / 11858472 / string - concatenation - of - two - pandas - columns
    df["range_mean"] = df.reset_index().index+2
    df["range_mean"] = df.agg(lambda x: f'''=IFERROR(vlookup(B{x['range_mean']},users_table[#All],4,false),"NA")''', axis=1)




    df_users_genres = pd.merge(df, df_genrestable, how="inner", on="ID")
    df_anime = pd.DataFrame()
    #print(df)
    df_anime_inf = df[['ID', 'Title', 'type', "season year", "season"]].drop_duplicates()
    df_anime_stats = df.groupby(['ID']).agg(anime_total=("Title","count"),
                                            anime_scored=("score","count"),
                                            anime_mean=("score","mean"),
                                            user_watched_mean=("user_mean","mean"))

    df_anime_stats["user_notscored_mean"]=(df_mean["user_mean"].mean()*df_mean["user_mean"].count()-(
        df_anime_stats["anime_scored"]*df_anime_stats["user_watched_mean"]))/(df_mean["user_mean"].count()-df_anime_stats["anime_scored"])
    df_anime_stats["anime_mean"].replace([np.nan], 0, inplace=True)
    df_anime_stats["user_notscored_mean"].replace([np.nan], 0, inplace=True)

    df_anime_stats["weighted_mean"]=(df_anime_stats["user_notscored_mean"]*(df_mean["user_mean"].count()-
                                        df_anime_stats["anime_scored"])+df_anime_stats["anime_mean"]*\
                                        df_anime_stats["anime_scored"])/df_mean["user_mean"].count()
    df_anime = pd.merge(df_anime_inf, df_anime_stats, how="inner", on="ID")
    df_anime["aux"] = df_anime.reset_index().index + 6
    df_anime["range_watched"]=df_anime.agg(lambda x: f'''=SUMPRODUCT((tb_userslist[series_animedb_id]=B{x['aux']})*
        (tb_userslist[my_finish_date]<=$E$2)*(tb_userslist[my_start_date]>=$E$1)*(tb_userslist[my_start_date]<>"NA"))'''
                                           , axis=1)

    df_anime["range_scored"]=df_anime.agg(lambda x: f'''=SUMPRODUCT((tb_userslist[series_animedb_id]=B{x['aux']})*
        (tb_userslist[my_score]<=10)*(tb_userslist[my_finish_date]<=$E$2)*(tb_userslist[my_start_date]>=$E$1)*
        (tb_userslist[my_start_date]<>"NA"))''', axis=1)
    df_anime["range_mean"]= df_anime.agg(lambda x: f'''=iferror(AVERAGEIFS(tb_userslist[my_score],tb_userslist[series_animedb_id],
        B{x['aux']},tb_userslist[my_start_date],">="&$E$1,tb_userslist[my_finish_date],"<="&$E$2),0)''', axis=1)
    df_anime["range_notscored_mean"]=df_anime.agg(lambda x: f'''=IF((COUNTIFS(users_table[username],"<>"&"",
        users_table[watched date range],">"&0)-M{x['aux']})>0,(AVERAGE(users_table[mean date range])*
        COUNTIFS(users_table[username],"<>"&"",users_table[watched date range],">"&0)-
        SUMPRODUCT((tb_userslist[series_animedb_id]=B{x['aux']})*(tb_userslist[my_score]<=10)*tb_userslist[user mean]*
        (tb_userslist[my_finish_date]<=$E$2)*(tb_userslist[my_start_date]>=$E$1)*(tb_userslist[my_start_date]<>"NA")))/
        (COUNTIFS(users_table[username],"<>"&"",users_table[watched date range],">"&0)-M{x['aux']}),0)''', axis=1)

    df_anime["range_weighted_mean"] = df_anime.agg(lambda x:f'''=(((COUNTIFS(users_table[username],"<>"&"",
        users_table[User mean score],">"&0)-M{x['aux']})*O{x['aux']})+(N{x['aux']}*L{x['aux']}))/
        COUNTIFS(users_table[username],"<>"&"",users_table[User mean score],">"&0)''', axis=1)
    df_anime.drop(columns=['user_watched_mean','aux'], inplace=True)
    '''=(((CONTAR.SE.S(users_table[username];"<>"&"";users_table[User mean score];">"&0)-[@[date watched scored]])*
    [@[date range not scored score]])+([@[date range mean score]]*[@[date range watched]]))/
    CONTAR.SE.S(users_table[username];"<>"&"";users_table[User mean score];">"&0)'''
    print(df_anime)
    #df_anime["id"] = pd.unique(df["ID"])





    n = len(df_anime.index) + 2
    while workbook['anime stats'].cell(row=n, column=2).value is not None:
        print(n)
        workbook['anime stats'].delete_rows(n, 1)
        n=n+1

    with pd.ExcelWriter(f"comparison/ test.xlsx", engine='openpyxl') as writer:


        # adds workbook and sheets to writer
        writer.book = workbook
        writer.sheets = dict((ws.title, ws) for ws in workbook.worksheets)

        df.to_excel(writer, sheet_name='users list', header=False, startrow=1,na_rep='NA')
        df_users_genres.to_excel(writer, sheet_name='users genres', header=False, startrow=1,na_rep='NA')
        df_anime.to_excel(writer,sheet_name="anime stats", header=False, startrow=5,na_rep='NA')

        writer.sheets["users list"].tables['tb_userslist'].ref = 'A1:T' + str(len(df) + 1)
        writer.sheets["users genres"].tables['table_genres'].ref = 'A1:U' + str(len(df_users_genres) + 1)
        writer.sheets["anime stats"].tables['tb_anime'].ref = 'A5:P' + str(len(df_anime) + 5)
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
