from openpyxl import load_workbook
from openpyxl import Workbook
import datetime
import json
from dateutil.parser import parse
from calendar import monthrange
import pandas as pd
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, DataBarRule, IconSetRule
import numpy as np


def insert_table(df, df_genrestable, workbook, username):
    # df = pd.json_normalize(json.dumps(anilist))

    sheet = workbook['user_list']
    table = sheet.tables['tb_list']
    tablerange = table.ref

    merged = pd.merge(df, df_genrestable, how="inner", on="ID")
    with pd.ExcelWriter(f"userlist/python1 - {username}.xlsx", engine='openpyxl') as writer:


        # adds workbook and sheets to writer
        writer.book = workbook
        writer.sheets = dict((ws.title, ws) for ws in workbook.worksheets)

        df.to_excel(writer, sheet_name='user_list', header=False, startrow=1)
        merged.to_excel(writer, sheet_name='genres_table', header=False, startrow=1)

        writer.sheets["user_list"].tables['tb_list'].ref = 'A1:Q' + str(len(df) + 1)
        writer.sheets["genres_table"].tables['tb_anime_genres'].ref = 'A1:R' + str(len(merged) + 1)
    #print(merged)

    print("please")


def to_excel(userlist, username):
    file = "python1.xlsx"
    workbook = load_workbook(filename=file)
    sheet = workbook['user_list']
    m = 1
    df= pd.DataFrame()
    for page in userlist:
        page_df = pd.json_normalize(page['data'])

        page_df.columns = ["ID", "Title","pic big","pic medium","start Date","End Date","mean",
                    "type","genres","num ep","season year", "season","source","ep duration","studios","status"
                    ,"score","ep watched","rewatched","update","start date","finish date"]
        df =pd.concat([df,page_df])

    #pd.set_option("display.max_rows", None, "display.max_columns", None)
    pd.set_option("display.max_columns", None)
    df_genreslist = df["genres"].explode().apply(pd.Series)["name"].unique()



    df_genrestable = df[["ID","genres"]].copy()
    df_genrestable=df_genrestable.explode("genres")
    df_genrestable= pd.concat([df_genrestable.drop(["genres"], axis=1), df_genrestable["genres"].apply(pd.Series)], axis=1)
    df_genrestable.drop(columns = ["id",0],inplace=True)



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
    df = df[["ID", "Title","type","source","mean","start date","finish date","score","season year", "season","ep watched",
             "ep duration"]]
    df["show duration"] = df["ep duration"]*df["ep watched"]
    print(df.dtypes)
    #in case the dates are not valid
    df["days watching"] = (pd.to_datetime(df["finish date"], errors= "coerce")-
                           pd.to_datetime(df["start date"], errors="coerce"))/ np.timedelta64(1, 'D')+1
    df["hours a day"] = pd.to_numeric(df["show duration"], errors= "coerce")/df["days watching"]
    df["episodes a day"] = df["ep watched"] / df["days watching"]



    n = len(df.index)
    sheet.tables['tb_list'].ref = 'A1:O' + str(n)
    print(sheet.tables['tb_list'].ref)
    n = n + 1
    # eleminar valores de profile anterior tiver mais entries
    while sheet.cell(row=n, column=2).value is not None:
        sheet.delete_rows(n, 1)

    # testar numero de entries
    comp = n - 1
    print(f'{comp} shows completos')

    oldest=pd.to_datetime(df['start date'], errors="coerce").min().year

    #add all dates starting from the 1st of january of the oldest year entry to ""Dias"" woorksheet
    workbook = insert_dates(oldest, workbook)

    # resize the table to fit the new entries
    workbook = resize_table(workbook, oldest)

    workbook = resize_genres_tables(workbook, df_genreslist)

    # add genres to table
    insert_table(df, df_genrestable, workbook, username)
    # workbook.save(filename=f"/userlist/python1 - {username}.xlsx")


def insert_dates(oldest: int, workbook):
    sheet = workbook['days']

    date = datetime.datetime(year=oldest, month=1, day=1)

    hoje = datetime.datetime.today()

    # (datetime.timedelta(days=1))
    i = 2
    sheet.cell(row=i, column=1).value = date
    while date <= hoje:
        i = i + 1
        date = date + datetime.timedelta(days=1)

        sheet.cell(row=i, column=1).value = date

        # sheet.cell(row=i, column=2).value = '=SOMA.SE.S(user_list!N:N;user_list!E:E;"<="&A{};user_list!F:F;">="&A{}'.format(i, i)
        cell = "B{}".format(i)
        sheet[cell] = '=SUMIFS(user_list!P:P,user_list!G:G,"<="&A{},user_list!H:H,">="&A{})'.format(i, i)
        cell = "C{}".format(i)
        sheet[cell] = '=SUMIFS(user_list!Q:Q,user_list!G:G,"<="&A{},user_list!H:H,">="&A{})'.format(i, i)
        cell = "D{}".format(i)
        sheet[
            cell] = '=SUMIFS(user_list!P:P,user_list!G:G,"<="&A{},user_list!H:H,">="&A{},user_list!J:J,days!G{})+SUMIFS(user_list!P:P,user_list!G:G,"<="&A{},user_list!H:H,">="&A{},user_list!J:J,days!G{}-1)'.format(
            i, i, i, i, i, i)
        cell = "E{}".format(i)
        sheet[cell] = '=COUNTIF(user_list!H:H,A{})'.format(i)
        cell = "F{}".format(i)
        sheet[cell] = '=MONTH(A{})'.format(i)
        cell = "G{}".format(i)
        sheet[cell] = '=YEAR(A{})'.format(i)

        # print(type(sheet.cell(row=i, column=1).value))
    sheet.tables['tb_days'].ref = 'A1:G' + str(i)
    i = i + 1

    while sheet.cell(row=i, column=2).value is not None:
        sheet.delete_rows(i, 1)
        i = i + 1
    return workbook


# updates table size to fit the new entries
def resize_table(workbook, oldest: int):
    sheet = workbook['progression']
    table = sheet.tables['tb_progression_month']

    # RESIZE THE MONTH TABLE
    year = oldest
    today_year = datetime.datetime.now().year
    row = 3
    cell = "B3"
    sheet[cell] = oldest
    cell = "C3"
    datetime_month = datetime.datetime.strptime("1", "%m")
    sheet[cell] = datetime_month.strftime("%b")
    row = row + 1

    for i in range(2, 13):
        cellyear = "B{}".format(row)
        cellmonth = "C{}".format(row)
        cellmonthnumber = "D{}".format(row)
        cellep = "E{}".format(row)
        cellhours = "F{}".format(row)
        cellfinish = "G{}".format(row)
        cellpercentage = "H{}".format(row)
        cellhoursday = "I{}".format(row)

        sheet[cellyear] = '=B{}'.format(row - 1)
        datetime_moth = datetime.datetime.strptime(str(i), "%m")
        sheet[cellmonth] = datetime_moth.strftime("%b")
        sheet[cellmonthnumber] = i
        sheet[cellep] = "=SUMIFS(days!C:C,days!G:G,B{},days!F:F,D{})".format(row, row)
        sheet[cellhours] = "=SUMIFS(days!B:B,days!G:G,B{},days!F:F,D{})".format(row, row)
        sheet[cellfinish] = "=SUMIFS(days!E:E,days!F:F,progression!D{},days!G:G,progression!B{})".format(
            row, row)
        sheet[cellpercentage] = "=IF(F{}>0,SUMIFS(days!D:D,days!G:G,B{},days!F:F,D{})/F{},0)".format(row, row, row, row)
        sheet[cellhoursday] = "=F{}/DAY(EOMONTH(DATE(B{},D{},1),0))".format(row, row, row)
        row = row + 1
    year = year +1
    if oldest < today_year:
        while year <= today_year:
            for i in range(1, 13):
                cellyear = "B{}".format(row)
                cellmonth = "C{}".format(row)
                cellmonthnumber = "D{}".format(row)
                cellep = "E{}".format(row)
                cellhours = "F{}".format(row)
                cellfinish = "G{}".format(row)
                cellpercentage = "H{}".format(row)
                cellhoursday = "I{}".format(row)

                sheet[cellyear] = '=B{}+1'.format(row - 12)
                datetime_moth = datetime.datetime.strptime(str(i), "%m")
                sheet[cellmonth] = datetime_moth.strftime("%b")
                sheet[cellmonthnumber] = i
                sheet[cellep] = "=SUMIFS(days!C:C,days!G:G,B{},days!F:F,D{})".format(row, row)
                sheet[cellhours] = "=SUMIFS(days!B:B,days!G:G,B{},days!F:F,D{})".format(row, row)
                sheet[cellfinish] = "=SUMIFS(days!E:E,days!F:F,progression!D{},days!G:G,progression!B{})".format(
                    row, row)
                sheet[cellpercentage] = "=IF(F{}>0,SUMIFS(days!D:D,days!G:G,B{},days!F:F,D{})/F{},0)".format(row, row,
                                                                                                             row, row)
                sheet[cellhoursday] = "=F{}/DAY(EOMONTH(DATE(B{},D{},1),0))".format(row, row, row)
                row = row + 1
            year = year + 1
        table.ref = 'B2:J' + str(row - 1)

    # RESIZE THE SEASON TABLE
    seasonyear = oldest
    seasonrow = 3
    cellseasonyear = "L{}".format(seasonrow)
    sheet[cellseasonyear] = "=B3"
    cellseason = "M{}".format(seasonrow)
    sheet[cellseason] = "WINTER"
    seasonrow = seasonrow+1

    for i in range(2, 5):
        cellseasonyear = "L{}".format(seasonrow)
        sheet[cellseasonyear] = "=L{}".format(seasonrow-1)
        cellseason = "M{}".format(seasonrow)
        if i == 2:
            sheet[cellseason] = "SPRING"
        elif i == 3:
            sheet[cellseason] = "SUMMER"
        elif i == 4:
            sheet[cellseason] = "FALL"

        cellseasonep = "N{}".format(seasonrow)
        sheet[cellseasonep] = '''=SUMIFS(tb_progression_month[ep],tb_progression_month[month2],
        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
        cellseasonhours = "O{}".format(seasonrow)
        sheet[cellseasonhours] = '''=SUMIFS(tb_progression_month[hours spent],tb_progression_month[month2],
        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
        cellseasonfinish = "P{}".format(seasonrow)
        sheet[cellseasonfinish] = '''=SUMIFS(tb_progression_month[finished shows],tb_progression_month[month2],
        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
        cellseasonpercentage = "Q{}".format(seasonrow)
        sheet[cellseasonpercentage] = '''=IF(O{}>0,SUMIFS(days!D:D,days!F:F,SE(M{}="WINTER","<=3",
        IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),days!F:F,IF(M{}="WINTER",">=1",
        IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),days!G:G,L{})/O{},
        0)'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
        cellseasonhoursday = "R{}".format(seasonrow)
        # I got lazy and considered each month month with 30 days
        sheet[cellseasonhoursday] = '''=SUMIFS(tb_progression_month[hours spent],tb_progression_month[month2],
        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],L{})
        /90'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
        seasonrow = seasonrow + 1
    seasonyear = seasonyear + 1
    if oldest <= today_year:
        while seasonyear <= today_year:
            for i in range(1,5):
                cellseasonyear = "L{}".format(seasonrow)
                sheet[cellseasonyear] = "=L{}+1".format(seasonrow - 4)
                cellseason = "M{}".format(seasonrow)
                if i == 1:
                    sheet[cellseason] = "WINTER"
                elif i ==2:
                    sheet[cellseason] = "SPRING"
                elif i == 3:
                    sheet[cellseason] = "SUMMER"
                elif i == 4:
                    sheet[cellseason] = "FALL"

                cellseasonep = "N{}".format(seasonrow)
                sheet[cellseasonep] = '''=SUMIFS(tb_progression_month[ep],tb_progression_month[month2],
                        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
                        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
                        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
                cellseasonhours = "O{}".format(seasonrow)
                sheet[cellseasonhours] = '''=SUMIFS(tb_progression_month[hours spent],tb_progression_month[month2],
                        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
                        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
                        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
                cellseasonfinish = "P{}".format(seasonrow)
                sheet[cellseasonfinish] = '''=SUMIFS(tb_progression_month[finished shows],tb_progression_month[month2],
                        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
                        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
                        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
                cellseasonpercentage = "Q{}".format(seasonrow)
                sheet[cellseasonpercentage] = '''=IF(O{}>0,SUMIFS(days!D:D,days!F:F,SE(M{}="WINTER","<=3",
                        IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),days!F:F,IF(M{}="WINTER",">=1",
                        IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),days!G:G,L{})/O{},
                        0)'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow,
                                     seasonrow, seasonrow)
                cellseasonhoursday = "R{}".format(seasonrow)
                # I got lazy and considered each month month with 30 days
                sheet[cellseasonhoursday] = '''=SUMIFS(tb_progression_month[hours spent],tb_progression_month[month2],
                        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
                        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],L{})
                        /90'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)

                seasonrow = seasonrow +1
            seasonyear = seasonyear + 1

    sheet.tables['tb_progression_season'].ref = 'L2:R' + str(seasonrow - 1)


    # i had to add the conditional formatting in the code because even though the table size is increase the conditional formatting
    # remains the same

    sheet.conditional_formatting.add("E3:E{}".format(row - 1),
                                     ColorScaleRule(start_type='min', start_color='f8696b',
                                                    mid_type='percentile', mid_value=50, mid_color='ffeb84',
                                                    end_type='max', end_color='63be7b'))
    sheet.conditional_formatting.add("F3:F{}".format(row - 1),
                                     ColorScaleRule(start_type='min', start_color='f8696b',
                                                    mid_type='percentile', mid_value=50, mid_color='ffeb84',
                                                    end_type='max', end_color='63be7b'))
    sheet.conditional_formatting.add("G3:G{}".format(row - 1),
                                     ColorScaleRule(start_type='min', start_color='f8696b',
                                                    mid_type='percentile', mid_value=50, mid_color='ffeb84',
                                                    end_type='max', end_color='5a8bc6'))

    sheet.conditional_formatting.add("H3:H{}".format(row - 1),
                                     DataBarRule(start_type='num', start_value=0, end_type='num', end_value='1',
                                                 color="FF638EC6", showValue="None", minLength=None, maxLength=None))
    sheet.conditional_formatting.add("I3:I{}".format(row - 1),
                                     IconSetRule('3Symbols', 'percent', [0, 33, 67], showValue=None, percent=None))

    return workbook


def resize_genres_tables(workbook, genreslist):
    sheet = workbook['challenge']
    row = 3
    for genre in genreslist:
        print(genre)
        cellgenres = "J{}".format(row)
        sheet[cellgenres] = genre
        celltotal = "K{}".format(row)
        sheet[celltotal] = '''=COUNTIFS(tb_anime_genres[genre],J{}, tb_anime_genres[my_start_date],
        ">=" & $B$1,tb_anime_genres[my_finish_date],"<="&$B$2)'''.format(row)
        cellpercentage = "L{}".format(row)
        sheet[cellpercentage] = '=K{}/$C$8'.format(row)
        cellaverage = "M{}".format(row)
        sheet[cellaverage] = '''=IF(K{}=0,0,AVERAGEIFS(tb_anime_genres[my_score],tb_anime_genres[genre],J{}, 
        tb_anime_genres[my_start_date],">=" & $B$1,tb_anime_genres[my_finish_date],"<="&$B$2))'''.format(row, row)
        cellhours = "N{}".format(row)
        sheet[cellhours] = '''=SUMIFS(tb_anime_genres[show duration],tb_anime_genres[genre],J{}, tb_anime_genres[my_start_date],
        ">=" & $B$1,tb_anime_genres[my_finish_date],"<="&$B$2)'''.format(row)
        cellweightedmean = "O{}".format(row)
        sheet[cellweightedmean] = '''=IF(K{}=0,"",M{}*(K{}/$C$8)+$C$17*($C$8-K{})/$C$8)'''.format(row, row, row, row)
        row = row+1

    sheet.tables['tb_challenge'].ref = 'I2:P' + str(row - 1)

    sheet = workbook['genres_stats']
    row = 2
    for genre in genreslist:
        cellgenres = "A{}".format(row)
        sheet[cellgenres] = genre
        celltotal = "B{}".format(row)
        sheet[celltotal] = '''=COUNTIF(tb_anime_genres[genre],A{})'''.format(row)
        cellpercentage = "C{}".format(row)
        sheet[cellpercentage] = '=B{}/COUNTIF(tb_list[series_title],"<>"&"")'.format(row)
        cellaverage = "D{}".format(row)
        sheet[cellaverage] = '''=AVERAGEIFS(tb_anime_genres[my_score],tb_anime_genres[genre],A{})'''.format(row)
        cellhours = "E{}".format(row)
        sheet[cellhours] = '''=SUMIFS(tb_anime_genres[show duration],tb_anime_genres[genre],A{})'''.format(row)
        cellweightedmean = "F{}".format(row)
        sheet[cellweightedmean] = '''=D{}*(B{}/COUNTIF(tb_list[series_title], "<>"&""))+
        AVERAGEIF(tb_list[my_score],"<>"&0,tb_list[my_score])*(COUNTIF(tb_list[series_title],"<>"&"")-
        B{})/COUNTIF(tb_list[series_title],"<>"&"")'''.format(row, row, row)
        row = row + 1

    sheet.tables['tb_genres'].ref = 'A1:F' + str(row - 1)

    return workbook

# todo: add the genres to the challenge worksheet
# I don't know who to get a list of all genres so I made a list of all genres in the user_list and then add to the tables
# and since a new genre may be add i have to do it every time a update a userstats
def add_genre_to_tabel(genrelist):
    pass
