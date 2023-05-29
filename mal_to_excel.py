from openpyxl import load_workbook
from openpyxl import Workbook
import datetime
import json
import pandas as pd
from openpyxl.styles import Color, PatternFill, Font, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, DataBarRule, IconSetRule
import numpy as np


def insert_table(df, df_genrestable,df_days, workbook, username,source):
    # df = pd.json_normalize(json.dumps(anilist))


    # Risize the list table
    n = len(df.index)+1
    while workbook['user_list'].cell(row=n, column=2).value is not None:
        workbook['user_list'].delete_rows(n, 1)
        n=n+1

    merged = pd.merge(df, df_genrestable, how="inner", on="ID")
    # Resize the genres table
    n = len(merged.index)+1

    while workbook['genres_table'].cell(row=n, column=2).value is not None:
        workbook['genres_tablet'].delete_rows(n, 1)
        n=n+1

    with pd.ExcelWriter(f"userlist/{username}_{source}.xlsx", engine='openpyxl') as writer:

        # adds workbook and sheets to writer
        writer.book= workbook
        writer.sheets = dict((ws.title, ws) for ws in workbook.worksheets)

        df.to_excel(writer, sheet_name='user_list', header=False, startrow=1, index=False)
        merged.to_excel(writer, sheet_name='genres_table', header=False, startrow=1, index=False)
        df_days.to_excel(writer, sheet_name="days", header=False, startrow=1, index=False)

        writer.sheets["user_list"].tables['tb_list'].ref = 'A1:P' + str(len(df) + 1)
        writer.sheets["genres_table"].tables['tb_anime_genres'].ref = 'A1:Q' + str(len(merged) + 1)

        #writer.sheets["progression"].column_dimensions["N"].number_format ="0.00"
        #round numbers
        for r in range(2,len(df_days)+3):
            writer.sheets["days"][f'B{r}'].number_format = "0.00"
            writer.sheets["days"][f'C{r}'].number_format = "0.00"
            writer.sheets["days"][f'D{r}'].number_format = "0.00"




def to_excel(userlist, username):
    file = "user_stats_template.xlsx"
    workbook = load_workbook(filename=file)

    m = 1
    df = pd.DataFrame()
    for page in userlist:
        page_df = pd.json_normalize(page['data'])

        page_df.columns = ["ID", "Title", "pic big", "pic medium", "start Date", "End Date", "mean", "type", "genres",
                           "num ep", "season year", "season", "source", "ep duration", "studios", "status", "score",
                           "ep watched", "rewatched", "update", "start date", "finish date"]
        df = pd.concat([df, page_df])


    pd.set_option("display.max_columns", None)

    df_genreslist = df["genres"].explode().apply(pd.Series)
    df_genreslist["name"].fillna("no Genre", inplace=True)
    df_genreslist = df_genreslist["name"].unique()

    # df_genreslist.fillna("no Genre", inplace=True)

    df_genrestable = df[["ID", "genres"]].copy()
    df_genrestable = df_genrestable.explode("genres")
    df_genrestable = pd.concat([df_genrestable.drop(["genres"], axis=1), df_genrestable["genres"].apply(pd.Series)], axis=1)
    df_genrestable.drop(columns=["id", 0], inplace=True)
    df_genrestable["name"].fillna("no Genre", inplace=True)

    #filter wierd mal dates like 30/2/2022
    df['start Date'] = pd.to_datetime(df['start Date'], errors="coerce")
    df['start date'] = pd.to_datetime(df['start date'], errors="coerce")
    df['finish date'] = pd.to_datetime(df['finish date'], errors="coerce")
    df["ep duration"] = df["ep duration"]/3600

    df["aux"] = pd.to_numeric(pd.DatetimeIndex(df['start Date']).month, errors="coerce")
    df["aux"] = df["aux"].replace([[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]], ["winter", "spring", "summer", "fall"])
    df["season"].fillna(df['aux'], inplace=True)
    df['aux'] = pd.to_numeric(pd.DatetimeIndex(df['start Date']).year, errors="coerce")
    df["season year"].fillna(df['aux'], inplace=True)

    df.drop(columns=["pic big", "pic medium", "start Date", "End Date"
        , "genres", "num ep", "studios", "status",
                     "rewatched", "update", "aux"])
    df = df[["ID", "Title", "type", "source", "mean", "start date", "finish date", "score", "season year", "season",
             "ep watched", "ep duration"]]
    df["show duration"] = df["ep duration"]*df["ep watched"]

    # in case the dates are not valid
    df["days watching"] = (pd.to_datetime(df["finish date"], errors= "coerce") -
                           pd.to_datetime(df["start date"], errors="coerce")) / np.timedelta64(1, 'D')+1
    df["hours a day"] = pd.to_numeric(df["show duration"], errors= "coerce")/df["days watching"]
    df["episodes a day"] = df["ep watched"] / df["days watching"]


    oldest = pd.to_datetime(df['start date'], errors="coerce").min().year

    # get dataframe with date from the oldest date to today
    df_days = resize_days_table(oldest, workbook)

    # resize the table to fit the new entries
    workbook = resize_table_progression(workbook, oldest)

    workbook = resize_challenge_tables(workbook, df_genreslist)

    # add genres to table
    insert_table(df, df_genrestable, df_days, workbook, username, "mal")


def resize_days_table(oldest: int, workbook):


    oldest = datetime.datetime(year=oldest, month=1, day=1)
    today = datetime.datetime.today()
    df_days = pd.DataFrame()
    df_days["dates"] = pd.date_range(start=oldest,end=today)
    i = 2


    df_days["aux"] = df_days.reset_index().index + 1

    df_days["hours"] = df_days.agg(lambda x: f'''=SUMIFS(tb_list[Hours a Day],tb_list[Start Date],"<="&A{x['aux']},
                                    tb_list[Finish Date],">="&A{x['aux']})''', axis=1)
    df_days["ep"] = df_days.agg(lambda x: f'''=SUMIFS(tb_list[EP a Day],tb_list[Start Date],"<="&A{x['aux']},
                                    tb_list[Finish Date],">="&A{x['aux']})''', axis=1)

    df_days["hours season"] = df_days.agg(lambda x: f'''=SUMIFS(tb_list[Hours a Day],tb_list[Start Date],
                                        "<="&A{x['aux']},tb_list[Finish Date],">="&A{x['aux']},tb_list[Year],
                                        days!G{x['aux']})+SUMIFS(tb_list[Hours a Day],tb_list[Start Date],
                                        "<="&A{x['aux']},tb_list[Finish Date],">="&A{x['aux']},tb_list[Year],
                                        days!G{x['aux']}-1)''', axis=1)
    df_days["finished shows"] = df_days.agg(lambda x: f'''=COUNTIF(user_list!G:G,A{x['aux']})''', axis=1)
    df_days["month"] = df_days.agg(lambda x: f'''=MONTH(A{x['aux']})''', axis=1)
    df_days["year"] = df_days.agg(lambda x: f'''=YEAR(A{x['aux']})''', axis=1)

    df_days.drop(columns=["aux"], inplace=True)


    return df_days

# todo in these function of resizig the tables it's probably a good idea to make a dataframe with the formulae and then
#  add to excel instead of iterating the through every cell
# updates table size to fit the new entries
def resize_table_progression(workbook, oldest: int):
    sheet = workbook['progression']
    table = sheet.tables['tb_progression_month']
    #to add border when year changes
    border = Border(bottom=Side(style='medium'))

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
        # add border when year change
        if i == 12:
            tb_range = sheet['B{}'.format(row):'I{}'.format(row)]
            for cell in tb_range:
                for x in cell:
                    x.border = border

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
                # add border when year change
                if i == 12:
                    tb_range = sheet['B{}'.format(row):'I{}'.format(row)]
                    for cell in tb_range:
                        for x in cell:
                            x.border = border

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
    cellseasonep = "N{}".format(seasonrow)
    sheet[cellseasonep] = '''=SUMIFS(tb_progression_month[ep],tb_progression_month[month2],
    IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
    IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
    L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
    cellseasonhours = "O{}".format(seasonrow)
    sheet[cellseasonhours] = '''=SUMIFS(tb_progression_month[Hours Watched],tb_progression_month[month2],
    IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
    IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
    L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
    cellseasonfinish = "P{}".format(seasonrow)
    sheet[cellseasonfinish] = '''=SUMIFS(tb_progression_month[finished shows],tb_progression_month[month2],
    IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
    IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
    L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
    cellseasonpercentage = "Q{}".format(seasonrow)
    sheet[cellseasonpercentage] = '''=IF(O{}>0,SUMIFS(days!D:D,days!F:F,IF(M{}="WINTER","<=3",
    IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),days!F:F,IF(M{}="WINTER",">=1",
    IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),days!G:G,L{})/O{},0)'''.format(seasonrow, seasonrow,
                                                                                         seasonrow, seasonrow,
                                                                                         seasonrow, seasonrow,
                                                                                         seasonrow, seasonrow,
                                                                                         seasonrow)
    cellseasonhoursday = "R{}".format(seasonrow)
    # I got lazy and considered each month month with 30 days
    sheet[cellseasonhoursday] = '''=SUMIFS(tb_progression_month[Hours Watched],tb_progression_month[month2],
    IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
    IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],L{})
    /90'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow)
    seasonrow = seasonrow+1

    for i in range(2, 5):

        cellseason = "M{}".format(seasonrow)
        if i == 2:
            sheet[cellseason] = "SPRING"
        elif i == 3:
            sheet[cellseason] = "SUMMER"
        elif i == 4:
            sheet[cellseason] = "FALL"
            tb_range = sheet['L{}'.format(seasonrow):'R{}'.format(seasonrow)]
            # add border when year change
            for cell in tb_range:
                for x in cell:
                    x.border = border

        cellseasonep = "N{}".format(seasonrow)
        sheet[cellseasonep] = '''=SUMIFS(tb_progression_month[ep],tb_progression_month[month2],
        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow-i+1)
        cellseasonhours = "O{}".format(seasonrow)
        sheet[cellseasonhours] = '''=SUMIFS(tb_progression_month[Hours Watched],tb_progression_month[month2],
        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow-i+1)
        cellseasonfinish = "P{}".format(seasonrow)
        sheet[cellseasonfinish] = '''=SUMIFS(tb_progression_month[finished shows],tb_progression_month[month2],
        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow-i+1)
        cellseasonpercentage = "Q{}".format(seasonrow)
        sheet[cellseasonpercentage] = '''=IF(O{}>0,SUMIFS(days!D:D,days!F:F,IF(M{}="WINTER","<=3",
        IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),days!F:F,IF(M{}="WINTER",">=1",
        IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),days!G:G,L{})/O{},0)'''.format(seasonrow, seasonrow,
                                                                                             seasonrow, seasonrow,
                                                                                             seasonrow, seasonrow,
                                                                                             seasonrow, seasonrow-i+1,
                                                                                             seasonrow)
        cellseasonhoursday = "R{}".format(seasonrow)
        # I got lazy and considered each month month with 30 days
        sheet[cellseasonhoursday] = '''=SUMIFS(tb_progression_month[Hours Watched],tb_progression_month[month2],
        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],L{})
        /90'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow-i+1)
        seasonrow = seasonrow + 1

    seasonyear = seasonyear + 1
    if oldest <= today_year:
        while seasonyear <= today_year:

            for i in range(1,5):


                cellseason = "M{}".format(seasonrow)
                if i == 1:
                    sheet[cellseason] = "WINTER"
                    cellseasonyear = "L{}".format(seasonrow)
                    sheet[cellseasonyear] = "=L{}+1".format(seasonrow - 4)
                elif i ==2:
                    sheet[cellseason] = "SPRING"
                elif i == 3:
                    sheet[cellseason] = "SUMMER"
                elif i == 4:
                    sheet[cellseason] = "FALL"
                    tb_range = sheet['L{}'.format(seasonrow):'R{}'.format(seasonrow)]
                    # add border when year change
                    for cell in tb_range:
                        for x in cell:
                            x.border = border

                cellseasonep = "N{}".format(seasonrow)
                sheet[cellseasonep] = '''=SUMIFS(tb_progression_month[ep],tb_progression_month[month2],
                        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
                        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
                        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow-i+1)
                cellseasonhours = "O{}".format(seasonrow)
                sheet[cellseasonhours] = '''=SUMIFS(tb_progression_month[Hours Watched],tb_progression_month[month2],
                        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
                        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
                        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow-i+1)
                cellseasonfinish = "P{}".format(seasonrow)
                sheet[cellseasonfinish] = '''=SUMIFS(tb_progression_month[finished shows],tb_progression_month[month2],
                        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
                        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],
                        L{})'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow-i+1)
                cellseasonpercentage = "Q{}".format(seasonrow)
                sheet[cellseasonpercentage] = '''=IF(O{}>0,SUMIFS(days!D:D,days!F:F,IF(M{}="WINTER","<=3",
                        IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),days!F:F,IF(M{}="WINTER",">=1",
                        IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),days!G:G,L{})/O{},
                        0)'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow,
                                     seasonrow-i+1, seasonrow)
                cellseasonhoursday = "R{}".format(seasonrow)
                # I got lazy and considered each month month with 30 days
                sheet[cellseasonhoursday] = '''=SUMIFS(tb_progression_month[Hours Watched],tb_progression_month[month2],
                        IF(M{}="WINTER","<=3",IF(M{}="SPRING","<=6",IF(M{}="SUMMER","<=9","<=12"))),tb_progression_month[month2],
                        IF(M{}="WINTER",">=1",IF(M{}="SPRING",">=4",IF(M{}="SUMMER",">=7",">=10"))),tb_progression_month[year],L{})
                        /90'''.format(seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow, seasonrow-i+1)

                seasonrow = seasonrow +1
            seasonyear = seasonyear + 1

    sheet.tables['tb_progression_season'].ref = 'L2:R' + str(seasonrow - 1)

    # i had to add the conditional formatting in the code because even though the table size is increase the conditional
    # formatting remains the same
    # Apparently if you enter the entire column in the conditional formatting you don't need this but i'll leave it here
    # for the month table the season has de conditional formatting in the entire column

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

# I don't know who to get a list of all genres so I made a list of all genres in the user_list and then add to the tables
# and since a new genre may be add i have to do it every time
def resize_challenge_tables(workbook, genreslist):
    sheet = workbook['challenge']
    row = 3
    for genre in genreslist:
        cellgenres = "J{}".format(row)
        sheet[cellgenres] = genre
        celltotal = "K{}".format(row)
        sheet[celltotal] = '''=COUNTIFS(tb_anime_genres[Genres],J{}, tb_anime_genres[Start Date],
        ">=" & $B$1,tb_anime_genres[Finish Date],"<="&$B$2)'''.format(row)
        cellpercentage = "L{}".format(row)
        sheet[cellpercentage] = '=K{}/$C$7'.format(row)
        sheet[cellpercentage].number_format = "0.00%"

        cellaverage = "M{}".format(row)
        sheet[cellaverage] = '''=IF(K{}=0,0,AVERAGEIFS(tb_anime_genres[Score],tb_anime_genres[Genres],J{}, 
        tb_anime_genres[Start Date],">=" & $B$1,tb_anime_genres[Finish Date],"<="&$B$2))'''.format(row, row)
        cellhours = "N{}".format(row)
        sheet[cellhours] = '''=SUMIFS(tb_anime_genres[Show Duration (hours)],tb_anime_genres[GENRES],J{}, tb_anime_genres[Start Date],
        ">=" & $B$1,tb_anime_genres[Finish Date],"<="&$B$2)'''.format(row)
        cellweightedmean = "O{}".format(row)
        sheet[cellweightedmean] = '''=IF(K{}=0,0,M{}*(K{}/$C$7)+$C$16*($C$7-K{})/$C$7)'''.format(row, row, row, row)
        row = row+1

    sheet.tables['tb_challenge'].ref = 'I2:P' + str(row - 1)

    sheet = workbook['genres_stats']
    row = 2
    for genre in genreslist:
        cellgenres = "A{}".format(row)
        sheet[cellgenres] = genre
        celltotal = "B{}".format(row)
        sheet[celltotal] = '''=COUNTIF(tb_anime_genres[Genres],A{})'''.format(row)
        cellpercentage = "C{}".format(row)
        sheet[cellpercentage] = '=B{}/COUNTIF(tb_list[Title],"<>"&"")'.format(row)
        sheet[cellpercentage].number_format="0.00%"
        cellaverage = "D{}".format(row)
        sheet[cellaverage] = '''=AVERAGEIFS(tb_anime_genres[Score],tb_anime_genres[Genres],A{})'''.format(row)
        cellhours = "E{}".format(row)
        sheet[cellhours] = '''=SUMIFS(tb_anime_genres[Show Duration (hours)],tb_anime_genres[Genres],A{})'''.format(row)
        cellweightedmean = "F{}".format(row)
        sheet[cellweightedmean] = '''=D{}*(B{}/COUNTIF(tb_list[Title], "<>"&""))+
        AVERAGEIF(tb_list[Score],"<>"&0,tb_list[Score])*(COUNTIF(tb_list[Title],"<>"&"")-
        B{})/COUNTIF(tb_list[Title],"<>"&"")'''.format(row, row, row)
        row = row + 1

    sheet.tables['tb_genres'].ref = 'A1:F' + str(row - 1)

    return workbook

