import requests
import json
import pandas as pd

def request_list (username: str, access_token: str):


    url = f'https://api.myanimelist.net/v2/users/{username}/animelist'
    PARAMS  = {
        'fields': "list_status,start_date,end_date,"+
                  "mean,media_type"+
                  ",genres,num_episodes,start_season,source,average_episode_duration,"+
                  "studios",

        # 'fields': "list_status,alternative_titles,start_date,end_date," +
        #           "mean,rank,popularity,num_list_users,num_scoring_users,nsfw,created_at,updated_at,media_type" +
        #           ",status,genres,num_episodes,start_season,broadcast,source,average_episode_duration," +
        #           "rating,pictures,background,related_anime,related_manga,recommendations,studios,statistics",

        'status': 'completed',
        'limit': '1000',
        "nsfw": "true"


    }
    #resquest frist page of the list
    response = requests.get(url, params = PARAMS,headers={
        'Authorization': f'Bearer {access_token}'
    })

    #requests following pages
    page = response.json()
    userlist = []
    userlist.append(page)
    response.close()








    response.raise_for_status()
    while "next" in page["paging"].keys():
        response = requests.get(page["paging"]["next"], params=PARAMS, headers={
            'Authorization': f'Bearer {access_token}'
        })
        page = response.json()
        #print(json.dumps(page, indent=2))
        userlist.append(page)
        print("another one!")


    response_list= [userlist, response.status_code]

    return response_list





def anime_info(access_token: str):
    url = f'https://api.myanimelist.net/v2/anime/30230'

    PARAMS = {
        'fields'

    }

    response = requests.get(url, params=PARAMS, headers={
        'Authorization': f'Bearer {access_token}'
    })
    print(response.url)
    page = response.json()

    response.close()
    # print(json.dumps(page, indent=2))
    response.raise_for_status()


def get_genre(access_token: str) -> list:
    url = f'https://api.myanimelist.net/v2/genres'

    PARAMS = {
        "fields": "genres"
    }

    response = requests.get(url, headers={
        'Authorization': f'Bearer {access_token}'
    })
    # print(json.dumps(response.json(), indent=2))
    print(response.raise_for_status())


def ani_to_mal(access_token):
    url = f'https://api.myanimelist.net/v2'
    PARAMS = {
        'anime_id': "34223",
        'fields': "list_status,alternative_titles,start_date,end_date",

    }

    response = requests.get(url, params=PARAMS, headers={
        'Authorization': f'Bearer {access_token}'
    })

    page = response.json()
    userlist = []
    userlist.append(page)
    response.close()
    # print(json.dumps(userlist, indent=2))
    response.raise_for_status()
