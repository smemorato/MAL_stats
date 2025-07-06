import requests
import json



def request_list (username: str, access_token: str):


    url = f'https://api.myanimelist.net/v2/users/{username}/animelist'
    PARAMS  = {
        'fields': "list_status,start_date,end_date,"+
                  "mean,media_type"+
                  ",genres,num_episodes,start_season,source,average_episode_duration,"+
                  "studios",
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
    if response.status_code == 200:
        while "next" in page["paging"].keys():
            response = requests.get(page["paging"]["next"], params=PARAMS, headers={
                'Authorization': f'Bearer {access_token}'
            })
            page = response.json()
            userlist.append(page)


    response_list = [userlist, response.status_code]

    return response_list





def anime_info(access_token: str):
    url = f'https://api.myanimelist.net/v2/anime/30230'

    PARAMS = {
        'fields'

    }

    response = requests.get(url, params=PARAMS, headers={
        'Authorization': f'Bearer {access_token}'
    })
    page = response.json()

    response.close()
    response.raise_for_status()


def get_genre(access_token: str) -> list:
    url = f'https://api.myanimelist.net/v2/genres'

    PARAMS = {
        "fields": "genres"
    }

    response = requests.get(url, headers={
        'Authorization': f'Bearer {access_token}'
    })


##todo from a list from anilist use information from mal
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
    response.raise_for_status()
