from  mal_token import refresh_token, get_new_code_verifier,print_new_authorisation_url,\
    generate_new_token,print_user_info, CLIENT_ID
import json
import requests
import webbrowser



def request_plan_watch_number (season, year, access_token: str):


    url = f'https://api.myanimelist.net/v2/anime/season/{year}/{season}'
    PARAMS  = {
        "limit" : "200",
        "fields" : "id,title,main_picture,alternative_titles,start_date,end_date,synopsis,mean,rank,popularity,num_list_users,num_scoring_users,nsfw,created_at,updated_at,media_type,status,genres,my_list_status,num_episodes,start_season,broadcast,source,average_episode_duration,rating,pictures,background,related_anime,related_manga,recommendations,studios,statistics'"
    }
    #resquest frist page of the list
    response = requests.get(url, params = PARAMS,headers={
        'Authorization': f'Bearer {access_token}'
    })
    #print(json.dumps(response.json(),indent=3))
    #requests following pages
    page = response.json()


    return page

def request_anime_status (id, access_token: str):

    url = f'https://api.myanimelist.net/v2/anime/{id}'
    PARAMS  = {
        "fields" : "mean,statistics,start_season"
        


    }
    #resquest frist page of the list
    response = requests.get(url, params = PARAMS,headers={
        'Authorization': f'Bearer {access_token}'
    })
    #print(json.dumps(response.json(),indent=3))
    #requests following pages
    page = response.json()


    return page




if CLIENT_ID is None:
    print("please add a client id")
else:
    try:
        with open('token.json', 'r') as file:
            print(file)
            token = json.load(file)
            token = refresh_token(token)

    except Exception as err:
        print(err)
        # todo make this a function in mal_api
        code_verifier = code_challenge = get_new_code_verifier()
        print_new_authorisation_url(code_challenge)
        authorisation_code = input('Copy-paste the Authorisation Code: ').strip()
        token = generate_new_token(authorisation_code, code_verifier)
        print_user_info(token['access_token'])
        
    file = open('items.txt','w')
    list1=[]
    for x in range (2001,2005):
        for y in range (1,5):
            if y == 1:
                season = "winter"
            elif y == 2:
                season= "spring"
            elif y == 3:
                season = "summer"
            elif y ==4:
                season = "fall"
            print(season)
            print(x)
            response = request_plan_watch_number(season,x, token["access_token"])
            print(response)
            #print(json.dumps(response, indent=0))
            for entry in response["data"]:
                #response2 = request_anime_status(entry["node"]["id"], token["access_token"])
                #print(json.dumps(response2, indent=2))
                try:
                    if int(response2["statistics"]["status"]["completed"])< int(response2["statistics"]["status"]["plan_to_watch"]):
                        try:
                            
                            try:
                                if response2["mean"]>6.5:
                                    
                                    list1.append(response2["id"])
                                    file.write(str(response2["id"])+"\n")
                            except:
                                print(response2)
                        except:
                            print(response2)
                
                except:
                    print(response2)


    if response[1] == 200:
        print("ok")

    elif response[1] == 404:
        print("username doesn't exist")
        print("please insert a valid username")
        print(f'error {response[1]}')
    else:

        print(response[1])