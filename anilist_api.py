import requests
import json
import pandas


def get_user_id (username: str):
    query = '''
    query ($name: String,) { 
      User (name: $name ) { 
        id
        }   
    
    }
    '''
    variables = {
        'name': username
    }


    url = 'https://graphql.anilist.co'

    # Make the HTTP Api request
    response = requests.post(url, json={'query': query, 'variables': variables})
    return response



def request_list(user_id: int):
    query = '''
    query ($userId: Int,) { # Define which variables will be used in the query (id)
      MediaListCollection (userId: $userId, type: ANIME,status: COMPLETED) { 
          lists{
            entries{
                status
                score
                startedAt{
                    year
                    month
                    day
                }
                completedAt{
                    year
                    month
                    day
                }
                media{
                    id
                    title{
                        romaji
                        english
                    }
                    status
                    episodes
                    format
                    seasonYear
                    season
                    source
                    tags{
                        name
                    }
                    duration
                    meanScore
                    genres
                    startDate{
                        year
                        month
                    }
                }
            }
            status
        }
      }
    }
    '''

    # Define our query variables and values that will be used in the query request
    variables = {
        'userId': user_id
    }
    url = 'https://graphql.anilist.co'

    # Make the HTTP Api request
    response = requests.post(url, json={'query': query, 'variables': variables})


    return response


