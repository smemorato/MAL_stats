import mal_api
import anilist_api
from  mal_token import refresh_token, get_new_code_verifier,print_new_authorisation_url,\
    generate_new_token,print_user_info, CLIENT_ID
import json
import mal_to_excel

mal = True
username = "dasdas"


if mal:
    try:
        with open('token.json', 'r') as file:
            print(file)
            token = json.load(file)
            token = refresh_token(token)

    except Exception as err:
        print(err)
        code_verifier = code_challenge = get_new_code_verifier()
        print_new_authorisation_url(code_challenge)
        authorisation_code = input('Copy-paste the Authorisation Code: ').strip()
        token = generate_new_token(authorisation_code, code_verifier)
        print_user_info(token['access_token'])

    response = mal_api.request_list(username, token["access_token"])
    mal_to_excel.to_excel(response[0], username)
else:
    response = anilist_api.get_user_id("bob")
