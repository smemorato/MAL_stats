import argparse
import json

import anilist_api
from  mal_token import refresh_token, get_new_code_verifier,print_new_authorisation_url,\
    generate_new_token,print_user_info, CLIENT_ID
import mal_api
import mal_to_excel
import anilist_to_excel



if __name__ == "__main__":
    parser = argparse.ArgumentParser(     prog='PROG',
        description='''Create an excel file with the stats of a given user''',
        epilog='''
        Let's hope  it works!!!''')


    parser.add_argument("-u", "--username",
                        required= True,
                        help="type username",
                        metavar= " ")
    parser.add_argument("-s", "--source",
                        required= True,
                        choices = ["anilist","mal"],
                        default="mal",
                        help="pick the source of the username either mal or anilist (default: %(default)s)'",
                        metavar= " ")
    args = parser.parse_args()


    if args.source == "mal":
        print("mal selected")
        if CLIENT_ID is None:
            print("please add a client id")
        else:
            # try to open token file generate new if file not found
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

            response = mal_api.request_list(args.username, token["access_token"])

            if response[1] == 200:
                mal_to_excel.to_excel(response[0], args.username)
            elif response[1] == 404:
                print("username doesn't exist")
                print("please insert a valid username")
                print(f'error {response[1]}')
            else:

                print(response[1])

    elif args.source == "anilist":
        response = anilist_api.get_user_id(args.username)
        if response.status_code == 200:
            data_id = json.loads(response.text)

            user_id = data_id['data']['User']['id']
            userlist = anilist_api.request_list(user_id)

            anilist_to_excel.to_excel(userlist,args.username)

        elif response.status_code == 404:
            print("username doesn't exist")
            print("please insert a valid username")
            print(response.text)
        else:
            print(response.text)

