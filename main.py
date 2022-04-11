import requests

def get_data():
    r = requests.get("https://line13.bkfon-resources.com/live/currentLine/ru")
    convert = r.json()
    id_list=[]
    name_list=[]
    team1_list=[]
    team2_list=[]
    final_id_list=[]
    final_name= []


    for sport in convert['sports']:
        id_list.append(sport['id'])
        name_list.append(sport['name'])

    for event in convert['events']:
        if event['sportId'] in id_list and event['name'] == '' and event.get('team2'):

                team1_list.append(event['team1'])

                team2_list.append(event['team2'])
                final_id_list.append(event['id'])


    for i in range(len(final_id_list)):
        var=final_id_list[i]
        if var in id_list:
            index = id_list.index(var)
            time_index = id_list.index(var)
            final_name.append(name_list[index])


    for i in range(len(name_list)):
        print(f'{final_id_list[i]} {team1_list[i]} {team2_list[i]} {name_list[i]}')

    return final_id_list, team1_list, team2_list, name_list

get_data()





