import os

import pandas as pd
import tbapy

import csv

tba = tbapy.TBA('opXlAfkuD4tQbDm2iskpBHdyYQbarWsQoeSG8w6MSKQ0c8jtbOnbREQu7z7nfUCK')


def removeFile(file):
    if os.path.exists(file):
        os.remove(file)


def getMatch(EVENT):
    removeFile("Match_Folder/Match Schedule_" + EVENT + ".csv")

    schedule_keys = tba.event_matches(EVENT, simple=True, keys=True)
    df = pd.concat([pd.DataFrame([i], columns=['Match #']) for i in schedule_keys], ignore_index=True)
    df = df[df['Match #'].str.contains("qm")]
    df['sort'] = df['Match #'].str.extract('([^qm]*$)', expand=False).astype(int)
    df.sort_values('sort', inplace=True, ascending=True)
    df = df.drop('sort', axis=1)

    with open("Match Schedule_" + EVENT + ".csv", "w", newline="") as f:
        headers = ['Match #', 'Red 1', 'Red 2', 'Red 3', 'Blue 1', 'Blue 2', 'Blue 3']
        writer = csv.DictWriter(f, fieldnames=headers)

        writer.writeheader()

        for m in df['Match #']:
            mat = tba.match(m, simple=True)
            writer.writerow({'Match #': mat.match_number,
                             'Red 1': mat.alliances['blue']['team_keys'][0][3:],
                             'Red 2': mat.alliances['blue']['team_keys'][1][3:],
                             'Red 3': mat.alliances['blue']['team_keys'][2][3:],
                             'Blue 1': mat.alliances['red']['team_keys'][0][3:],
                             'Blue 2': mat.alliances['red']['team_keys'][1][3:],
                             'Blue 3': mat.alliances['red']['team_keys'][2][3:]})
        f.close()


# schedule = ['2022azva', '2022cafr']
schedule = ['2021mttd']
for i in schedule:
    getMatch(i)
