import os.path
import requests
import time
import random
import json
from openpyxl import Workbook
from requests import session

URL = input("Enter the URL of the webpage: ")
Filename = input("What do you want the Excel worksheet's file name to be? ")
Headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer {INSERT YOUR API KEY HERE}'
}
Request_Delay = (2, 15)
Max_Match_Results = 10
Retrys = 5
Timeout = 15


def get_all_teams(eventid):
    all_teams = []
    page = 1

    while True:
        url = f"https://www.robotevents.com/api/v2/events/{eventid}/teams?grade%5B%5D=High%20School&grade%5B%5D=Middle%20School"
        url += f"&page={page}"

        response = requests.get(url, headers=Headers)
        if response.status_code == 429:
            while True:
                time.sleep(Timeout)
                response = requests.get(url, headers=Headers)
                if response.status_code == 200:
                    break
        json_data = response.json()
        all_teams.extend(json_data["data"])
        if json_data["meta"]["current_page"] >= json_data["meta"]["last_page"]:
            break
        page += 1
    return all_teams


def get_teams(url):
    attempts = 0
    while attempts < Retrys:
        attempts+=1
        sessiona = requests.Session()
        sessiona.headers.update(Headers)
        print("")
        event_code = url.split("/")[-1].split(".")[0]
        print(event_code)
        eventcodelink = f'https://www.robotevents.com/api/v2/events?sku[]={event_code}'
        eventcode = sessiona.get(eventcodelink)
        if eventcode.status_code == 200:
            json_data = eventcode.json()
            data_table = json_data["data"][0]
            eventid = data_table["id"]
            print(f'The event id is {eventid}')
            data = get_all_teams(eventid)
            if data:
                print("Successfully got event data")
                print("Getting all team data, may take a while (depends on how many teams there are)")
                return data
        else:
            print(f"Failed to get 200 status code in {attempts} attempts.")
    return None


def get_best_ranking(team_id):
    url = f"https://www.robotevents.com/api/v2/teams/{team_id}/rankings?season%5B%5D=197"

    best_rank = None
    while True:
        response = requests.get(url, Headers)
        if response.status_code == 429:
            while True:
                time.sleep(Timeout)
                response = requests.get(url, Headers)
                if response.status_code == 200:
                    break
        try:
            response = response.json()
        except:
            response = {"data": [], "meta": {}}
        if not isinstance(response, dict):
            response = {"data": [], "meta": {}}
        data = response.get("data", [])
        for r in data:
            rank = r.get("rank")
            if rank is not None:
                if best_rank is None or rank < best_rank:
                    best_rank = rank
        next_page = response.get("meta", {}).get("next_page_url")
        if not next_page:
            break
        url = next_page

    return best_rank


def get_team_data(event_data):
    results = []
    for team in event_data:
        t_id = team.get("id")
        if not t_id:
            continue
        nsession = requests.Session()
        nsession.headers.update(Headers)

        base = f"https://www.robotevents.com/api/v2/teams/{t_id}/skills?season%5B%5D=197"
        skills_total = base
        skills_driver = base + "&type%5B%5D=driver"
        skills_programming = base + "&type%5B%5D=programming"

        def get_highest(session2, nurl):
            try:
                resp = session2.get(nurl, headers=Headers, timeout=Timeout)
                if resp.status_code == 429:
                    while True:
                        time.sleep(Timeout)
                        resp = session2.get(nurl, headers=Headers, timeout=Timeout)
                        if resp.status_code == 200:
                            break
                if not resp.text.strip():
                    return 0
                json_data = resp.json()
            except Exception:
                return 0
            if not isinstance(json_data, dict):
                return 0
            runs = json_data.get("data", [])
            if not runs:
                return 0
            return max(v.get("score", 0) for v in runs if isinstance(v, dict))

        highest_driver = get_highest(nsession, skills_driver)
        highest_programming = get_highest(nsession, skills_programming)
        total = highest_driver + highest_programming

        awards = []
        url = f"https://www.robotevents.com/api/v2/teams/{t_id}/awards?season=197&per_page=250"
        while True:
            try:
                aj = nsession.get(url, headers=Headers, timeout=Timeout)
                if aj.status_code == 429:
                    while True:
                        time.sleep(Timeout)
                        aj = nsession.get(url, headers=Headers, timeout=Timeout)
                        if aj.status_code == 200:
                            break
                if not aj.text.strip():
                    aj_json = {"data": [], "meta": {}}
                else:
                    aj_json = aj.json()
            except Exception:
                aj_json = {"data": [], "meta": {}}
            if not isinstance(aj_json, dict):
                aj_json = {"data": [], "meta": {}}
            for a in aj_json.get("data", []):
                awards.append(a.get("title", ""))
            nxt = aj_json.get("meta", {}).get("next_page_url")
            if not nxt:
                break
            url = nxt

        best_rank = None
        url = f"https://www.robotevents.com/api/v2/teams/{t_id}/rankings?season=197&per_page=250"
        while True:
            try:
                rj = nsession.get(url, headers=Headers, timeout=Timeout)
                if rj.status_code == 429:
                    time.sleep(Timeout)
                    rj = nsession.get(url, headers=Headers, timeout=Timeout)
                    if rj.status_code == 200:
                        break
                if not rj.text.strip():
                    rj_json = {"data": [], "meta": {}}
                else:
                    rj_json = rj.json()
            except Exception:
                rj_json = {"data": [], "meta": {}}
            if not isinstance(rj_json, dict):
                rj_json = {"data": [], "meta": {}}
            for r in rj_json.get("data", []):
                rk = r.get("rank")
                if rk is not None:
                    if best_rank is None or rk < best_rank:
                        best_rank = rk
            nxt = rj_json.get("meta", {}).get("next_page_url")
            if not nxt:
                break
            url = nxt

        results.append({
            "team_id": t_id,
            "highest_total_skills": total,
            "highest_driver_skills": highest_driver,
            "highest_programming_skills": highest_programming,
            "awards": ", ".join(awards),
            "best_rank": best_rank
        })
    return results


def save_teams_to_excel(event_data, team_data, filename):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append([
        "Team ID", "Team Number", "Team Name", "Organization", "Grade", "Location",
        "Highest Total Skills Score", "Highest Driver", "Highest Programming",
        "Best Rank at a Tournement", "Awards"
    ])

    td_map = {t["team_id"]: t for t in team_data}

    for event in event_data:
        t_id = event.get("id")
        location_data = event.get("location", {})
        location_parts = [
            location_data.get("city"),
            location_data.get("region"),
            location_data.get("country")
        ]
        location_string = ", ".join([p for p in location_parts if p])

        t = td_map.get(t_id, {})

        worksheet.append([
            event.get("id"),
            event.get("number"),
            event.get("team_name"),
            event.get("organization"),
            event.get("grade"),
            location_string,
            t.get("highest_total_skills", ""),
            t.get("highest_driver_skills", ""),
            t.get("highest_programming_skills", ""),
            t.get("best_rank", ""),
            t.get("awards", "")
        ])

    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass
        adjusted_width = max_length + 2
        worksheet.column_dimensions[column_letter].width = adjusted_width

    workbook.save(f'{filename}.xlsx')


if __name__ == "__main__":
    edata = get_teams(URL)
    tdata = get_team_data(edata)
    save_teams_to_excel(edata, tdata, Filename)

