import os
import io
from datetime import datetime
from xml.etree.ElementTree import Element, SubElement, tostring

import psycopg2
from psycopg2.extras import RealDictCursor
from flask import Flask, request, render_template, Response
from openpyxl import Workbook
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)


MATCH_TYPE_MAP = {
    "P": "Practice",
    "Q": "Qualification",
    "M": "Match",
    "F": "Final"
}

FIELD_ROWS = [


("Autonomous","Shooter Type","autonomousShootertype"),
("Autonomous","Shots Taken","autonomousShotstaken"),
("Autonomous","Shot Volumes","autonomousShotvolumes"),
("Autonomous","Subjective Accuracy","autonomousSubjectiveaccuracy"),

("Teleop","Fuel Count","teleopFuelcount"),
("Teleop","Human Fuel Count","teleopHumanfuelcount"),
("Teleop","Pass Bump","teleopPassbump"),
("Teleop","Pass Trench","teleopPasstrench"),
("Teleop","Fetch Ball Preference","teleopFetchballpreference"),
("Teleop","Shots Taken","teleopShotstaken"),
("Teleop","Shot Volumes","teleopShotvolumes"),
("Teleop","Subjective Accuracy","teleopSubjectiveaccuracy"),

("End&AfterGame","Tower Status","endAndAfterGameTowerstatus"),
("End&AfterGame","Climbing Time","endAndAfterGameClimbingtime"),
("End&AfterGame","Ranking Points","endAndAfterGameRankingpoint"),
("End&AfterGame","Coop Point","endAndAfterGameCooppoint"),
("End&AfterGame","Autonomous Move","endAndAfterGameAutonomousmove"),
("End&AfterGame","Teleop Move","endAndAfterGameTeleopmove"),
("End&AfterGame","Comments","endAndAfterGameComments"),

]

def get_db():

    return psycopg2.connect(
        host=os.getenv("POSTGRES_HOST","localhost"),
        port=os.getenv("POSTGRES_PORT","5432"),
        user=os.getenv("POSTGRES_USER","postgres"),
        password=os.getenv("POSTGRES_PASSWORD","postgres"),
        dbname=os.getenv("POSTGRES_DB","postgres")
    )


def format_value(v):

    if v is None:
        return ""

    if isinstance(v,bool):
        return "Yes" if v else "No"

    return str(v)

def get_value(record, key):
    return record.get(key)


    if key.startswith("teleop"):

        tele=record.get("teleop") or {}

        mapping={
        "teleopFuelCount":"fuelCount",
        "teleopHumanFuelCount":"humanFuelCount",
        "teleopPassBump":"passBump",
        "teleopPassTrench":"passTrench",
        "teleopFetchBallPreference":"fetchBallPreference",
        "teleopShotsTaken":"shotsTaken",
        "teleopShotVolumes":"shotVolumes",
        "teleopSubjectiveAccuracy":"subjectiveAccuracy"
        }

        return tele.get(mapping[key])


    if key.startswith("endAndAfterGame"):

        end=record.get("endAndAfterGame") or {}

        mapping={
        "endAndAfterGameTowerStatus":"towerStatus",
        "endAndAfterGameClimbingTime":"climbingTime",
        "endAndAfterGameRankingPoint":"rankingPoint",
        "endAndAfterGameCoopPoint":"coopPoint",
        "endAndAfterGameAutonomousMove":"autonomousMove",
        "endAndAfterGameTeleopMove":"teleopMove",
        "endAndAfterGameComments":"comments"
        }

        return end.get(mapping[key])


    return record.get(key)


def load_match(event_id,match_type,match_number):

    db_type=MATCH_TYPE_MAP.get(match_type)

    conn=get_db()

    with conn.cursor(cursor_factory=RealDictCursor) as cur:

        cur.execute("""

        SELECT *
        FROM team_match_record
        WHERE "scoutEventId"=%s
        AND "matchType"=%s
        AND "matchNumber"=%s
        ORDER BY "alliance","teamNumber"

        """,(event_id,db_type,match_number))

        rows=cur.fetchall()

    conn.close()

    match={"red":[],"blue":[]}

    for r in rows:

        if "red" in str(r["alliance"]).lower():
            match["red"].append(r)
        else:
            match["blue"].append(r)

    print(rows[0].keys())

    return match


def build_matrix(match):

    columns=[]

    for t in match["red"]:
        columns.append({"label":f"Red {t['teamNumber']}","record":t})

    for t in match["blue"]:
        columns.append({"label":f"Blue {t['teamNumber']}","record":t})

    while len(columns)<6:

        if len(columns)<3:
            columns.append({"label":"Red","record":None})
        else:
            columns.append({"label":"Blue","record":None})

    rows=[]

    for cat,name,key in FIELD_ROWS:

        r=[cat,name]

        for c in columns:

            if c["record"]:
                v=get_value(c["record"],key)
                r.append(format_value(v))
            else:
                r.append("")

        rows.append(r)

    return columns,rows


def create_excel(match):

    wb=Workbook()
    ws=wb.active

    columns,rows=build_matrix(match)

    ws.cell(row=1,column=1,value="Section")
    ws.cell(row=1,column=2,value="Field")

    for i,c in enumerate(columns,start=3):
        ws.cell(row=1,column=i,value=c["label"])

    for r_i,row in enumerate(rows,start=2):

        for c_i,val in enumerate(row,start=1):
            ws.cell(row=r_i,column=c_i,value=val)

    bio=io.BytesIO()
    wb.save(bio)

    return bio.getvalue()


def create_xml(match):

    root=Element("TeamMatchRecord")

    columns,rows=build_matrix(match)

    headers=SubElement(root,"Headers")

    SubElement(headers,"Header").text="Section"
    SubElement(headers,"Header").text="Field"

    for c in columns:
        SubElement(headers,"Header").text=c["label"]

    body=SubElement(root,"Rows")

    for row in rows:

        r=SubElement(body,"Row")

        for cell in row:
            SubElement(r,"Cell").text=str(cell)

    return tostring(root,encoding="utf-8",xml_declaration=True)


@app.route("/")
def index():

    event_id=request.args.get("scout_event_id")
    match_type=request.args.get("match_type")
    match_number=request.args.get("match_number")

    columns=[]
    rows=[]

    if event_id and match_type and match_number:

        match=load_match(event_id,match_type,int(match_number))
        columns,rows=build_matrix(match)

    return render_template(
        "index.html",
        columns=columns,
        rows=rows,
        scout_event_id=event_id,
        match_type=match_type,
        match_number=match_number
    )


@app.get("/export/excel")
def export_excel():

    event_id=request.args.get("event_id")
    match_type=request.args.get("match_type")
    match_number=int(request.args.get("match_number"))

    match=load_match(event_id,match_type,match_number)

    data=create_excel(match)

    name=f"match_{match_number}_{datetime.now().strftime('%Y%m%d')}.xlsx"

    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":f'attachment; filename="{name}"'}
    )


@app.get("/export/xml")
def export_xml():

    event_id=request.args.get("event_id")
    match_type=request.args.get("match_type")
    match_number=int(request.args.get("match_number"))

    match=load_match(event_id,match_type,match_number)

    data=create_xml(match)

    name=f"match_{match_number}_{datetime.now().strftime('%Y%m%d')}.xml"

    return Response(
        data,
        mimetype="application/xml",
        headers={"Content-Disposition":f'attachment; filename="{name}"'}
    )


if __name__=="__main__":

    app.run(host="0.0.0.0",port=5000,debug=True)