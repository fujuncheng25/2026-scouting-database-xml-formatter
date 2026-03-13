import io
import os
from collections import OrderedDict
from datetime import datetime
from typing import Any
from xml.etree.ElementTree import Element, SubElement, tostring

import psycopg2
from dotenv import load_dotenv
from flask import Flask, Response, render_template, request
from openpyxl import Workbook
from psycopg2.extras import RealDictCursor

load_dotenv()

app = Flask(__name__)


FIELD_ROWS = [
    ("Info", "id", "id"),
    ("Info", "matchNumber", "match_number"),
    ("Info", "matchType", "match_type"),
    ("Info", "alliance", "alliance"),
    ("Info", "teamNumber", "team_number"),
    ("Info", "matchKey", "match_key"),
    ("Autonomous", "Shooter Type", "auto_shooter_type"),
    ("Autonomous", "Shots Taken", "auto_shots_taken"),
    ("Autonomous", "Shot Volumes", "auto_shot_volumes"),
    ("Autonomous", "Subjective Accuracy", "auto_subjective_accuracy"),
    ("Teleop", "Fuel Count", "teleop_fuel_count"),
    ("Teleop", "Human Fuel Count", "teleop_human_fuel_count"),
    ("Teleop", "Pass Bump", "teleop_pass_bump"),
    ("Teleop", "Pass Trench", "teleop_pass_trench"),
    ("Teleop", "Fetch Ball Preference", "teleop_fetch_ball_preference"),
    ("Teleop", "Shots Taken", "teleop_shots_taken"),
    ("Teleop", "Shot Volumes", "teleop_shot_volumes"),
    ("Teleop", "Subjective Accuracy", "teleop_subjective_accuracy"),
    ("End&AfterGame", "Tower Status", "end_tower_status"),
    ("End&AfterGame", "Climbing Time", "end_climbing_time"),
    ("End&AfterGame", "Ranking Points", "end_ranking_point"),
    ("End&AfterGame", "Cop Point", "end_coop_point"),
    ("End&AfterGame", "Autonomous Move", "end_autonomous_move"),
    ("End&AfterGame", "Teleop Move", "end_teleop_move"),
    ("End&AfterGame", "Game Comments", "end_comments"),
]


def quote_ident(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def get_db_connection():
    return psycopg2.connect(
        host=os.getenv("POSTGRES_HOST", "localhost"),
        port=int(os.getenv("POSTGRES_PORT", "5432")),
        user=os.getenv("POSTGRES_USER", "postgres"),
        password=os.getenv("POSTGRES_PASSWORD", "postgres"),
        dbname=os.getenv("POSTGRES_DB", "postgres"),
    )


def format_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    return str(value)


def get_value(row: dict[str, Any], candidates: list[str]) -> Any:
    lowered = {str(key).lower(): key for key in row.keys()}
    for name in candidates:
        if name in row:
            return row[name]
        lowered_name = name.lower()
        if lowered_name in lowered:
            return row[lowered[lowered_name]]
    return None


def get_nested_dict(row: dict[str, Any], candidates: list[str]) -> dict[str, Any]:
    value = get_value(row, candidates)
    return value if isinstance(value, dict) else {}


def pick_from_nested(payload: dict[str, Any], candidates: list[str]) -> Any:
    lowered = {str(key).lower(): key for key in payload.keys()}
    for name in candidates:
        if name in payload:
            return payload[name]
        lowered_name = name.lower()
        if lowered_name in lowered:
            return payload[lowered[lowered_name]]
    return None


def resolve_columns(column_names: set[str]) -> dict[str, str]:
    actual_by_lower = {name.lower(): name for name in column_names}

    def pick(*candidates: str, required: bool = False) -> str | None:
        for candidate in candidates:
            exact = actual_by_lower.get(candidate.lower())
            if exact is not None:
                return exact
        if required:
            raise RuntimeError(
                f"Required column not found. Expected one of: {', '.join(candidates)}"
            )
        return None

    return {
        "id": pick("id", required=True),
        "event": pick("scoutEventId", "scout_event_id", required=True),
        "match_number": pick("matchNumber", "match_number", required=True),
        "match_type": pick("matchType", "match_type", required=True),
        "alliance": pick("alliance", required=True),
        "team": pick("teamNumber", "team_number", required=True),
        "match_key": pick("matchKey", "match_key"),
        "autonomous_json": pick("autonomous"),
        "teleop_json": pick("teleop"),
        "end_json": pick("endAndAfterGame", "end_and_after_game"),
        "auto_shooter_type": pick("autonomousShooterType", "autonomous_shooter_type"),
        "auto_shots_taken": pick("autonomousShotsTaken", "autonomous_shots_taken"),
        "auto_shot_volumes": pick("autonomousShotVolumes", "autonomous_shot_volumes"),
        "auto_subjective_accuracy": pick(
            "autonomousSubjectiveAccuracy", "autonomous_subjective_accuracy"
        ),
        "teleop_fuel_count": pick("teleopFuelCount", "teleop_fuel_count"),
        "teleop_human_fuel_count": pick(
            "teleopHumanFuelCount", "teleop_human_fuel_count"
        ),
        "teleop_pass_bump": pick("teleopPassBump", "teleop_pass_bump"),
        "teleop_pass_trench": pick("teleopPassTrench", "teleop_pass_trench"),
        "teleop_fetch_ball_preference": pick(
            "teleopFetchBallPreference", "teleop_fetch_ball_preference"
        ),
        "teleop_shots_taken": pick("teleopShotsTaken", "teleop_shots_taken"),
        "teleop_shot_volumes": pick("teleopShotVolumes", "teleop_shot_volumes"),
        "teleop_subjective_accuracy": pick(
            "teleopSubjectiveAccuracy", "teleop_subjective_accuracy"
        ),
        "end_tower_status": pick("endAndAfterGameTowerStatus", "end_and_after_game_tower_status"),
        "end_climbing_time": pick("endAndAfterGameClimbingTime", "end_and_after_game_climbing_time"),
        "end_ranking_point": pick("endAndAfterGameRankingPoint", "end_and_after_game_ranking_point"),
        "end_coop_point": pick("endAndAfterGameCoopPoint", "end_and_after_game_coop_point"),
        "end_autonomous_move": pick(
            "endAndAfterGameAutonomousMove", "end_and_after_game_autonomous_move"
        ),
        "end_teleop_move": pick("endAndAfterGameTeleopMove", "end_and_after_game_teleop_move"),
        "end_comments": pick("endAndAfterGameComments", "end_and_after_game_comments"),
    }


def load_records_by_event(scout_event_id: str) -> OrderedDict[int, dict[str, Any]]:
    with get_db_connection() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            cursor.execute(
                """
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema = 'public' AND table_name = 'team_match_record'
                """
            )
            columns = {row["column_name"] for row in cursor.fetchall()}
            mapping = resolve_columns(columns)

            team_col = quote_ident(mapping["team"])
            event_col = quote_ident(mapping["event"])
            match_number_col = quote_ident(mapping["match_number"])
            alliance_col = quote_ident(mapping["alliance"])

            query = f"""
            SELECT r.*, t.number AS joined_team_number
            FROM team_match_record r
            LEFT JOIN team t ON t.number = r.{team_col}
            WHERE r.{event_col} = %s
            ORDER BY r.{match_number_col} ASC, r.{alliance_col} ASC, COALESCE(t.number, r.{team_col}) ASC
            """

            cursor.execute(query, (scout_event_id,))
            rows = cursor.fetchall()

    matches: OrderedDict[int, dict[str, Any]] = OrderedDict()

    for row in rows:
        autonomous_payload = (
            get_nested_dict(row, [mapping["autonomous_json"]])
            if mapping["autonomous_json"]
            else {}
        )
        teleop_payload = (
            get_nested_dict(row, [mapping["teleop_json"]])
            if mapping["teleop_json"]
            else {}
        )
        end_payload = (
            get_nested_dict(row, [mapping["end_json"]])
            if mapping["end_json"]
            else {}
        )

        normalized = {
            "id": get_value(row, [mapping["id"]]),
            "match_number": get_value(row, [mapping["match_number"]]),
            "match_type": get_value(row, [mapping["match_type"]]),
            "alliance": get_value(row, [mapping["alliance"]]),
            "team_number": get_value(row, ["joined_team_number", mapping["team"]]),
            "match_key": get_value(row, [mapping["match_key"]]) if mapping["match_key"] else None,
            "auto_shooter_type": (
                get_value(row, [mapping["auto_shooter_type"]])
                if mapping["auto_shooter_type"]
                else pick_from_nested(autonomous_payload, ["shooterType", "shooter_type"])
            ),
            "auto_shots_taken": (
                get_value(row, [mapping["auto_shots_taken"]])
                if mapping["auto_shots_taken"]
                else pick_from_nested(autonomous_payload, ["shotsTaken", "shots_taken"])
            ),
            "auto_shot_volumes": (
                get_value(row, [mapping["auto_shot_volumes"]])
                if mapping["auto_shot_volumes"]
                else pick_from_nested(autonomous_payload, ["shotVolumes", "shot_volumes"])
            ),
            "auto_subjective_accuracy": (
                get_value(row, [mapping["auto_subjective_accuracy"]])
                if mapping["auto_subjective_accuracy"]
                else pick_from_nested(
                    autonomous_payload,
                    ["subjectiveAccuracy", "subjective_accuracy"],
                )
            ),
            "teleop_fuel_count": (
                get_value(row, [mapping["teleop_fuel_count"]])
                if mapping["teleop_fuel_count"]
                else pick_from_nested(teleop_payload, ["fuelCount", "fuel_count"])
            ),
            "teleop_human_fuel_count": (
                get_value(row, [mapping["teleop_human_fuel_count"]])
                if mapping["teleop_human_fuel_count"]
                else pick_from_nested(
                    teleop_payload,
                    ["humanFuelCount", "human_fuel_count"],
                )
            ),
            "teleop_pass_bump": (
                get_value(row, [mapping["teleop_pass_bump"]])
                if mapping["teleop_pass_bump"]
                else pick_from_nested(teleop_payload, ["passBump", "pass_bump"])
            ),
            "teleop_pass_trench": (
                get_value(row, [mapping["teleop_pass_trench"]])
                if mapping["teleop_pass_trench"]
                else pick_from_nested(teleop_payload, ["passTrench", "pass_trench"])
            ),
            "teleop_fetch_ball_preference": (
                get_value(row, [mapping["teleop_fetch_ball_preference"]])
                if mapping["teleop_fetch_ball_preference"]
                else pick_from_nested(
                    teleop_payload,
                    ["fetchBallPreference", "fetch_ball_preference"],
                )
            ),
            "teleop_shots_taken": (
                get_value(row, [mapping["teleop_shots_taken"]])
                if mapping["teleop_shots_taken"]
                else pick_from_nested(teleop_payload, ["shotsTaken", "shots_taken"])
            ),
            "teleop_shot_volumes": (
                get_value(row, [mapping["teleop_shot_volumes"]])
                if mapping["teleop_shot_volumes"]
                else pick_from_nested(teleop_payload, ["shotVolumes", "shot_volumes"])
            ),
            "teleop_subjective_accuracy": (
                get_value(row, [mapping["teleop_subjective_accuracy"]])
                if mapping["teleop_subjective_accuracy"]
                else pick_from_nested(
                    teleop_payload,
                    ["subjectiveAccuracy", "subjective_accuracy"],
                )
            ),
            "end_tower_status": (
                get_value(row, [mapping["end_tower_status"]])
                if mapping["end_tower_status"]
                else pick_from_nested(end_payload, ["towerStatus", "tower_status"])
            ),
            "end_climbing_time": (
                get_value(row, [mapping["end_climbing_time"]])
                if mapping["end_climbing_time"]
                else pick_from_nested(end_payload, ["climbingTime", "climbing_time"])
            ),
            "end_ranking_point": (
                get_value(row, [mapping["end_ranking_point"]])
                if mapping["end_ranking_point"]
                else pick_from_nested(end_payload, ["rankingPoint", "ranking_point"])
            ),
            "end_coop_point": (
                get_value(row, [mapping["end_coop_point"]])
                if mapping["end_coop_point"]
                else pick_from_nested(end_payload, ["coopPoint", "coop_point"])
            ),
            "end_autonomous_move": (
                get_value(row, [mapping["end_autonomous_move"]])
                if mapping["end_autonomous_move"]
                else pick_from_nested(
                    end_payload,
                    ["autonomousMove", "autonomous_move"],
                )
            ),
            "end_teleop_move": (
                get_value(row, [mapping["end_teleop_move"]])
                if mapping["end_teleop_move"]
                else pick_from_nested(end_payload, ["teleopMove", "teleop_move"])
            ),
            "end_comments": (
                get_value(row, [mapping["end_comments"]])
                if mapping["end_comments"]
                else pick_from_nested(end_payload, ["comments"])
            ),
        }

        match_number = normalized["match_number"]
        if match_number is None:
            continue

        if match_number not in matches:
            matches[match_number] = {
                "match_number": match_number,
                "match_type": normalized["match_type"],
                "red": OrderedDict(),
                "blue": OrderedDict(),
            }

        alliance_value = str(normalized["alliance"] or "").strip().lower()
        bucket = "red" if "red" in alliance_value else "blue"
        team_key = normalized["team_number"]
        if team_key is None:
            team_key = f"unknown-{normalized['id']}"

        if team_key not in matches[match_number][bucket]:
            matches[match_number][bucket][team_key] = normalized

    return matches


def build_matrix(match_data: dict[str, Any]) -> tuple[list[dict[str, Any]], list[list[str]]]:
    red_teams = list(match_data["red"].values())
    blue_teams = list(match_data["blue"].values())

    columns: list[dict[str, Any]] = []
    for team in red_teams:
        columns.append(
            {
                "label": f"Red Alliance {team['team_number']}",
                "record": team,
            }
        )
    for team in blue_teams:
        columns.append(
            {
                "label": f"Blue Alliance {team['team_number']}",
                "record": team,
            }
        )

    while len(columns) < 6:
        if len(columns) < 3:
            columns.append({"label": "Red Alliance", "record": None})
        else:
            columns.append({"label": "Blue Alliance", "record": None})

    rows: list[list[str]] = []
    for category, field_name, key in FIELD_ROWS:
        row = [category, field_name]
        for column in columns:
            record = column["record"]
            row.append(format_value(record.get(key)) if record else "")
        rows.append(row)

    return columns, rows


def create_excel(matches: OrderedDict[int, dict[str, Any]]) -> bytes:
    workbook = Workbook()
    first_sheet = True

    for match_number, match_data in matches.items():
        if first_sheet:
            sheet = workbook.active
            first_sheet = False
        else:
            sheet = workbook.create_sheet()

        sheet.title = f"M{match_number}"[:31]
        columns, rows = build_matrix(match_data)

        sheet.cell(row=1, column=1, value="Info")
        sheet.cell(row=1, column=2, value=f"Match {match_number}")
        for index, column in enumerate(columns, start=3):
            sheet.cell(row=1, column=index, value=column["label"])

        for row_index, row_data in enumerate(rows, start=2):
            for col_index, value in enumerate(row_data, start=1):
                sheet.cell(row=row_index, column=col_index, value=value)

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def create_xml(matches: OrderedDict[int, dict[str, Any]]) -> bytes:
    root = Element("TeamMatchRecords")

    for match_number, match_data in matches.items():
        page = SubElement(root, "Page")
        page.set("matchNumber", str(match_number))
        page.set("matchType", format_value(match_data.get("match_type")))

        columns, rows = build_matrix(match_data)

        headers = SubElement(page, "Headers")
        SubElement(headers, "Header").text = "Category"
        SubElement(headers, "Header").text = "Field"
        for column in columns:
            SubElement(headers, "Header").text = column["label"]

        body = SubElement(page, "Rows")
        for row in rows:
            row_el = SubElement(body, "Row")
            for value in row:
                SubElement(row_el, "Cell").text = value

    return tostring(root, encoding="utf-8", xml_declaration=True)


@app.get("/")
def index():
    scout_event_id = (request.args.get("scout_event_id") or "").strip()
    selected_match_param = (request.args.get("match_number") or "").strip()

    matches: OrderedDict[int, dict[str, Any]] = OrderedDict()
    match_numbers: list[int] = []
    selected_match: int | None = None
    columns = []
    rows = []
    error = None

    if scout_event_id:
        try:
            matches = load_records_by_event(scout_event_id)
            match_numbers = list(matches.keys())
            if match_numbers:
                if selected_match_param.isdigit() and int(selected_match_param) in matches:
                    selected_match = int(selected_match_param)
                else:
                    selected_match = match_numbers[0]

                columns, rows = build_matrix(matches[selected_match])
        except Exception as exc:
            error = str(exc)

    return render_template(
        "index.html",
        scout_event_id=scout_event_id,
        match_numbers=match_numbers,
        selected_match=selected_match,
        columns=columns,
        rows=rows,
        error=error,
    )


@app.get("/export/excel")
def export_excel():
    scout_event_id = (request.args.get("scout_event_id") or "").strip()
    if not scout_event_id:
        return Response("Missing scout_event_id", status=400)

    matches = load_records_by_event(scout_event_id)
    if not matches:
        return Response("No match records found", status=404)

    content = create_excel(matches)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"team_match_record_{scout_event_id}_{timestamp}.xlsx"

    return Response(
        content,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/export/xml")
def export_xml():
    scout_event_id = (request.args.get("scout_event_id") or "").strip()
    if not scout_event_id:
        return Response("Missing scout_event_id", status=400)

    matches = load_records_by_event(scout_event_id)
    if not matches:
        return Response("No match records found", status=404)

    content = create_xml(matches)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"team_match_record_{scout_event_id}_{timestamp}.xml"

    return Response(
        content,
        mimetype="application/xml",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
