"""Account-bound survey storage. No outbound services or answer logging."""
from collections import Counter
from datetime import datetime, timezone, timedelta
import secrets
import time

from utils.db import get_db_connection

SURVEY_ID = "service-ux-v1"
SURVEY_TITLE = "Оценка удобства и визуального оформления сервиса"
SNOOZE_SECONDS = 24 * 60 * 60
QUESTIONS = [
    {"id": "q1", "label": "Насколько вам удобно пользоваться сервисом?", "low": "совсем неудобно", "high": "очень удобно"},
    {"id": "q2", "label": "Насколько быстро вы находите нужную функцию?", "low": "очень долго", "high": "очень быстро"},
    {"id": "q3", "label": "Насколько понятны названия разделов, кнопок и действий?", "low": "совсем непонятны", "high": "полностью понятны"},
    {"id": "q4", "label": "Насколько вам нравится цветовая гамма сервиса?", "low": "совсем не нравится", "high": "очень нравится"},
    {"id": "q5", "label": "Насколько легко читать тексты и различать элементы интерфейса?", "low": "очень трудно", "high": "очень легко"},
    {"id": "q6", "label": "Как вы оцениваете общее визуальное оформление сервиса?", "low": "очень плохо", "high": "отлично"},
]
QUESTION_IDS = tuple(q["id"] for q in QUESTIONS)


def _now():
    return time.time()


def _iso(timestamp):
    return datetime.fromtimestamp(timestamp, timezone.utc).isoformat().replace("+00:00", "Z") if timestamp is not None else None


def _metadata():
    return {"id": SURVEY_ID, "title": SURVEY_TITLE, "questions": QUESTIONS}


def _ensure(conn, account_key):
    conn.execute("INSERT OR IGNORE INTO ux_survey_state(account_key,survey_id) VALUES (?,?)", (account_key, SURVEY_ID))


def _state(conn, account_key):
    row = conn.execute("SELECT * FROM ux_survey_state WHERE account_key=? AND survey_id=?", (account_key, SURVEY_ID)).fetchone()
    response = conn.execute("SELECT submitted_at FROM ux_survey_responses WHERE account_key=? AND survey_id=?", (account_key, SURVEY_ID)).fetchone()
    completed = response is not None
    ratings = {key: row[f"draft_{key}"] if row and not completed else None for key in QUESTION_IDS}
    comment = row["draft_comment"] if row and not completed else ""
    next_at = row["next_invite_at"] if row and not completed else None
    return {
        "survey": _metadata(), "completed": completed,
        "submitted_at": _iso(response["submitted_at"]) if completed else None,
        "draft": {"ratings": ratings, "comment": comment},
        "draft_version": row["draft_version"] if row else 0,
        "has_draft": any(value is not None for value in ratings.values()) or bool(comment),
        "eligible": not completed and (next_at is None or next_at <= _now()),
        "next_invite_at": _iso(next_at),
    }


def state(account_key):
    with get_db_connection() as conn:
        conn.execute("BEGIN")
        return _state(conn, account_key)


def claim_invitation(account_key):
    with get_db_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        _ensure(conn, account_key)
        current = _state(conn, account_key)
        if not current["eligible"]:
            return {"show": False, "token": None, "state": current}
        token = secrets.token_urlsafe(24)
        conn.execute("UPDATE ux_survey_state SET next_invite_at=?,invitation_token=? WHERE account_key=? AND survey_id=?",
                     (_now() + SNOOZE_SECONDS, token, account_key, SURVEY_ID))
        return {"show": True, "token": token, "state": _state(conn, account_key)}


class SurveyConflict(Exception):
    def __init__(self, code, current):
        super().__init__(code)
        self.code = code
        self.state = current


def invitation_shown(account_key, token):
    with get_db_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        now = _now()
        result = conn.execute("UPDATE ux_survey_state SET first_shown_at=COALESCE(first_shown_at,?),last_shown_at=? "
                              "WHERE account_key=? AND survey_id=? AND invitation_token=?",
                              (now, now, account_key, SURVEY_ID, token))
        if not result.rowcount:
            raise SurveyConflict("invitation_expired", _state(conn, account_key))
        return {"ok": True}


def defer(account_key):
    with get_db_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        _ensure(conn, account_key)
        if not _state(conn, account_key)["completed"]:
            conn.execute("UPDATE ux_survey_state SET next_invite_at=? WHERE account_key=? AND survey_id=?",
                         (_now() + SNOOZE_SECONDS, account_key, SURVEY_ID))
        return _state(conn, account_key)


def _validate(ratings, comment, *, complete):
    if not isinstance(comment, str) or len(comment) > 2000 or "\x00" in comment:
        raise ValueError("Invalid survey comment")
    if set(ratings) - set(QUESTION_IDS):
        raise ValueError("Unknown survey question")
    for key in QUESTION_IDS:
        value = ratings.get(key)
        if value is None and not complete:
            continue
        if type(value) is not int or not 0 <= value <= 5:
            raise ValueError("Invalid survey rating")


def save_draft(account_key, ratings, comment, version):
    _validate(ratings, comment, complete=False)
    with get_db_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        _ensure(conn, account_key)
        current = _state(conn, account_key)
        if current["completed"]:
            raise SurveyConflict("already_submitted", current)
        if current["draft_version"] != version:
            raise SurveyConflict("draft_conflict", current)
        conn.execute("UPDATE ux_survey_state SET draft_q1=?,draft_q2=?,draft_q3=?,draft_q4=?,draft_q5=?,draft_q6=?,"
                     "draft_comment=?,draft_version=draft_version+1,draft_updated_at=? WHERE account_key=? AND survey_id=?",
                     (*[ratings.get(key) for key in QUESTION_IDS], comment, _now(), account_key, SURVEY_ID))
        return _state(conn, account_key)


def submit(account_key, ratings, comment):
    _validate(ratings, comment, complete=True)
    with get_db_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        _ensure(conn, account_key)
        previous = conn.execute("SELECT submitted_at FROM ux_survey_responses WHERE account_key=? AND survey_id=?",
                                (account_key, SURVEY_ID)).fetchone()
        if previous:
            return {"ok": True, "completed": True, "already_submitted": True, "submitted_at": _iso(previous["submitted_at"])}
        now = _now()
        conn.execute("INSERT INTO ux_survey_responses(account_key,survey_id,q1,q2,q3,q4,q5,q6,comment,submitted_at) "
                     "VALUES (?,?,?,?,?,?,?,?,?,?)", (account_key, SURVEY_ID, *[ratings[key] for key in QUESTION_IDS], comment, now))
        conn.execute("UPDATE ux_survey_state SET draft_q1=NULL,draft_q2=NULL,draft_q3=NULL,draft_q4=NULL,draft_q5=NULL,"
                     "draft_q6=NULL,draft_comment='',draft_version=draft_version+1,draft_updated_at=NULL,next_invite_at=NULL "
                     "WHERE account_key=? AND survey_id=?", (account_key, SURVEY_ID))
        return {"ok": True, "completed": True, "already_submitted": False, "submitted_at": _iso(now)}


def analytics():
    with get_db_connection() as conn:
        # A single read transaction gives all figures the same snapshot.
        conn.execute("BEGIN")
        responses = conn.execute("SELECT r.*,s.first_shown_at FROM ux_survey_responses r JOIN ux_survey_state s "
                                 "USING(account_key,survey_id) WHERE r.survey_id=? ORDER BY submitted_at DESC,account_key",
                                 (SURVEY_ID,)).fetchall()
        invited = conn.execute("SELECT COUNT(*) FROM ux_survey_state WHERE survey_id=? AND first_shown_at IS NOT NULL", (SURVEY_ID,)).fetchone()[0]
    exposed_completed = sum(row["first_shown_at"] is not None for row in responses)
    questions = []
    for question in QUESTIONS:
        values = [row[question["id"]] for row in responses]
        distribution = Counter(values)
        questions.append({**question, "count": len(values), "average": round(sum(values) / len(values), 1) if values else None,
                          "distribution": [{"score": score, "count": distribution[score]} for score in range(6)]})
    moscow = timezone(timedelta(hours=3))
    daily = Counter(datetime.fromtimestamp(row["submitted_at"], moscow).date().isoformat() for row in responses)
    return {
        "survey": _metadata(), "generated_at": _iso(_now()), "completed_count": len(responses),
        "invited_count": invited, "completed_invited_count": exposed_completed,
        "completion_rate": round(100 * exposed_completed / invited, 1) if invited else None,
        "questions": questions, "daily": [{"date": date, "count": daily[date]} for date in sorted(daily)],
        "comments": [{"submitted_at": _iso(row["submitted_at"]), "text": row["comment"]} for row in responses if row["comment"].strip()],
    }
