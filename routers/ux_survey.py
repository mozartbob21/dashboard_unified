"""Authenticated survey and administrator-only aggregate results."""
from typing import Annotated

from fastapi import APIRouter, Depends, HTTPException, Request, Response
from fastapi.responses import JSONResponse
from pydantic import BaseModel, ConfigDict, Field, StrictInt, StrictStr, field_validator

from core.web import templates
from services import ux_survey as survey
from services.auth.accounts import require_account_manager
from services.auth.security import get_user_from_token

def private_response(response: Response):
    response.headers["Cache-Control"] = "no-store"


router = APIRouter(dependencies=[Depends(private_response)])
Score = Annotated[StrictInt, Field(ge=0, le=5)]


def require_participant(request: Request):
    user = getattr(request.state, "user", None) or get_user_from_token(request.cookies.get("access_token"))
    if not user:
        raise HTTPException(401, "Необходимо войти в систему")
    if user.get("kc_sub"):
        return "keycloak:" + str(user["kc_sub"])
    if type(user.get("id")) is not int or not user.get("is_active", True):
        raise HTTPException(401, "Необходимо войти в систему")
    return f"local:{user['id']}"


class Payload(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Ratings(Payload):
    q1: Score
    q2: Score
    q3: Score
    q4: Score
    q5: Score
    q6: Score


class DraftRatings(Payload):
    q1: Score | None = None
    q2: Score | None = None
    q3: Score | None = None
    q4: Score | None = None
    q5: Score | None = None
    q6: Score | None = None


class CommentPayload(Payload):
    comment: StrictStr = Field(default="", max_length=2000)

    @field_validator("comment")
    @classmethod
    def no_null_character(cls, value):
        # SQLite length() stops at NUL; reject it instead of weakening the DB limit.
        if "\x00" in value:
            raise ValueError("Недопустимый символ в комментарии")
        return value


class Submission(CommentPayload):
    ratings: Ratings


class Draft(CommentPayload):
    ratings: DraftRatings
    draft_version: Annotated[StrictInt, Field(ge=0)]


class Shown(Payload):
    token: StrictStr = Field(min_length=1, max_length=100)


def _mutation(function, *args):
    try:
        return function(*args)
    except survey.SurveyConflict as error:
        return JSONResponse(status_code=409, content={"code": error.code, "state": error.state}, headers={"Cache-Control": "no-store"})


@router.get("/api/ux-survey/state")
def survey_state(account_key=Depends(require_participant)):
    return survey.state(account_key)


@router.post("/api/ux-survey/invitation/claim")
def invitation_claim(account_key=Depends(require_participant)):
    return survey.claim_invitation(account_key)


@router.post("/api/ux-survey/invitation/shown")
def invitation_shown(payload: Shown, account_key=Depends(require_participant)):
    return _mutation(survey.invitation_shown, account_key, payload.token)


@router.post("/api/ux-survey/later")
@router.post("/api/ux-survey/open")
def postpone(account_key=Depends(require_participant)):
    return survey.defer(account_key)


@router.put("/api/ux-survey/draft")
def draft_save(payload: Draft, account_key=Depends(require_participant)):
    return _mutation(survey.save_draft, account_key, payload.ratings.model_dump(), payload.comment, payload.draft_version)


@router.post("/api/ux-survey/submit")
def submit(payload: Submission, account_key=Depends(require_participant)):
    return survey.submit(account_key, payload.ratings.model_dump(), payload.comment)


@router.get("/api/ux-survey/admin", dependencies=[Depends(require_account_manager)])
def admin_analytics():
    return survey.analytics()


@router.get("/feedback/survey", dependencies=[Depends(require_participant)])
def survey_page(request: Request):
    return templates.TemplateResponse(request=request, name="ux_survey.html", context={"survey": survey._metadata()}, headers={"Cache-Control": "no-store"})


@router.get("/users/feedback", dependencies=[Depends(require_account_manager)])
def admin_page(request: Request):
    return templates.TemplateResponse(request=request, name="ux_survey_admin.html", context={}, headers={"Cache-Control": "no-store"})
