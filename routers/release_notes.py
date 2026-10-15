"""Private release-note APIs; account identity never comes from the request body."""
from fastapi import APIRouter, Depends, HTTPException, Request, Response
from pydantic import BaseModel, ConfigDict

from services import release_notes
from services.auth.security import get_user_from_token


def private_response(response: Response):
    response.headers["Cache-Control"] = "private, no-store"


def require_user(request: Request):
    user = getattr(request.state, "user", None) or get_user_from_token(request.cookies.get("access_token"))
    if not user:
        raise HTTPException(401, "Необходимо войти в систему")
    return user


router = APIRouter(prefix="/api/releases", dependencies=[Depends(private_response)])


class Acknowledgement(BaseModel):
    model_config = ConfigDict(extra="forbid")


@router.get("/pending")
def pending(user=Depends(require_user)):
    return release_notes.pending(user)


@router.get("/history")
def history(user=Depends(require_user)):
    return release_notes.history(user)


@router.post("/{release_id}/read")
def mark_read(release_id: str, payload: Acknowledgement | None = None, user=Depends(require_user)):
    return release_notes.acknowledge(user, release_id)
