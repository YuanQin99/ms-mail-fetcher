from fastapi import APIRouter, Depends, Query
from sqlalchemy.orm import Session

from app.crud.mail import get_mail_detail, list_mails
from app.db.database import get_db
from app.schemas.schemas import MailDetailResponse, MailListResponse

router = APIRouter(prefix='/api/accounts', tags=['mail'])


@router.get('/{account_id}/mail/{folder}', response_model=MailListResponse)
def list_mails_route(
    account_id: int,
    folder: str,
    is_active: bool = Query(True),
    page: int = Query(1, ge=1),
    page_size: int = Query(10, ge=1, le=100),
    db: Session = Depends(get_db),
):
    return list_mails(db=db, account_id=account_id, folder=folder, page=page, page_size=page_size, is_active=is_active)


@router.get('/{account_id}/mail/{folder}/{message_id}', response_model=MailDetailResponse)
def get_mail_detail_route(
    account_id: int,
    folder: str,
    message_id: str,
    is_active: bool = Query(True),
    db: Session = Depends(get_db),
):
    return get_mail_detail(db=db, account_id=account_id, folder=folder, message_id=message_id, is_active=is_active)
