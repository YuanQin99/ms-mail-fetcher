from fastapi import HTTPException
from sqlalchemy.orm import Session

from app.crud.accounts import get_account_or_404
from app.schemas.schemas import MailDetailResponse, MailListResponse
from app.utils.outlook_imap_client import (
    INBOX_FOLDER_NAME,
    JUNK_FOLDER_NAME,
    get_email_detail_by_uid,
    get_emails_by_folder_paginated,
)


def resolve_folder(folder: str) -> str:
    folder_lower = folder.lower()
    if folder_lower == 'inbox':
        return INBOX_FOLDER_NAME
    if folder_lower == 'spam':
        return JUNK_FOLDER_NAME
    raise HTTPException(status_code=400, detail='folder 仅支持 inbox 或 spam')


def list_mails(db: Session, account_id: int, folder: str, page: int, page_size: int, is_active: bool) -> MailListResponse:
    account = get_account_or_404(db, account_id, is_active)
    target_folder = resolve_folder(folder)
    page_number = page - 1

    mail_result = get_emails_by_folder_paginated(
        email_address=account.email,
        refresh_token=account.refresh_token,
        client_id=account.client_id,
        target_folder=target_folder,
        page_number=page_number,
        emails_per_page=page_size,
    )

    if not mail_result.get('success'):
        raise HTTPException(status_code=400, detail=mail_result.get('error_msg', '读取邮件失败'))

    return MailListResponse(
        account_id=account.id,
        email=account.email,
        folder=folder.lower(),
        page=page,
        page_size=page_size,
        total=mail_result.get('total_emails', 0),
        items=mail_result.get('emails', []),
    )


def get_mail_detail(db: Session, account_id: int, folder: str, message_id: str, is_active: bool) -> MailDetailResponse:
    account = get_account_or_404(db, account_id, is_active)
    target_folder = resolve_folder(folder)

    detail_result = get_email_detail_by_uid(
        email_address=account.email,
        refresh_token=account.refresh_token,
        client_id=account.client_id,
        target_uid=message_id,
        target_folder=target_folder,
    )

    if not detail_result.get('success'):
        raise HTTPException(status_code=400, detail=detail_result.get('error_msg', '读取邮件详情失败'))

    return MailDetailResponse(
        account_id=account.id,
        email=account.email,
        folder=folder.lower(),
        message_id=message_id,
        detail=detail_result.get('detail', {}),
    )
