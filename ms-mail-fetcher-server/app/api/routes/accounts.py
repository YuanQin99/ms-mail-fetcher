from datetime import datetime
import re
from fastapi import APIRouter, Depends, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import PlainTextResponse
from sqlalchemy.orm import Session

from app.crud.accounts import (
    archive_account,
    archive_all_active_accounts,
    create_account,
    delete_account,
    export_accounts_text,
    import_accounts,
    list_accounts,
    refresh_all_account_tokens,
    update_account,
)
from app.db.database import get_db
from app.schemas.schemas import AccountCreate, AccountOut, AccountUpdate, ImportResult, PaginatedAccounts

router = APIRouter(prefix='/api/accounts', tags=['accounts'])


@router.get('', response_model=PaginatedAccounts)
def list_accounts_route(
    is_active: bool = Query(True),
    search: str | None = Query(None),
    type: str | None = Query(None),
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=200),
    db: Session = Depends(get_db),
):
    return list_accounts(
        db=db,
        is_active=is_active,
        search=search,
        account_type=type,
        page=page,
        page_size=page_size,
    )


@router.post('', response_model=AccountOut)
def create_account_route(payload: AccountCreate, db: Session = Depends(get_db)):
    return create_account(db=db, payload=payload)


@router.put('/archive-all')
def archive_all_active_accounts_route(db: Session = Depends(get_db)):
    return archive_all_active_accounts(db=db)


@router.post('/import', response_model=ImportResult)
async def import_accounts_route(
    text: str | None = Form(None),
    file: UploadFile | None = File(None),
    is_active: bool = Form(True),
    account_type: str | None = Form(None),
    db: Session = Depends(get_db),
):
    content = (text or '').strip()
    if file is not None:
        raw = await file.read()
        content = raw.decode('utf-8', errors='ignore').strip()

    if not content:
        raise HTTPException(status_code=400, detail='请提供导入文本或 txt 文件')

    return import_accounts(
        db=db,
        content=content,
        is_active=is_active,
        account_type=account_type,
    )


@router.get('/export')
def export_accounts_route(
    is_active: bool = Query(True),
    search: str | None = Query(None),
    type: str | None = Query(None),
    ids: str | None = Query(None),
    filename_prefix: str | None = Query(None),
    db: Session = Depends(get_db),
):
    account_ids = None
    if ids:
        account_ids = [int(item) for item in ids.split(',') if item.strip().isdigit()]

    output = export_accounts_text(
        db=db,
        is_active=is_active,
        search=search,
        account_type=type,
        account_ids=account_ids,
    )

    raw_prefix = (filename_prefix or 'accounts_export').strip() or 'accounts_export'
    safe_prefix = re.sub(r'[^A-Za-z0-9_-]+', '_', raw_prefix).strip('_') or 'accounts_export'
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
    file_name = f'{safe_prefix}_{timestamp}.txt'

    return PlainTextResponse(
        content=output,
        headers={'Content-Disposition': f'attachment; filename="{file_name}"'},
    )


@router.put('/{account_id}/archive')
def archive_account_route(account_id: int, db: Session = Depends(get_db)):
    return archive_account(db=db, account_id=account_id)


@router.put('/refresh-all-tokens')
def refresh_all_account_tokens_route(
    is_active: bool = Query(True),
    search: str | None = Query(None),
    type: str | None = Query(None),
    db: Session = Depends(get_db),
):
    return refresh_all_account_tokens(
        db=db,
        is_active=is_active,
        search=search,
        account_type=type,
    )


@router.put('/{account_id}', response_model=AccountOut)
def update_account_route(
    account_id: int,
    payload: AccountUpdate,
    is_active: bool = Query(True),
    db: Session = Depends(get_db),
):
    payload = payload.model_copy(update={'is_active': is_active})
    return update_account(db=db, account_id=account_id, payload=payload)


@router.delete('/{account_id}')
def delete_account_route(
    account_id: int,
    is_active: bool = Query(True),
    db: Session = Depends(get_db),
):
    return delete_account(db=db, account_id=account_id, is_active=is_active)
