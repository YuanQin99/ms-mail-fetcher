from datetime import datetime

from fastapi import HTTPException
from sqlalchemy import and_, or_
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from app.crud.account_types import ensure_account_type_exists, normalize_code
from app.models.models import Account, ArchivedAccount
from app.schemas.schemas import AccountCreate, AccountOut, AccountUpdate, ImportResult, PaginatedAccounts
from app.utils.outlook_imap_client import refresh_oauth_token_manually


def to_account_out(account: Account | ArchivedAccount) -> AccountOut:
    days_since_refresh = max((datetime.utcnow() - account.last_refresh_time).days, 0)
    return AccountOut(
        id=account.id,
        email=account.email,
        password=account.password,
        client_id=account.client_id,
        refresh_token=account.refresh_token,
        last_refresh_time=account.last_refresh_time,
        account_type=account.account_type,
        remark=account.remark,
        is_active=account.is_active,
        days_since_refresh=days_since_refresh,
        source_account_id=getattr(account, 'source_account_id', None),
    )


def get_account_model(is_active: bool):
    return Account if is_active else ArchivedAccount


def get_account_or_404(db: Session, account_id: int, is_active: bool) -> Account | ArchivedAccount:
    model = get_account_model(is_active)
    account = db.query(model).filter(model.id == account_id).first()
    if not account:
        raise HTTPException(status_code=404, detail='账号不存在')
    return account


def parse_import_text(text: str):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    parsed = []
    errors = []

    for idx, line in enumerate(lines, start=1):
        parts = [part.strip() for part in line.split('----')]
        if len(parts) != 4:
            errors.append(f'第 {idx} 行格式错误: {line}')
            continue
        email, password, client_id, refresh_token = parts
        if not all(parts):
            errors.append(f'第 {idx} 行存在空字段: {line}')
            continue
        parsed.append(
            {
                'email': email,
                'password': password,
                'client_id': client_id,
                'refresh_token': refresh_token,
            }
        )

    return parsed, errors


def list_accounts(
    db: Session,
    is_active: bool,
    search: str | None,
    account_type: str | None,
    page: int,
    page_size: int,
) -> PaginatedAccounts:
    model = get_account_model(is_active)
    filters = [model.is_active == is_active]

    if search:
        keyword = f"%{search.strip()}%"
        filters.append(or_(model.email.like(keyword), model.remark.like(keyword)))

    if account_type:
        filters.append(model.account_type == account_type)

    query = db.query(model).filter(and_(*filters)).order_by(model.id.desc())
    total = query.count()
    items = query.offset((page - 1) * page_size).limit(page_size).all()

    return PaginatedAccounts(
        items=[to_account_out(item) for item in items],
        total=total,
        page=page,
        page_size=page_size,
    )


def create_account(db: Session, payload: AccountCreate) -> AccountOut:
    normalized_type = normalize_code(payload.account_type) if payload.account_type else None
    ensure_account_type_exists(db, normalized_type)

    account = Account(
        email=payload.email,
        password=payload.password,
        client_id=payload.client_id,
        refresh_token=payload.refresh_token,
        account_type=normalized_type,
        remark=payload.remark,
        is_active=True,
    )
    db.add(account)

    try:
        db.commit()
    except IntegrityError:
        db.rollback()
        raise HTTPException(status_code=409, detail='邮箱已存在')

    db.refresh(account)
    return to_account_out(account)


def update_account(db: Session, account_id: int, payload: AccountUpdate) -> AccountOut:
    target_is_active = payload.is_active if payload.is_active is not None else True
    account = get_account_or_404(db, account_id, target_is_active)

    updates = payload.model_dump(exclude_unset=True)
    updates.pop('is_active', None)
    if 'account_type' in updates:
        account_type = updates.get('account_type')
        normalized_type = normalize_code(account_type) if account_type else None
        ensure_account_type_exists(db, normalized_type)
        updates['account_type'] = normalized_type

    for key, value in updates.items():
        setattr(account, key, value)

    db.add(account)
    db.commit()
    db.refresh(account)
    return to_account_out(account)


def import_accounts(
    db: Session,
    content: str,
    is_active: bool,
    account_type: str | None,
) -> ImportResult:
    normalized_type = normalize_code(account_type) if account_type else None
    ensure_account_type_exists(db, normalized_type)
    model = get_account_model(is_active)

    parsed, errors = parse_import_text(content)
    inserted = 0
    skipped = 0

    emails = [item['email'] for item in parsed]
    existing_emails = {
        email
        for (email,) in db.query(model.email)
        .filter(model.email.in_(emails))
        .all()
    }

    for item in parsed:
        if item['email'] in existing_emails:
            skipped += 1
            continue

        account = model(
            email=item['email'],
            password=item['password'],
            client_id=item['client_id'],
            refresh_token=item['refresh_token'],
            account_type=normalized_type,
            is_active=is_active,
        )
        db.add(account)
        inserted += 1

    db.commit()
    return ImportResult(inserted=inserted, skipped=skipped, errors=errors)


def export_accounts_text(
    db: Session,
    is_active: bool,
    search: str | None,
    account_type: str | None,
    account_ids: list[int] | None = None,
) -> str:
    model = get_account_model(is_active)
    filters = [model.is_active == is_active]
    if search:
        keyword = f"%{search.strip()}%"
        filters.append(or_(model.email.like(keyword), model.remark.like(keyword)))
    if account_type:
        filters.append(model.account_type == account_type)
    if account_ids:
        filters.append(model.id.in_(account_ids))

    items = db.query(model).filter(and_(*filters)).order_by(model.id.desc()).all()
    return '\n'.join(
        [f"{item.email}----{item.password}----{item.client_id}----{item.refresh_token}" for item in items]
    )


def archive_account(db: Session, account_id: int) -> dict:
    account = get_account_or_404(db, account_id, True)

    archived = ArchivedAccount(
        email=account.email,
        password=account.password,
        client_id=account.client_id,
        refresh_token=account.refresh_token,
        last_refresh_time=account.last_refresh_time,
        account_type=account.account_type,
        remark=account.remark,
        source_account_id=account.id,
        is_active=False,
    )
    db.add(archived)
    db.delete(account)
    db.commit()
    return {'message': '已归档'}


def archive_all_active_accounts(db: Session) -> dict:
    active_accounts = db.query(Account).filter(Account.is_active == True).order_by(Account.id.asc()).all()

    for account in active_accounts:
        db.add(
            ArchivedAccount(
                email=account.email,
                password=account.password,
                client_id=account.client_id,
                refresh_token=account.refresh_token,
                last_refresh_time=account.last_refresh_time,
                account_type=account.account_type,
                remark=account.remark,
                source_account_id=account.id,
                is_active=False,
            )
        )
        db.delete(account)

    db.commit()
    return {'message': '活跃账号已全部归档', 'count': len(active_accounts)}


def delete_account(db: Session, account_id: int, is_active: bool) -> dict:
    account = get_account_or_404(db, account_id, is_active)

    db.delete(account)
    db.commit()
    return {'message': '删除成功'}


def refresh_all_account_tokens(
    db: Session,
    is_active: bool,
    search: str | None,
    account_type: str | None,
) -> dict:
    model = get_account_model(is_active)
    filters = [model.is_active == is_active]
    if search:
        keyword = f"%{search.strip()}%"
        filters.append(or_(model.email.like(keyword), model.remark.like(keyword)))
    if account_type:
        filters.append(model.account_type == account_type)

    accounts = db.query(model).filter(and_(*filters)).order_by(model.id.desc()).all()

    success_count = 0
    failed_count = 0
    errors: list[str] = []

    for account in accounts:
        refresh_result = refresh_oauth_token_manually(account.client_id, account.refresh_token)
        if not refresh_result.get('success'):
            failed_count += 1
            errors.append(f"{account.email}: {refresh_result.get('error_msg', '刷新失败')}")
            continue

        new_refresh_token = refresh_result.get('new_refresh_token')
        if not new_refresh_token:
            failed_count += 1
            errors.append(f'{account.email}: 刷新成功但未返回 refresh_token')
            continue

        account.refresh_token = new_refresh_token
        account.last_refresh_time = datetime.utcnow()
        db.add(account)
        success_count += 1

    db.commit()

    return {
        'total': len(accounts),
        'success': success_count,
        'failed': failed_count,
        'errors': errors,
    }


def migrate_inactive_accounts_to_archive_table(db: Session) -> int:
    inactive_accounts = db.query(Account).filter(Account.is_active == False).order_by(Account.id.asc()).all()
    moved = 0

    for account in inactive_accounts:
        exists = db.query(ArchivedAccount.id).filter(ArchivedAccount.email == account.email).first()
        if not exists:
            db.add(
                ArchivedAccount(
                    email=account.email,
                    password=account.password,
                    client_id=account.client_id,
                    refresh_token=account.refresh_token,
                    last_refresh_time=account.last_refresh_time,
                    account_type=account.account_type,
                    remark=account.remark,
                    source_account_id=account.id,
                    is_active=False,
                )
            )
            moved += 1
        db.delete(account)

    if inactive_accounts:
        db.commit()

    return moved
