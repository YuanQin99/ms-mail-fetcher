import re

from fastapi import HTTPException
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from app.models.models import Account, AccountType, ArchivedAccount
from app.schemas.schemas import AccountTypeCreate, AccountTypeOut, AccountTypeUpdate

HEX_COLOR_PATTERN = re.compile(r"^#[0-9a-fA-F]{6}$")

DEFAULT_ACCOUNT_TYPES = [
    {"code": "team", "label": "Team", "color": "#409EFF"},
    {"code": "member", "label": "Member", "color": "#67C23A"},
    {"code": "plus", "label": "Plus", "color": "#E6A23C"},
    {"code": "idle", "label": "Idle", "color": "#909399"},
]


def normalize_code(code: str) -> str:
    return code.strip().lower()


def validate_color(color: str) -> str:
    value = color.strip()
    if not HEX_COLOR_PATTERN.match(value):
        raise HTTPException(status_code=400, detail='颜色必须是 #RRGGBB 格式')
    return value


def ensure_default_account_types(db: Session) -> None:
    existing_codes = {item.code for item in db.query(AccountType).all()}
    changed = False

    for item in DEFAULT_ACCOUNT_TYPES:
        code = item['code']
        if code in existing_codes:
            continue
        db.add(AccountType(code=code, label=item['label'], color=item['color']))
        changed = True

    if changed:
        db.commit()


def list_account_types(db: Session) -> list[AccountTypeOut]:
    items = db.query(AccountType).order_by(AccountType.id.asc()).all()
    return [AccountTypeOut.model_validate(item) for item in items]


def ensure_account_type_exists(db: Session, code: str | None) -> None:
    if not code:
        return

    normalized = normalize_code(code)
    exists = db.query(AccountType.id).filter(AccountType.code == normalized).first()
    if not exists:
        raise HTTPException(status_code=400, detail=f'账号类型不存在: {normalized}')


def create_account_type(db: Session, payload: AccountTypeCreate) -> AccountTypeOut:
    code = normalize_code(payload.code)
    if not code:
        raise HTTPException(status_code=400, detail='类型编码不能为空')

    model = AccountType(
        code=code,
        label=payload.label.strip() or code,
        color=validate_color(payload.color),
    )
    db.add(model)

    try:
        db.commit()
    except IntegrityError:
        db.rollback()
        raise HTTPException(status_code=409, detail='类型编码已存在')

    db.refresh(model)
    return AccountTypeOut.model_validate(model)


def update_account_type(db: Session, account_type_id: int, payload: AccountTypeUpdate) -> AccountTypeOut:
    model = db.query(AccountType).filter(AccountType.id == account_type_id).first()
    if not model:
        raise HTTPException(status_code=404, detail='账号类型不存在')

    if payload.label is not None:
        model.label = payload.label.strip() or model.code

    if payload.color is not None:
        model.color = validate_color(payload.color)

    db.add(model)
    db.commit()
    db.refresh(model)
    return AccountTypeOut.model_validate(model)


def delete_account_type(db: Session, account_type_id: int) -> dict:
    model = db.query(AccountType).filter(AccountType.id == account_type_id).first()
    if not model:
        raise HTTPException(status_code=404, detail='账号类型不存在')

    db.query(Account).filter(Account.account_type == model.code).update({Account.account_type: None})
    db.query(ArchivedAccount).filter(ArchivedAccount.account_type == model.code).update({ArchivedAccount.account_type: None})
    db.delete(model)
    db.commit()
    return {'message': '账号类型已删除'}
