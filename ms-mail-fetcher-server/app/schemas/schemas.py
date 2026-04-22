from datetime import datetime

from pydantic import BaseModel, Field


class AccountBase(BaseModel):
    email: str
    password: str
    client_id: str
    refresh_token: str
    account_type: str | None = None
    remark: str | None = None
    is_active: bool = True


class AccountCreate(AccountBase):
    pass


class AccountUpdate(BaseModel):
    password: str | None = None
    client_id: str | None = None
    refresh_token: str | None = None
    account_type: str | None = None
    remark: str | None = None
    is_active: bool | None = None


class AccountOut(AccountBase):
    id: int
    last_refresh_time: datetime
    days_since_refresh: int = Field(default=0)
    source_account_id: int | None = None

    model_config = {"from_attributes": True}


class PaginatedAccounts(BaseModel):
    items: list[AccountOut]
    total: int
    page: int
    page_size: int


class ImportResult(BaseModel):
    inserted: int
    skipped: int
    errors: list[str]


class AccountTypeBase(BaseModel):
    code: str
    label: str
    color: str = "#409EFF"


class AccountTypeCreate(AccountTypeBase):
    pass


class AccountTypeUpdate(BaseModel):
    label: str | None = None
    color: str | None = None


class AccountTypeOut(AccountTypeBase):
    id: int

    model_config = {"from_attributes": True}


class MailItem(BaseModel):
    uid: str
    subject: str
    from_name: str
    from_email: str
    date: str
    folder: str


class MailListResponse(BaseModel):
    account_id: int
    email: str
    folder: str
    page: int
    page_size: int
    total: int
    items: list[MailItem]


class MailDetail(BaseModel):
    subject: str
    from_: str = Field(alias="from")
    to: str
    date: str
    body_text: str
    body_html: str


class MailDetailResponse(BaseModel):
    account_id: int
    email: str
    folder: str
    message_id: str
    detail: MailDetail


class UiPreferences(BaseModel):
    sidebar_collapsed: bool = False
    window_width: int = Field(default=1280, ge=1100)
    window_height: int = Field(default=860, ge=760)


class UiPreferencesUpdate(BaseModel):
    sidebar_collapsed: bool | None = None
    window_width: int | None = Field(default=None, ge=1100)
    window_height: int | None = Field(default=None, ge=760)
