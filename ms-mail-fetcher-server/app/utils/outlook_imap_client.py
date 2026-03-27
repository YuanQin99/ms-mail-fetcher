import imaplib
import email
from email.header import decode_header
from email import utils as email_utils
import requests

# --- 常量配置 ---
IMAP_SERVER = 'outlook.live.com'
IMAP_PORT = 993
TOKEN_URL = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token'
INBOX_FOLDER_NAME = "INBOX"
JUNK_FOLDER_NAME = "Junk"


def _looks_like_html(content: str) -> bool:
    if not content:
        return False
    text = content.lstrip().lower()
    return (
        text.startswith("<!doctype html")
        or text.startswith("<html")
        or ("<body" in text and "</body>" in text)
    )


def decode_header_value(header_value):
    """辅助函数：解码邮件头中的中文字符等"""
    if header_value is None: return ""
    decoded_string = ""
    try:
        parts = decode_header(str(header_value))
        for part, charset in parts:
            if isinstance(part, bytes):
                try:
                    decoded_string += part.decode(charset if charset else 'utf-8', 'replace')
                except LookupError:
                    decoded_string += part.decode('utf-8', 'replace')
            else:
                decoded_string += str(part)
    except Exception:
        return str(header_value)
    return decoded_string


# =======================================================
# 内部隐藏辅助函数：临时换取 access_token (对外不可见/不关心)
# =======================================================
def _get_temp_access_token(client_id, refresh_token):
    """内部使用的工具，默默用 refresh_token 换一个临时的 access_token 来干活"""
    try:
        response = requests.post(TOKEN_URL, data={
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
            'scope': 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access'
        })
        response.raise_for_status()
        return response.json().get('access_token')
    except Exception as e:
        print(f"内部获取临时 access_token 失败: {e}")
        return None


# =======================================================
# 专属工具：手动刷新并获取新的 Refresh Token (你自己定期调)
# =======================================================
def refresh_oauth_token_manually(client_id, current_refresh_token):
    """
    专门用来刷新 Token 的工具。
    返回包含新 refresh_token 的字典，拿到后你自己保存到本地或数据库。
    """
    result = {
        "success": False,
        "new_refresh_token": "",
        "new_access_token": "",
        "error_msg": ""
    }
    try:
        response = requests.post(TOKEN_URL, data={
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': current_refresh_token,
            'scope': 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access'
        })
        response.raise_for_status()
        token_data = response.json()

        result["new_access_token"] = token_data.get('access_token', "")
        result["new_refresh_token"] = token_data.get('refresh_token', "")

        if result["new_access_token"] and result["new_refresh_token"]:
            result["success"] = True
        else:
            result["error_msg"] = "微软接口未返回完整的 token 数据"

    except Exception as e:
        result["error_msg"] = f"刷新 Token 失败: {e}"

    return result


# =======================================================
# 核心业务 1：分页获取邮件列表 (参数完全还原为你的要求)
# =======================================================
def get_emails_by_folder_paginated(email_address, refresh_token, client_id, target_folder=INBOX_FOLDER_NAME,
                                   page_number=0, emails_per_page=10):
    """
    分页获取 Outlook 指定文件夹的邮件列表。
    只返回干净的邮件数据，不包含任何 token 刷新的杂质。
    """
    result = {
        "success": False,
        "error_msg": "",
        "total_emails": 0,
        "emails": []
    }

    # 1. 内部默默拿临时钥匙
    access_token = _get_temp_access_token(client_id, refresh_token)
    if not access_token:
        result["error_msg"] = "未能获取到有效的 access_token，请检查你的 refresh_token 是否有效"
        return result

    imap_conn = None
    try:
        imap_conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        auth_string = f"user={email_address}\1auth=Bearer {access_token}\1\1"
        typ, _ = imap_conn.authenticate('XOAUTH2', lambda x: auth_string.encode('utf-8'))

        if typ != 'OK':
            result["error_msg"] = "IMAP 认证失败，请确认凭证有效"
            return result

        typ, _ = imap_conn.select(target_folder, readonly=True)
        if typ != 'OK':
            result["error_msg"] = f"选择文件夹 '{target_folder}' 失败"
            return result

        typ, uid_data = imap_conn.uid('search', None, "ALL")
        if typ != 'OK' or not uid_data[0]:
            result["success"] = True
            return result

        uids = uid_data[0].split()
        result["total_emails"] = len(uids)
        uids.reverse()

        start_index = page_number * emails_per_page
        end_index = start_index + emails_per_page
        page_uids = uids[start_index:end_index]

        emails_list = []
        for uid_bytes in page_uids:
            uid_str = uid_bytes.decode('utf-8', 'replace')
            typ, msg_data = imap_conn.uid('fetch', uid_bytes, '(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE FROM)])')

            subject_str = "(No Subject)"
            formatted_date_str = "(No Date)"
            from_name = "(Unknown)"
            from_email = ""

            if typ == 'OK' and msg_data and msg_data[0] is not None:
                header_content_bytes = None
                if isinstance(msg_data[0], tuple) and len(msg_data[0]) == 2:
                    header_content_bytes = msg_data[0][1]
                elif isinstance(msg_data, list) and len(msg_data) > 1:
                    header_content_bytes = msg_data[1]

                if header_content_bytes:
                    header_message = email.message_from_bytes(header_content_bytes)
                    subject_str = decode_header_value(header_message.get('Subject', '(No Subject)'))
                    from_str = decode_header_value(header_message.get('From', '(Unknown Sender)'))

                    if '<' in from_str and '>' in from_str:
                        from_name = from_str.split('<')[0].strip().strip('"')
                        from_email = from_str.split('<')[1].split('>')[0].strip()
                    else:
                        from_email = from_str.strip()
                        if '@' in from_email:
                            from_name = from_email.split('@')[0]

                    date_header_str = header_message.get('Date')
                    if date_header_str:
                        try:
                            dt_obj = email_utils.parsedate_to_datetime(date_header_str)
                            if dt_obj: formatted_date_str = dt_obj.strftime('%Y-%m-%d %H:%M:%S')
                        except Exception:
                            pass

            emails_list.append({
                'uid': uid_str,
                'subject': subject_str,
                'from_name': from_name,
                'from_email': from_email,
                'date': formatted_date_str,
                'folder': target_folder
            })

        result["emails"] = emails_list
        result["success"] = True
        return result

    except Exception as e:
        result["error_msg"] = f"发生异常: {e}"
        return result
    finally:
        if imap_conn:
            try:
                imap_conn.close()
                imap_conn.logout()
            except:
                pass


# =======================================================
# 核心业务 2：获取特定邮件详情 (参数完全还原为你的要求)
# =======================================================
def get_email_detail_by_uid(email_address, refresh_token, client_id, target_uid, target_folder=INBOX_FOLDER_NAME):
    """
    根据 UID 获取特定邮件的完整内容。
    只返回纯净的详情字典，不含 token 更新逻辑。
    """
    result = {
        "success": False,
        "error_msg": "",
        "detail": {
            "subject": "",
            "from": "",
            "to": "",
            "date": "",
            "body_text": "",
            "body_html": ""
        }
    }

    # 1. 内部默默拿临时钥匙
    access_token = _get_temp_access_token(client_id, refresh_token)
    if not access_token:
        result["error_msg"] = "未能获取到有效的 access_token"
        return result

    imap_conn = None
    try:
        imap_conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        auth_string = f"user={email_address}\1auth=Bearer {access_token}\1\1"
        typ, _ = imap_conn.authenticate('XOAUTH2', lambda x: auth_string.encode('utf-8'))

        if typ != 'OK':
            result["error_msg"] = "IMAP 认证失败"
            return result

        typ, _ = imap_conn.select(target_folder, readonly=True)
        if typ != 'OK':
            result["error_msg"] = f"选择文件夹 '{target_folder}' 失败"
            return result

        uid_bytes = target_uid.encode('utf-8') if isinstance(target_uid, str) else target_uid
        typ, msg_data = imap_conn.uid('fetch', uid_bytes, '(RFC822)')

        if typ != 'OK' or not msg_data or msg_data[0] is None:
            result["error_msg"] = f"未在 {target_folder} 找到 UID 为 {target_uid} 的邮件"
            return result

        raw_email_bytes = None
        if isinstance(msg_data[0], tuple) and len(msg_data[0]) == 2:
            raw_email_bytes = msg_data[0][1]
        elif isinstance(msg_data, list):
            for item in msg_data:
                if isinstance(item, tuple) and len(item) == 2:
                    raw_email_bytes = item[1];
                    break

        if not raw_email_bytes:
            result["error_msg"] = "解析邮件数据结构失败"
            return result

        email_message = email.message_from_bytes(raw_email_bytes)

        result["detail"]["subject"] = decode_header_value(email_message.get('Subject', '(No Subject)'))
        result["detail"]["from"] = decode_header_value(email_message.get('From', '(Unknown Sender)'))
        result["detail"]["to"] = decode_header_value(email_message.get('To', '(Unknown Recipient)'))
        result["detail"]["date"] = email_message.get('Date', '(Unknown Date)')

        body_text = ""
        body_html = ""

        if email_message.is_multipart():
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if "attachment" not in content_disposition:
                    try:
                        charset = part.get_content_charset() or 'utf-8'
                        payload = part.get_payload(decode=True)
                        if payload:
                            decoded_str = payload.decode(charset, errors='replace')
                            if content_type == "text/plain":
                                body_text += decoded_str
                            elif content_type == "text/html":
                                body_html += decoded_str
                    except Exception:
                        pass
        else:
            try:
                charset = email_message.get_content_charset() or 'utf-8'
                payload = email_message.get_payload(decode=True)
                if payload:
                    decoded_str = payload.decode(charset, errors='replace')
                    content_type = email_message.get_content_type()
                    if content_type == "text/html":
                        body_html = decoded_str
                    else:
                        body_text = decoded_str
            except Exception:
                pass

        body_text = body_text.strip()
        body_html = body_html.strip()

        # 兜底：部分邮件服务会把 HTML 正文错误标为 text/plain。
        # 遇到这种情况，把看起来像 HTML 的正文转到 body_html，供前端渲染。
        if not body_html and _looks_like_html(body_text):
            body_html = body_text
            body_text = ""

        result["detail"]["body_text"] = body_text
        result["detail"]["body_html"] = body_html
        result["success"] = True
        return result

    except Exception as e:
        result["error_msg"] = f"解析异常: {e}"
        return result
    finally:
        if imap_conn:
            try:
                imap_conn.close()
                imap_conn.logout()
            except:
                pass


# --- 用法示例 ---
if __name__ == "__main__":
    TEST_CLIENT_ID = '9e5f94bc-e8a4-4e73-b8be-63364c29d753'
    TEST_EMAIL = 'AdrianJones7591@outlook.com'
    TEST_REFRESH_TOKEN = 'M.C536_SN1.0.U.-CsWcHph3Kdy2aP9mEIHeES4HDzQj7Fi1WYFq!PZQ6dR!nKznaRGG2!V6SuZFfyIddv9U7ohjp9X4iUu2G978J84tQXM4KFduPV!lGvVClMUehH44yVN*hrIEZl5PqEnKaMWvkvVXZXk8dgCEeSJXLgfgnMGRBmtLg9OvEewOTKFE6l8mI38SSvaIbLD!Z7fnMVzecZJHVFeO1qUekXwUEB7iHdRNPtmI3*CoRQ46OtYhWNDD9j*4*w7Gnpjwvao*55q!ekBGAK7CjdaouLWvXGBvl3MAEhy8gX687P1KSqNRAtnMbPVY0cHSeUFMDhlCGSZ!U!MqG6WuKjpVo4RlUvBCzA29MfS!eFEPUWcrYlPbGtVTDGefRZUrV9lGMhesX8AJeSmb4hFPTN7RvpjcjfmwgOcuPycbJwfFWDLOQrZBaZ7g26h8ruVOmHORDlbDWA$$'

    # 场景 1：日常读取邮件，不再烦恼 token 更新的返回值，干净清爽！
    print(">>> 测试：静默获取邮件列表")
    list_res = get_emails_by_folder_paginated(
        TEST_EMAIL, TEST_REFRESH_TOKEN, TEST_CLIENT_ID,
        target_folder=INBOX_FOLDER_NAME, page_number=0, emails_per_page=3
    )
    if list_res["success"]:
        print(f"成功获取 {len(list_res['emails'])} 封邮件。返回字典里再也没有 token 相关的杂乱字段了！")
    else:
        print(f"获取失败: {list_res['error_msg']}")

    # 场景 2：假设过了一个月，你想手动刷新并持久化 token 了
    print("\n>>> 测试：手动调用专职刷新工具")
    token_res = refresh_oauth_token_manually(TEST_CLIENT_ID, TEST_REFRESH_TOKEN)
    if token_res["success"]:
        print(f"刷新成功！新的 Refresh Token: {token_res['new_refresh_token'][:20]}...")
        # TODO: 这里写你保存到本地 txt 或数据库的代码
    else:
        print(f"刷新失败: {token_res['error_msg']}")
