# import time
import json
import random
import requests
from concurrent.futures import ThreadPoolExecutor
from util import multi_accounts_task, GracefulKiller

MIN_INVOKE_TIMES = 176
MAX_INVOKE_TIMES = 237
EXECUTOR_KILLER = GracefulKiller()


def config(path, data=None):
    if not data:
        with open(path, mode="r") as conf:
            return json.load(conf)

    # fast-fail
    json.loads(json.dumps(data))
    with open(path, mode="w") as conf:
        json.dump(data, conf)

    # with open(path, mode="r+") as conf:
    #     if not data:
    #         return json.load(conf)
    #     json.dump(data, conf, sort_keys=True, indent=4)


def get_access_token(app):
    try:
        return requests.post(
            "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            data={
                "grant_type": "refresh_token",
                "refresh_token": app["refresh_token"],
                "client_id": app["client_id"],
                "client_secret": app["client_secret"],
                "redirect_uri": app["redirect_uri"],
            },
        ).json()
    except Exception:
        return {}


def invoke_api(path):
    app = config(path)
    tokens = get_access_token(app)
    access_token = tokens.get("access_token", "")
    refresh_token = tokens.get("refresh_token", "")
    username = app["username"]

    if len(access_token) < 5 or len(refresh_token) < 5:
        return f"✘ 账号 [{username}] 调用失败."
    
    def generateStr(length):
        searchstr = ''
        chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
        for i in range(length):
            searchstr += chars[random.randint(0, len(chars) - 1)]
        return searchstr

    apis = [
        # User API 组
        'https://graph.microsoft.com/v1.0/groups',
        'https://graph.microsoft.com/v1.0/groups?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/groups?$select=displayName',
        'https://graph.microsoft.com/v1.0/groups?$select=id',
        'https://graph.microsoft.com/v1.0/groups?$select=displayName,id',
        'https://graph.microsoft.com/v1.0/groups?$select=displayName,id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/groups?$select=groupTypes',
        'https://graph.microsoft.com/v1.0/groups?$select=groupTypes&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/groups?$select=owners,groupTypes',
        'https://graph.microsoft.com/v1.0/groups?$select=owners,groupTypes&$top=' + str(random.randint(1, 20)),
        # site API 组
        'https://graph.microsoft.com/v1.0/sites/root',
        'https://graph.microsoft.com/v1.0/sites/root?$select=displayName',
        'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl',
        'https://graph.microsoft.com/v1.0/sites/root?$select=displayName,webUrl',
        'https://graph.microsoft.com/v1.0/sites/root?$select=displayName,webUrl&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/sites/root?$select=displayName,webUrl&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/sites/root?$select=displayName,webUrl&$top=' + str(random.randint(1, 20)),
        # me API 组
        'https://graph.microsoft.com/v1.0/me/',
        'https://graph.microsoft.com/v1.0/me?$select=displayName',
        'https://graph.microsoft.com/v1.0/me?$select=displayName,mail',
        'https://graph.microsoft.com/v1.0/me?$select=id',
        'https://graph.microsoft.com/v1.0/me?$select=displayName,mail&$top=' + str(random.randint(1, 20)),
        # me event API 组
        'https://graph.microsoft.com/v1.0/me/events',
        'https://graph.microsoft.com/v1.0/me/events?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=subject',
        'https://graph.microsoft.com/v1.0/me/events?$select=subject&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=subject&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=subject&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=organizer',
        'https://graph.microsoft.com/v1.0/me/events?$select=organizer&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=organizer&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=organizer&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=subject,organizer',
        'https://graph.microsoft.com/v1.0/me/events?$select=subject,organizer&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=subject,organizer&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/events?$select=subject,organizer&$top=' + str(random.randint(1, 20)),
        # me people API 组
        'https://graph.microsoft.com/v1.0/me/people',
        'https://graph.microsoft.com/v1.0/me/people?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=displayName',
        'https://graph.microsoft.com/v1.0/me/people?$select=id',
        'https://graph.microsoft.com/v1.0/me/people?$select=id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=displayName,id',
        'https://graph.microsoft.com/v1.0/me/people?$select=displayName,id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=displayName,id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=displayName,id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=displayName&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=displayName&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/people?$select=displayName&$top=' + str(random.randint(1, 20)),
        # me contacts API 组
        'https://graph.microsoft.com/v1.0/me/contacts',
        'https://graph.microsoft.com/v1.0/me/contacts?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/contacts?$select=assistantName',
        'https://graph.microsoft.com/v1.0/me/contacts?$select=id',
        'https://graph.microsoft.com/v1.0/me/contacts?$select=manager,nickName',
        'https://graph.microsoft.com/v1.0/me/contacts?$select=manager,nickName&$top=' + str(random.randint(1, 20)),
        # calendar API 组
        'https://graph.microsoft.com/v1.0/me/calendars',
        'https://graph.microsoft.com/v1.0/me/calendars?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id',
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=canEdit&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=canEdit&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=canEdit&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=allowedOnlineMeetingProviders',
        'https://graph.microsoft.com/v1.0/me/calendars?$select=allowedOnlineMeetingProviders&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=allowedOnlineMeetingProviders&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=allowedOnlineMeetingProviders&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id,isRemovable',
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id,isRemovable&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id,isRemovable&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id,isRemovable&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$select=id,canEdit&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/calendars?$top=' + str(random.randint(1, 20)),
        # drive API 组
        'https://graph.microsoft.com/v1.0/me/drive',
        'https://graph.microsoft.com/v1.0/me/drive?$select=name',
        'https://graph.microsoft.com/v1.0/me/drive?$select=name&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/drive?$select=name&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/drive?$select=name&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/drive?$select=owner,webUrl',
        'https://graph.microsoft.com/v1.0/me/drive?$select=owner,webUrl&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/drive?$select=owner,webUrl&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/drive?$select=owner,webUrl&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/drive?$select=driveType,webUrl&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/drive?$select=driveType,webUrl&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/drive?$select=driveType,webUrl&$top=' + str(random.randint(1, 20)),
        # drive onenote API 组
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks',
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=displayName',
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=displayName&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=displayName,isShared',
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=displayName,isShared&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$orderby=displayName desc',
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$orderby=displayName desc&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$orderby=createdDateTime desc',
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$orderby=createdDateTime desc&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$orderby=lastModifiedDateTime desc',
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$orderby=lastModifiedDateTime desc&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$orderby=lastModifiedDateTime desc&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$orderby=lastModifiedDateTime desc&$top=' + str(random.randint(1, 20)),
        # drive onenote section API 组
        'https://graph.microsoft.com/v1.0/me/onenote/sections',
        'https://graph.microsoft.com/v1.0/me/onenote/sections?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/sections?$select=displayName',
        'https://graph.microsoft.com/v1.0/me/onenote/sections?$select=displayName,createdDateTime',
        'https://graph.microsoft.com/v1.0/me/onenote/sections?$select=displayName,createdDateTime&$top=' + str(random.randint(1, 20)),
        # drive onenote pages API 组
        'https://graph.microsoft.com/v1.0/me/onenote/pages',
        'https://graph.microsoft.com/v1.0/me/onenote/pages?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/pages?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/pages?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/pages?$select=id',
        'https://graph.microsoft.com/v1.0/me/onenote/pages?$select=id,order&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/pages?$select=id,order&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/onenote/pages?$select=id,order&$top=' + str(random.randint(1, 20)),
        # drive onenote resources API 组
        'https://graph.microsoft.com/v1.0/me/messages',
        'https://graph.microsoft.com/v1.0/me/messages?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=createdDateTime',
        'https://graph.microsoft.com/v1.0/me/messages?$select=importance',
        'https://graph.microsoft.com/v1.0/me/messages?$select=importance,isRead&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=importance,isRead&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=importance,isRead&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=hasAttachments&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=hasAttachments&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=hasAttachments&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=toRecipients,from&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=toRecipients,from&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$select=toRecipients,from&$top=' + str(random.randint(1, 20)),
        # 带有搜索参数的API
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"',
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"',
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"',
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(2, 4)) + '"',
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(2, 4)) + '"',
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"&$select=toRecipients,from',
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"&$select=toRecipients,from&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"&$select=toRecipients,from&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"&$select=toRecipients,from&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/messages?$search="' + generateStr(random.randint(1, 2)) + '"&$select=toRecipients,from&$top=' + str(random.randint(1, 20)),
        # drive onenote resources API 组0
        'https://graph.microsoft.com/v1.0/me/mailFolders',
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules?$select=displayName,sequence,conditions&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules?$select=displayName,sequence,conditions&$top=' + str(random.randint(1, 20)),
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules?$select=displayName,sequence,conditions&$top=' + str(random.randint(1, 20)),
        # 搜索用户的邮件测试
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'high'",
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'low'",
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'normal'",
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'high' and isRead eq false",
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'low' and isRead eq false",
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'normal' and isRead eq false",
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'high' and isRead eq true",
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'low' and isRead eq true",
        "https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'normal' and isRead eq true",
    ]
    headers = {"Authorization": f"Bearer {access_token}"}

    def single_period(period):
        if EXECUTOR_KILLER.kill_now:
            return ""

        result = "=" * 100 + "\n"
        random.shuffle(apis)
        probability = random.random()
        for api in apis:
            if random.random() < probability:
                continue
            try:
                if requests.get(api, headers=headers).status_code == 200:
                    result += "{:>20s} | {:>6s} | {:<50s}\n".format(
                        f"账号: {username}", f"周期: {period}", f"成功: {api}"
                    )
            except Exception:
                # time.sleep(random.random()*3)
                pass

            if EXECUTOR_KILLER.kill_now:
                return result

        return result

    with ThreadPoolExecutor() as executor:
        max = random.randint(MIN_INVOKE_TIMES, MAX_INVOKE_TIMES)
        futures = [executor.submit(single_period, period) for period in range(1, max)]
        result = "".join((f.result() for f in futures))

    # save refresh_token
    app["refresh_token"] = refresh_token
    config(path, app)

    return f"{result}✔ 账号 [{username}] 调用成功."

if __name__ == "__main__":
    multi_accounts_task(invoke_api)
