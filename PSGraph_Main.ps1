.".\PSGraph_Methods.ps1"

##### トークン取得の実行 ############################################################################################
# --------- トークン取得 ---------
# アプリケーション許可 (Application permissions) で実行する場合
$tokenResponse = getApplicationAccessToken

# ユーザー委任 (Delegated permissions) で実行する場合
$tokenResponse = getDelegatedAccessToken



##### リクエストの実行 ############################################################################################
# ---------  GET メソッドの実行 ---------
# 変数: getURL に、リクエスト URL を設定します
$getURL = 'https://graph.microsoft.com/v1.0/XXXXXXXXXXXXXXXXXXXXXXXX'
# ＜例＞ $getURL = 'https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq "Team")' など

# アクセストークン ($tokenResponse) をもって、リクエスト URL (getURL) に対して GET リクエストを実行します
doGET $tokenResponse $getURL


# # ---------  POST メソッドの実行 ---------
# 変数: postURL に、リクエスト URL を設定します
$postURL = 'https://graph.microsoft.com/v1.0/XXXXXXXXXXXXXXXXXXXXXXXX'
# ＜例＞ $postURL = 'https://graph.microsoft.com/v1.0/groups' など

# 変数: postBody に、リクエスト Body を設定します
$postBody = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
# ＜例＞ $postBody = '
#     {
#         "description": "Self help community for library",
#         "displayName": "Library Assist",
#         "groupTypes": [
#         "Unified"
#         ],
#         "mailEnabled": true,
#         "mailNickname": "library",
#         "securityEnabled": false
#     }'

# アクセストークン ($tokenResponse) とリクエストBody ($postBody) をもって、リクエスト URL (postURL) に対して POST リクエストを実行します
doPOST $tokenResponse $postURL $postBody


# ---------  DELETE メソッドの実行 ---------
# 変数: deleteURL に、リクエスト URL を設定します
$deleteURL = 'https://graph.microsoft.com/v1.0/XXXXXXXXXXXXXXXXXXXXXXXX'

# アクセストークン ($tokenResponse) をもって、リクエスト URL (deleteURL) に対して DELETE リクエストを実行します
doDELETE $tokenResponse $deleteURL


# ---------  Incoming Webhook の実行 ---------
# 変数: webhookURL に、リクエスト URL (Webhook URL) を設定します
$webhookURL = 'https://XXXXXXXXXX.webhook.office.com/webhookb2/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'

# 変数: wpostBody に、リクエスト Body を設定します
$wpostBody = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
# ＜例＞ $wpostBody = '{
#     "text" : "This is a webhook post test from a PSGraph script."
# }'

# リクエストBody ($wpostBody) をもって、リクエスト URL (webhookURL) に対して POST リクエストを実行します
doWEBHOOK $webhookURL $wpostBody