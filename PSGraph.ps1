Add-Type -AssemblyName System.Web

##### リクエスト実行用メソッド ###################
# --------- GET メソッド ---------
# Get Method
function doGET (
    [Parameter(Mandatory=$true)]$Uri,
    [Parameter(Mandatory=$true)][string]$Mode,
    [Parameter(Mandatory=$false)][string]$clientid,
    [Parameter(Mandatory=$false)][string]$clientSecret,
    [Parameter(Mandatory=$false)][string]$tenantName,
    [Parameter(Mandatory=$false)][string]$UPN,
    [Parameter(Mandatory=$false)][string]$Password
    ){

    $tokenResponse = getAccessToken $Mode $clientid $clientSecret $tenantName
    
    $result = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $Uri -Method Get
    return $result
}
# --------------------------------

# --------- POST メソッド ---------
# POST Method
function doPOST (
    [Parameter(Mandatory=$true)]$Uri,
    [Parameter(Mandatory=$true)]$Body,
    [Parameter(Mandatory=$true)][string]$Mode,
    [Parameter(Mandatory=$false)][string]$clientid,
    [Parameter(Mandatory=$false)][string]$clientSecret,
    [Parameter(Mandatory=$false)][string]$tenantName,
    [Parameter(Mandatory=$false)][string]$UPN,
    [Parameter(Mandatory=$false)][string]$Password
    ){

    $tokenResponse = getAccessToken $Mode $clientid $clientSecret $tenantName

    $result = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $Uri -Method POST -Body $Body -ContentType 'application/json'
    return $result
}
# --------------------------------

# --------- DELETE メソッド ---------
function doDELETE (
    [Parameter(Mandatory=$true)]$Uri,
    [Parameter(Mandatory=$true)][string]$Mode,
    [Parameter(Mandatory=$false)][string]$clientid,
    [Parameter(Mandatory=$false)][string]$clientSecret,
    [Parameter(Mandatory=$false)][string]$tenantName,
    [Parameter(Mandatory=$false)][string]$UPN,
    [Parameter(Mandatory=$false)][string]$Password
    ){
    
    $tokenResponse = getAccessToken $Mode $clientid $clientSecret $tenantName

    $result = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $Uri -Method DELETE
    return $result
}
# --------------------------------

# --------- PATCH メソッド ---------
function doPATCH (
    [Parameter(Mandatory=$true)]$Uri,
    [Parameter(Mandatory=$true)]$Body,
    [Parameter(Mandatory=$true)][string]$Mode,
    [Parameter(Mandatory=$false)][string]$clientid,
    [Parameter(Mandatory=$false)][string]$clientSecret,
    [Parameter(Mandatory=$false)][string]$tenantName,
    [Parameter(Mandatory=$false)][string]$UPN,
    [Parameter(Mandatory=$false)][string]$Password
    ){
    
    $tokenResponse = getAccessToken $Mode $clientid $clientSecret $tenantName

    $result = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $Uri -Method PATCH -Body $Body -ContentType 'application/json'
    return $result
}
# --------------------------------




##### トークン取得用メソッド ###################

# --------- トークン取得を実行 ---------
function getAccessToken (
    [Parameter(Mandatory=$true)][string]$Mode,
    [Parameter(Mandatory=$false)][string]$clientid,
    [Parameter(Mandatory=$false)][string]$clientSecret,
    [Parameter(Mandatory=$false)][string]$tenantName
    ){
    
    if($Mode -eq "Application"){                        
        $tokenResponse = getApplicationAccessToken $clientid $clientSecret $tenantName      # アプリケーション許可 (Application permissions) で実行する場合
    }
    elseif($Mode -eq "Delegate"){
        $tokenResponse = getDelegatedAccessToken $clientid $clientSecret $tenantName        # ユーザー委任 (Delegated permissions) で実行する場合
    }
    return $tokenResponse
}

# -------------- アプリケーション許可 (Application permissions) 用メソッド -------------- 
function getApplicationAccessToken (
    [Parameter(Mandatory=$true)][string]$clientid,              #アプリケーションID
    [Parameter(Mandatory=$true)][string]$clientSecret,          #クライアント シークレット
    [Parameter(Mandatory=$true)][string]$tenantName             #テナント名
    ){

    $scope                  = "https://graph.microsoft.com/.default"

    $clientIDEncoded        = [System.Web.HttpUtility]::UrlEncode($clientid)
    $clientSecretEncoded    = [System.Web.HttpUtility]::UrlEncode($clientSecret)
    $scopeEncoded           = [System.Web.HttpUtility]::UrlEncode($scope)


    $body = "client_id=$clientIdEncoded&scope=$scopeEncoded&client_secret=$clientSecretEncoded&grant_type=client_credentials"
    $tokenResponse = Invoke-RestMethod https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token `
        -Method Post -ContentType "application/x-www-form-urlencoded" `
        -Body $body `
        -ErrorAction STOP

    return $tokenResponse
}
# ---------------------------------------------------------------------------

# -------------- ユーザー委任 (Delegated permissions) 用メソッド --------------
function getDelegatedAccessToken (
    [Parameter(Mandatory=$true)][string]$clientid,              #アプリケーションID
    [Parameter(Mandatory=$true)][string]$clientSecret,          #クライアント シークレット
    [Parameter(Mandatory=$true)][string]$tenantName,            #テナント名
    [Parameter(Mandatory=$true)][string]$UPN,                   #委任ユーザーのUPN
    [Parameter(Mandatory=$true)][string]$Password               #委任ユーザーのパスワード
    ){

    $scope                  = "https://graph.microsoft.com/.default"

    $clientIDEncoded        = [System.Web.HttpUtility]::UrlEncode($clientid)
    $clientSecretEncoded    = [System.Web.HttpUtility]::UrlEncode($clientSecret)
    $scopeEncoded           = [System.Web.HttpUtility]::UrlEncode($scope)
    $redirectUriEncoded     = [System.Web.HttpUtility]::UrlEncode($redirectUri)
    $usernameEncoded        = [System.Web.HttpUtility]::UrlEncode($UPN)
    $passwordEncoded        = [System.Web.HttpUtility]::UrlEncode($Password)
    
    # get Access Token
    $body = "username=$usernameEncoded&password=$passwordEncoded&grant_type=password&redirect_uri=$redirectUriEncoded&client_id=$clientIdEncoded&client_secret=$clientSecretEncoded&scope=$scopeEncoded"
    $tokenResponse = Invoke-RestMethod https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token `
        -Method Post -ContentType "application/x-www-form-urlencoded" `
        -Body $body `
        -ErrorAction STOP
    return $tokenResponse
}
# ------------------------------------------------------------------









##### 実行方法 ###################

# --------- 実行メソッドについて ---------
# GET    で実行する場合は【doGET】コマンドを実行ください
# POST   で実行する場合は【doPOST】コマンドを実行ください
# DELETE で実行する場合は【doDELETE】コマンドを実行ください
# PATCH  で実行する場合は【doPATCH】コマンドを実行ください
# ---------------------------------------


# --------- アクセス許可の種類について ---------
# アプリケーション許可 (Application permissions) で実行する場合は、-Mode オプションで "Application" と記載ください
# 例:  doGET -Mode Application -Uri "https://graph.microsoft.com/v1.0/teams" -clientid XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX -clientSecret XXXXXXXXXXXXXXXXXXXX -tenantName XXXXXXXX.onmicrosoft.com

# ユーザー委任 (Delegated permissions) で実行する場合は、-Mode オプションで "Delegate" と記載ください
# 例:  doGET -Mode Delegate -Uri "https://graph.microsoft.com/v1.0/me/joinedTeams" -clientid XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX -clientSecret XXXXXXXXXXXXXXXXXXXX -tenantName XXXXXXXX.onmicrosoft.com
# --------------------------------------------

# --------- リクエストBody の書き方について ---------
# 以下のような形式でリクエストBody を指定し、コマンドを実行ください

# $postBody = @"
# {
#    "template@odata.bind":"https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
#    "displayName":"サンプルチーム",
#    "description":"Graph API でのチーム作成テストです",
#     "members":[
#         {
#          "@odata.type":"#microsoft.graph.aadUserConversationMember",
#          "roles":[
#             "owner"
#          ],
#         "user@odata.bind":"https://graph.microsoft.com/v1.0/users('91f3cf31-cef6-47e8-a6a4-682f85694222')"
#         }
#     ],
#     "MemberSettings": {
#         "allowCreatePrivateChannels": false,
#         "allowDeleteChannels": false,
#         "allowCreateUpdateRemoveConnectors": false,
#         "allowuploadapps": false

#     }
# }
# "@

# doPOST -Mode Application -Uri "https://graph.microsoft.com/v1.0/teams" -clientid XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX -clientSecret XXXXXXXXXXXXXXXXXXXX -tenantName XXXXXXXX.onmicrosoft.com -Body $postBody

# --------------------------------------------