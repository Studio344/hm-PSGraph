##### トークン取得用のメソッド ############################################################################################
# -------------- アプリケーション許可 (Application permissions) -------------- 
function getApplicationAccessToken {
    $clientId       = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"      #アプリケーションID
    $clientSecret   = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"      #クライアント シークレット
    $tenantName     = "XXXXXXXXXX.onmicrosoft.com"              #テナント名

    $tokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $clientId
        Client_Secret = $clientSecret
    }
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $tokenBody
    return $tokenResponse
}
# ---------------------------------------------------------------------------

# -------------- ユーザー委任 (Delegated permissions) --------------
function getDelegatedAccessToken {
    # Azure AD 上に登録したアプリケーション情報
    $clientid       = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"      #アプリケーションID
    $clientSecret   = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"      #クライアント シークレット
    $tenantName     = "XXXXXXXXXX.onmicrosoft.com"              #テナント名
    $scope          = "https://graph.microsoft.com/.default"
    
    # 委任するユーザー アカウント情報
    $username = "XXXXXXXXXX@XXXXXXXXXX.work"
    $password = "XXXXXXXXXXXXXXXXXXXX"

    # UrlEncode the ClientID and ClientSecret and URL's for special characters 
    Add-Type -AssemblyName System.Web
    $clientIDEncoded        = [System.Web.HttpUtility]::UrlEncode($clientid)
    $clientSecretEncoded    = [System.Web.HttpUtility]::UrlEncode($clientSecret)
    $redirectUriEncoded     = [System.Web.HttpUtility]::UrlEncode($redirectUri)
    $scopeEncoded           = [System.Web.HttpUtility]::UrlEncode($scope)
    $usernameEncoded        = [System.Web.HttpUtility]::UrlEncode($username)
    $passwordEncoded        = [System.Web.HttpUtility]::UrlEncode($password)
    
    # get Access Token
    $body = "username=$usernameEncoded&password=$passwordEncoded&grant_type=password&redirect_uri=$redirectUriEncoded&client_id=$clientIdEncoded&client_secret=$clientSecretEncoded&scope=$scopeEncoded"
    $tokenResponse = Invoke-RestMethod https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token `
        -Method Post -ContentType "application/x-www-form-urlencoded" `
        -Body $body `
        -ErrorAction STOP
    return $tokenResponse
}
# ------------------------------------------------------------------



##### リクエスト実行用のメソッド ############################################################################################

# --------- GET メソッド ---------
# Get Method
function doGET ($tokenResponse, $getURL) {
    $result = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $getURL -Method Get
    $result
}
# --------------------------------

# --------- POST メソッド ---------
# POST Method
function doPOST ($tokenResponse, $postURL, $postBody) {
    $result = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $postURL -Method POST -Body $postBody -ContentType 'application/json'
    $result
}
# --------------------------------

# --------- DELETE メソッド ---------
function doDELETE ($tokenResponse, $deleteURL) {
    $result = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $deleteURL -Method DELETE
    $result
}
# --------------------------------

# --------- Webhook メソッド ---------
function doWEBHOOK ($webhookURL, $wpostBody) {
    $wpostBody = [Text.Encoding]::UTF8.GetBytes($wpostBody)
    $result = Invoke-RestMethod -Uri $webhookURL -Method POST -Body $wpostBody -ContentType 'application/json'
    $result
}
# -----------------------------------
