<%@ Language=VBScript %>
<%

Dim awskey, awsscret, requestSign, objKey, bucket, host, proto
Dim expire, sig, result, url1, url2
Dim sha1

awskey = "access key"
awsscret = "secret"

expire = 1457400000 'unix timestamp to expire
bucket = "bucket name"
objKey = "object key"
host = "s3.example.com"
proto = "https"

requestSign = "GET" & vbLf & vbLf & vbLf & expire & vbLf & "/" & bucket & "/" & objKey

set sha1 = GetObject("script:" & Server.MapPath("sha1.wsc"))
sha1.hexcase = 0

result = sha1.b64_hmac_sha1(awsscret, requestSign)
sig = Server.URLEncode(result & "=")
url1 = proto & "://" & host & "/" & bucket & "/" & objKey & "?AWSAccessKeyId=" & awskey & "&Signature=" & sig & "&Expires=" & expire
url2 = proto & "://" & bucket & "." & host & "/" & objKey & "?AWSAccessKeyId=" & awskey & "&Signature=" & sig & "&Expires=" & expire

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <title>S3</title>
</head>
<body>
            Request Signature  : <%= requestSign %> <br/>
            HMAC-SHA1 result: <%= result %> <br/>
            Signature result: <%= sig %> <br/>
            URL Path style: <a href="<%= url1 %>" target="_blank"><%= url1 %></a> <br/>
            URL Virtual style: <a href="<%= url2 %>" target="_blank"><%= url2 %></a> <br/>
</body>
</html>


<%
'Free resource
Set sha1 = Nothing
%>
