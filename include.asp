<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer				= true
Response.CharSet			= "utf-8"
Response.ContentType		= "text/html"
Response.CacheControl		= "no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires			= -1
Response.ExpiresAbsolute	= Now()-1
%>