@REM PASS THESE ARGUMENTS: cert-thumbprint certbinding-ip:port
@REM cert-thumbprint: win-acme's {CertThumbprint} variable
@REM certbinding-ip:port: the IP:port value of existing ScreenConnect 
@REM   cert in the output of this command:
@REM         netsh http show sslcert
@ECHO off
echo "Installing ScreenConnect TLS Certificate"
net stop "ScreenConnect Web Server"

ECHO Now installing cert %1
netsh http delete sslcert ipport=%2
netsh http add sslcert ipport=%2 certhash=%1 appid="{00000000-0000-0000-0000-000000000000}"

net start "ScreenConnect Web Server"
