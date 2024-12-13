# Set up self-hosted ScreenConnect LetsEncrypt TLS certificates

## Context and Background

When using a self-hosted ConnectWise ScreenConnect server, if you don't want to manually renew the TLS/SSL certificates annually, you must set up TLS certificates using LetsEncrypt. This is a feature that ConnectWise has rejected including, but the application does not use IIS directly, so the certificate that's used must be manually bound to the application initially and after renewal. This script is designed for Windows and is NOT relevant for Linux or macOS, where older versions of ScreenConnect ran and where there is more documentation available online for automating this process (and ConnectWise doesn't support SSO on any servers except Windows, which is important for some people).

### Introduction and Overview

There were too many possible options, libraries, and methods to try but no simple process with a very straightforward installation script that was tested for use on modern Windows versions with the default ScreenConnect web server configuration that didn't involve proxies or third parties like CloudFlare, so I assembled this process that requires the very well-written, easy-to-use, and frequently updated win-acme tool and a tiny script to install the certificate. Hopefully this provides the push to stop renewing certificates manually for ScreenConnect!

## What this is not

This process assumes you already have an operational [ScreenConnect](https://screenconnect.com) installation on your own self-hosted server, and that it's already configured with a valid TLS certificate, perhaps issued by RapidSSL or any other certificate authority where you buy certificates and manually retrieve and install them, but that you'd like to switch to using LetsEncrypt certificates instead.

This process assumes you are having LetsEncrypt configured for TLS, are using the built-in web server and not proxying the web server through a third party like CloudFlare or yourself using nginx or Caddy, so it doesn't walk you through that process. It also assumes you have locked down the TLS settings yourself and validated it using a service like [Qualys SSLLabs](https://www.ssllabs.com/ssltest/) in order to ensure only modern and secure TLS configurations are used.

## Alternate Solutions

We will assume you will not use a self-signed certificate, which is not recommended as it is not secure and is subject to being revoked. This is not a real option for a public server.

1. Use a third-party tool like CertifyTheWeb to obtain and renew a certificate from LetsEncrypt. This should work and is well documented, but as of this writing costs approximately $60 per year, which is substantially more than a basic RapidSSL certificate (though you'd have to spend the time renewing it manually).
2. Create your own PowerShell or command-line based script to obtain and renew a certificate from LetsEncrypt and apply it to ScreenConnect. There are various PowerShell modules for ACME v2 that would work; this solution is relatively close but it uses a third-party free tool called [win-acme](https://www.win-acme.com/) to obtain and renew the certificate, calling the Command Line script provided here to remove the old and install the new certificate for ScreenConnect use.

## This Solution

### win-acme and installation script
The script in this repository, `BindNewScreenConnectCert.cmd` (the installation script that win-acme runs after obtaining the certificate) should be placed into a folder with the extracted win-acme zip file contents. From the [win-acme](https://www.win-acme.com/) website, on your Windows server where ScreenConnect is running, place the extracted win-acme zip file contents in a permanent location. You may wish to create a folder like C:\certs to hold the files, or place it in an existing location. I'm using `C:\certs\win-acme-pluggable` in this example and you're welcome to copy mine or choose your own. The executable that matters is `C:\certs\win-acme-pluggable\wacs.exe` or the same file inside the extracted zip on your system.

### Validating for certificate to be issued
Configuring win-acme to obtain a certificate and validate properly is left up to you, but their documentation is relatively good. If you've run certbot on Linux it's not terribly different. Validation can be completed via multiple web-based or DNS-based method. If web-based, you'll need to allow port 80 for HTTP access to the server to handle validation, all the way through the firewall. win-acme has a plugin system for many DNS providers that can validate with DNS instead (including Microsoft Azure DNS, CloudFlare, and many others).

### Determine the bound IP and port of ScreenConnect server
If you run this command from an elevated command prompt, it will display the TLS certificates that are mapped to listening IP:port combinations:

`netsh http show sslcert`

The first field should be named "IP/port" and should be in a global "listen on all interfaces" format like `0.0.0.0:443` or a specific IP (public or private) like `10.0.0.2:443`. As noted above, you need ScreenConnect already functional and listening or to figure this out first. Some of the potential commands and screenshots of sample output are listed at [Replace/renew your SSL certificate in ConnectWise OnPrem](https://nadavsvirsky.medium.com/replace-renew-your-ssl-certificate-in-connectwise-onprem-82fc15352227). Gather the IP:port string from the above command and record it for the next step.

If you need to set this up for the first time, there are instructions at [ConnectWise Control (ScreenConnect) On-Premise SSL Installation Woes… Here’s the Secret to Using an Alternate Port (not 443), Working With SSL](https://asheroto.medium.com/screenconnect-on-premise-ssl-installation-woes-heres-the-secret-to-get-an-alternate-port-working-931f240ced92) although you can simply use the default port 443, but that's up to you. (Note the article recommends RapidSSL certificates and suggests the benefit is not renewing every 90 days; we are automating this process so that's not ever necessary but the manual setup steps are similar and useful! You don't have to use the ScreenConnectSSLConfigurator they recommend if you just want to get a certificate with win-acme in the first place which will be used in the same way, and win-acme handles all public and private keys and CSRs in addition to renewals.) The section starting "To show existing bindings" has the more useful information in this article, along with the next "Binding the Certificate — Method 1" section; Method 2 uses the ScreenConnectSSLConfigurator again but shouldn't be necessary.

### Run wacs.exe and obtain a new certificate
Open an elevated command prompt or PowerShell prompt (either is fine) and change to the folder you extracted the win-acme zip file contents to, then run `.\wacs.exe` from the command prompt. It will launch a menu and let you walk through obtaining a new certificate, using the `M` option to Create certificate (full options).

Walk through the wizard to obtain a new certificate. Choose your own validation options to match what you wish to use and have available. The "Store" you choose should be "Windows Certificate Store (Local Computer)" and when prompted for how you would like to store the certificate, choose "2: [My] - General computer store (for Exchange/RDS)" as the option. If you want to store the certificate a second way, choose a second set of options but this should not be necessary and you can choose "5: No (additional) store steps" (the default).

For the Installation step, choose "2: Start external script or program" and enter the full path to the BindNewScreenConnectCert.cmd script and press Enter, for example in the folder above:

`C:\certs\win-acme-pluggable\BindNewScreenConnectCert.cmd`

You'll then be prompted for Parameters. The script requires two parameters, the first is the thumbprint of the certificate to bind. The second is the IP:port to bind to (which it will unbind first). Enter the placeholder variable and the IP:port string all on one line but separate by a space as the Parameters value, exactly like this with the curly braces (customize only the IP:port portion), then press Enter:

`{CertThumbprint} 0.0.0.0:443`

Choose 3 to perform no additional installation steps and press Enter. The certificate will be obtained, assuming no errors, saved to the Local Computer Personal Certificates, then the installation script will be called by win-acme and the script will unbind the existing certificate (if any), bind the new one by thumbprint (both using `netsh` commands). **Warning: This process will stop the ScreenConnect Web Server service! After rebinding the new certificate, it will start it again. However, connections will be interrupted!**

By default, Task Scheduler will have a task added by win-acme to run daily to attempt certificate renewal. However, certificates won't be renewed until 55 days have elapsed by default. The `settings.json` file can be edited in the folder alongside `wacs.exe` and the settings are well-documented on the win-acme website. You may wish to provide different schedule times (perhaps not during the day, since the process stops the ScreenConnect Web Server service and restarts it after rebinding).

You can exit the win-acme menus now and run `netsh http show sslcert` to verify the new certificate is bound properly.

### Editing the certificate
You can re-run `wacs.exe` and choose Manage Renewals to walk through the reviewing or editing any of the settings you set above. After editing, the certificate will be reinstalled including reloading the ScreenConnect service! Note that by default the renewal will use the cached certificate and will not reach out to LetsEncrypt if it's been less than a day, by default, to avoid rate-limiting.

Renewals should run automatically, but you can edit `settings.json` to configure an SMTP server and email details to send failure and, optionally, success notifications to you via email, or you can monitor the expiration externally with a tool like UptimeRobot or others that will alert you if the renewal fails with enough time to remediate the issue.

### Examining the command line generated for renewals
The Renewals submenu in `wacs.exe` also has an "L" option that will display the command line you can run to renew the certificate manually; if nothing else this is nice to review what the command line options being used are, and see the script call and parameters that are set. It also displays the next renewal attempt date of all certificates being managed for review.

### Reviewing the script run in Event Log
win-acme does a great job of logging, to its own log and to the Windows Application Event Log. If you want to see the actual output of the `BindNewScreenConnectCert.cmd` script, you can open the Windows Event Log and search for Event ID 7703, the source will be "win-acme" if you'd like to set a filter. The output of the script will be in the "Message" field of the event on the General tab of this event ID when it's logged. You should see the service stopping, the binding deleted, re-added, and the service told to start again.

## Summary
These instructions could likely be more concise and clearer, possibly with some screenshots, but this is the first pass. If you have any suggestions for improvement, please let me know! Pull requests are also welcome. Hope this is helpful!

# References

A few links that were helpful in my research, not mentioned above:

- [win-acme documentation for installation scripts](https://www.win-acme.com/reference/plugins/installation/script)
- [official ScreenConnect docs on updating SSL cert](https://docs.connectwise.com/ConnectWise_ScreenConnect_Documentation/On-premises/Advanced_setup/SSL_certificate_installation/Install_and_bind_an_SSL_certificate_on_a_Windows_server)
- [Installing an SSL certificate manually on Windows](https://docs.connectwise.com/ConnectWise_ScreenConnect_Documentation/On-premises/Advanced_setup/SSL_certificate_installation/Install_and_bind_an_SSL_certificate_on_a_Windows_server)
- [official ScreenConnect docs on the SSL Configurator tool](https://docs.connectwise.com/ConnectWise_ScreenConnect_Documentation/On-premises/Advanced_setup/SSL_certificate_installation/SSL_Configurator)