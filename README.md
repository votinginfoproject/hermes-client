# hermes-client

Powershell upload script for the hermes VIP data processing tool.
Intended to be run by election officials on their systems to automate
periodically uploading data to VIP.

## Requirements

You need Powershell version 4 or later to run this script (see
http://social.technet.microsoft.com/wiki/contents/articles/21016.how-to-install-windows-powershell-4-0.aspx
for more information).

## Usage

First, put the files to upload into a `data` subdirectory of the directory
containing this script. Then, run the script, making sure to provide a username,
password, fips code, and election date:

  powershell -File hermes-upload.ps1 -username some_user -password 53cr37 -fips 8 -electionDate 2017-11-07

This will create a file called `vipFeed-8-2017-11-07.zip` and upload it to the
VIP server.
