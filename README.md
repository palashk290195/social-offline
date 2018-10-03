# social-offline
Project Social Offline for interacting with peers

Code for Social Offline:

Usage:

python google-sheets.py Google-form-response-sheet-id Option

Google form response sheet id ->  It is the id of google form response sheet, you can easily find it in the google sheets url: example 1nMidXBWZhAqFSfiQ_hrzwraO3XyDX1G426lTt2TGqDc

Option ->

Put 1, if you want to just calculate the matchings and print them

Put 2, if you want to calculate the matchings as well as send emails to the participants

Warning: Always use Option = '1' for the first time and verify the matchings first.

If you get an error:
Cannot access token.json: No such file or directory
  warnings.warn(_MISSING_FILE_MESSAGE.format(filename))
usage: google-sheets.py [--auth_host_name AUTH_HOST_NAME]
                        [--noauth_local_webserver]
                        [--auth_host_port [AUTH_HOST_PORT [AUTH_HOST_PORT ...]]]
                        [--logging_level {DEBUG,INFO,WARNING,ERROR,CRITICAL}]
google-sheets.py: error: unrecognized arguments:

Then, run python google-sheets.py and authenticate first. After that run the command with options.
