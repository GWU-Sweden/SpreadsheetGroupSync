
# SpreadsheetGroupSync

The following command will read e-mail addresses from the G-Suite Spreadsheet and add the approved members to the G-Suite Group as members, if they aren't already. Any other members in the group not in the spreadsheet will be removed.

```
python sync.py <MEMBERSHIP_SPREADSHEET_ID> <MEMBERS_GROUP_ID>
```

This is quite handy as e.g. G-Suite documents can be shared to G-Suite groups so that access management to "members only" documents is a little bit less of a hassle.

You probably need to change the `SPREADSHEET_RANGE_NAME`, `SPREADSHEET_EMAIL_INDEX`, `SPREADSHEET_STATUS_INDEX` in the script. A nice feature would be to add this to the command line maybe. :fox_face:
