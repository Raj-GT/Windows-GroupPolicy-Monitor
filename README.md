# Windows-GroupPolicy-Monitor
*PowerShell script to monitor the domain (or select OU) for Group Policy changes and take automatic backups and optionally, alert via e-mail*

When run (ideally on a recurring schedule via Task Scheduler) the script will check Group Policies linked under $watchedOU for changes and perform an automatic backup of new and changed policies. It will also generate individual HTML/XML reports of the policies and save it with the backups with an option to send a summary of changes via e-mail.

Each set of backup is created under it's own folder and kept indefinitely.

An article I wrote on an earlier version of these scripts is available at https://www.experts-exchange.com/articles/30751/Automating-Group-Policy-Backups.html