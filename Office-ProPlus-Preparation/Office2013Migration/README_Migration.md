##**Migration Scripts**

These scripts assist in grabbing registry key settings and creating log files for current Office 2013 users.

###**Set-GlobalDictionary_v0.2**

Overall View of what this VBS script does.

1. Searches the registry for a list of RoamingCustom.DIC dictionary files		
2. Creates a directory %UserProfile%\Appdata\Roaming\Microsoft\OfficeSettingMigration
3. Creates a debug log in the new directory so each operation of the script is captured
4. Copies all existing RoamingCustom.DIC files to the new directory as a backup
5. Copies the CUSTOM.DIC to the new directory as a backup
6. Copies all of the words from each of the RoamingCustom.DIC files to the CUSTOM.DIC file
7. Configures the users’ registry so that the user no longer uses the RoamingCustom.DIC, it is removed from the profile
8. Configures the users’ registry so that the CUSTOM.DIC is the only dictionary listed

Additional notes about this script

1. The script can be run multiple times without harm.  It will create new backups of the dictionary files and if it finds any new words, it will add them to the CUSTOM.DIC as an append operation
2. This script can be run at any time.  It does not require being run ‘just in time’ before the Office 2016 upgrade.  
3. Once this script has run, it does not need to be run again.  It leaves the user in a state where they are only using their CUSTOM.DIC, which is a file that will persist the migration
4. If Intel decides to keep some 2013 clients, this script can fix those clients too.  
5. Make sure Intel runs this script while the SignInOptions=3 is still set.  Once that changes and allows for users to determine their identity is federated, it will be too late.  For this reason, we recommend running this sooner, rather than later.
6. This script modifies the HKCU, and as such needs to run in the user context, no the system context.
7. The script runs so fast that it is undetectable unless you go look for the logs and the changes.


###**Manage-OfficeSettings_v0.4**

Overall View of what this VBS script does.

1. Logging - Default folder for logs and backup files is %appdata%\Microsoft\OfficeSettingMigration\
2. Log debug header - key environment variables, user settings and computer settings
3. Log current SignInOptions
4. Backup RoamingCustom.dic
5. Get initial set of identities present on the computer
6. Backup MRUs for each Office app and merge the result to create a single reg file (AD_MRURegBackup.reg)
7. Backup roaming settings and customizations (RoamingADSettingsBackup.reg)
8. Log the list of Office Apps that have MRU
9. Launch Apps and wait for AAD Auth (as signaled by presence of OrgId)
10. Word is launched first.
11. Excel, PowerPoint, Access and Publisher are launched if MRU exists for these apps
12. A reg file (OrgId_MRURegBackup.reg) is created with MRUs and OrgId settings. This file is merged with the registry to restore MRUs



###**Migrate_AD_MRU_and_Roaming_to_OrgId_Final_comments**

Overall description of this powershell script

1. Confirms that only one AD and one OrgId setting exists for all app MRUs and for roaming settings

2. If more than one AD exists, this script does nothing.
   If more than one AAD exists, this script does nothing.
   If no AD exists, this script does nothing.
   If no AAD exists (i.e the user hasn't been migrated to AAD, OR, the user hasn't launched the given Office app before), this script does nothing.

3. Takes a backup of the OrgId settings in all app MRus and roaming, writing it to the path provided in the "-BackupPath" argument

4. Makes the following changes:

    Most Recently Used (MRU)
    ------------------------
    Moves FROM:
	    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\**App**\User MRU\AD_**\
    Moves TO:
	    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\**App**\User MRU\OrgId_**\

    Roaming settings and customizations
    -----------------------------------
    Moves FROM:
	    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\Common\Roaming\Identities\**_AD
    Moves TO:
	    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\Common\Roaming\Identities\**_OrgId

5. If the "-Undo" argument is specified, applies the backup registry files located in the specified "-BackupPath" argument