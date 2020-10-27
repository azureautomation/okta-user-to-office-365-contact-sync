Okta User to Office 365 Contact Sync
====================================

            

This script will synchronise (including updates and deletes) your Okta users to Office 365 contacts if they do not already exist in Office 365.


Matching is done based on the email address of the user.


It only syncs if the user isn't already synced by AADConnect


It only syncs users with both a first and lastname and valid email address


It will update/sync the following fields:


  *  Email 
  *  Firstname 
  *  Lastname 
  *  Address 
  *  Country 
  *  DisplayName 
  *  Zip Code 
  *  City 
  *  Department 
  *  Title 

See my blog at [http://www.lieben.nu](http://www.lieben.nu)


Or check the guide at:[ http://www.lieben.nu/liebensraum/2017/12/setting-up-okta-user-office-365-contact-synchronisation/](http://www.lieben.nu/liebensraum/2017/12/setting-up-okta-user-office-365-contact-synchronisation/)


Or Git:[ https://gitlab.com/Lieben/oktaToOffice365ContactSync/blob/master/OktaContactSync.ps1](https://gitlab.com/Lieben/oktaToOffice365ContactSync/blob/master/OktaContactSync.ps1)

 

        
    
TechNet gallery is retiring! This script was migrated from TechNet script center to GitHub by Microsoft Azure Automation product group. All the Script Center fields like Rating, RatingCount and DownloadCount have been carried over to Github as-is for the migrated scripts only. Note : The Script Center fields will not be applicable for the new repositories created in Github & hence those fields will not show up for new Github repositories.
