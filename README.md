# LAD-search ( Local active directory )
VBA/EXCEL Local Active Directory user details search ( Display name, email, NetID )
- This excel VBA scrip opens in an additional excel instance, so you're not blocked using other excel files or macros

Import all 3 files to your excel macro VBA editor ( 4 files total but UserForm1.frm only get imported to VBA )
1) ThisWorkbook.cls - Some code on workbook Open ( file )
2) UserForm1.frm ( file )
3) UserForm1.frx ( file )
4) LAD.bas ( file )

Change the LAD constant in LAD module to your Domain and you're good to go ( Public Const LAD = "LDAP://YourDomainHere.com" )

This was created some years ago, so all improvements welcome ;)
