VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub EmailTestTemplate()
 Dim MyItem As Outlook.MailItem
 
 Set MyItem = Application.CreateItemFromTemplate("C:\Users\USER\Documents\Outlook\File.oft")
 MyItem.Display
End Sub
