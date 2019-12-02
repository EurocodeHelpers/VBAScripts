Public Sub PrintBentData()

'README!
'This macro produces PDF reports for each type of bent.
'By nature the macro is quite hacky - please check through the reports produced and do some manual checks to be sure you're happy with the output.
'Before running the macro please check youre happy with the below settings.

'Set up the global variables
Dim bentsRange As Range
Dim outputDirectory As String
Dim printingRange As Range

'------------------ SETTINGS----------------------------
Set bentsRange = Range("P1:DC1")                        'Sets out the list of bents for which data is extracted.
outputDirectory = "C:\Users\cox87208\Desktop\"      'Specify the folder directory to which the reports are outputted e.g. "C:\Users\twood\Desktop\"
Set printingRange = Range("D1:N465")
'------------------END SETTINGS----------------------------


'Get an array of bents based off the range of bents defined above
Dim arrayOfBentNumbers As Variant
arrayOfBentNumbers = Application.Transpose(Application.Transpose(bentsRange))

For Each bent In arrayOfBentNumbers
    
    'Filter by bent
    ActiveSheet.ListObjects("Table24").Range.AutoFilter Field:=8, Criteria1:=bent
           
    'Get the filename - we must separate by hyphen not backslash otherwise we will get an error when saving (i.e. 55/01 wont work, 55-01 will be ok.
    Dim splitString() As String
    splitString = Split(bent, "/")
    Dim fileName As String
    fileName = outputDirectory + splitString(0) + "-" + splitString(1) + ".pdf"
    
    'Finally print to pdf...
    printingRange.Select
    ChDir outputDirectory
    Selection.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
        fileName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
Next

End Sub


Peter Cox
Assistant Engineer
peter.cox@mottmac.com
 
Mott MacDonald
Mott MacDonald House
8-10 Sydenham Road
Croydon CR0 2EE 
United Kingdom 


Website   |   Twitter   |   LinkedIn   |   Facebook   |   Instagram   |   YouTube 

Mott MacDonald International Limited. Registered in England and Wales no. 2064414. Registered office: Mott MacDonald House, 8-10 Sydenham Road, Croydon CR0 2EE, United Kingdom 

The information contained in this e-mail is intended only for the person or entity to which it is addressed and may contain confidential and/or privileged material. If you are not the intended recipient of this e-mail, the use of this information or any disclosure, copying or distribution is prohibited and may be unlawful. If you received this in error, please contact the sender and delete the material from any computer. 


