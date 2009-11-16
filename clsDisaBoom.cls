VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "clsDisaBoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Const CLASS_NAME As String = "clsDisaBoom"

Public Function ExecFeed(ByVal clsUpload As clsCommonFeedMembers, ByVal ADORS As ADODB.Recordset, ByVal clsGetFieldVals As clsGetDBData, ByVal strFileName As String, ByVal intClientID As Integer) As String
        'Variables
        
        'contact_type_id  description
        '1                Via Email
        '2                Via Postal Mail
        '3                Via Phone
        '4                Via Fax
        '5                Via Web Site
        
3320    'On Error GoTo ExecFeedError


        Const DELIMITER = ","
        Const COUNTRY = "USA"
        Dim objTextFile As TextStream
        Dim objString As Auto_Feed_Members.clsStringBuilder
        Dim objfSys As FileSystemObject
        Dim lngReqEmediaID As Long
        'Required Fields
        Dim strJobTitle As String
        'Dim strHeader As String
        'Dim strFooter As String
        'Dim strDuties As String
        Dim strContactEmail As String
        Dim strCompanyName As String
        Dim strCity As String
        Dim strState As String
        'Dim strCompanyUrl As String
        'Dim strReplyUrl  As String
        Dim blnMediaSupportsImgTracking As Boolean
        Dim blnMediaIsProvisioned As Boolean
        Dim strImageTag As String
        
        Dim strUDLFile As String
        Dim lngReqID As Long
        Dim strListing As String
        'Dim strZipCode As String
        'Dim strCategory As String
        
        Dim strReqCode As String
        Dim strJobAction As String
        Dim intJobAction As Integer
        Dim strClosedDate As String
        Dim strUnpostDate As String
        'Dim strCloseDate As String
        Dim strEndDate As String
        Dim strPostDate As String
        
        Const DATE_FORMAT = "yyyy-mm-dd"
                   
3330    If Not ADORS.EOF Then
3340      Set objfSys = New FileSystemObject

3350      If ((objfSys.FileExists(strFileName)) And (clsUpload.OkToDelete = True)) Then
            'If the file exists, it is not the first record and we are ok'd to delete then
3360        objfSys.DeleteFile (strFileName)
3370      End If

      '40620     strEmediaID = ADORS.Fields("Emedia_ID").Value & ""

3380      Set objTextFile = clsUpload.OpenTextFile(objfSys, strFileName, ForAppending, True)
3390      Set objString = New Auto_Feed_Members.clsStringBuilder
          objString.Supports_Unicode = clsUpload.Supports_Unicode

'3391      strJobHireFrom = clsUpload.GetXMLSetting("CASE_MANAGEMENT", "E_" & clsUpload.EmediaID & "_JOB_FROM", "", "", False)
                    
3400      With ADORS
3410        clsUpload.IsXML = False

3420        Do While Not .EOF
              '*********************************************************************************************************
              'required fields
3421          strUDLFile = .Fields("UDLFile").value
              strJobAction = .Fields("JobAction").value
              
              intJobAction = IIf(((.Fields("JobAction").value = "A") Or (.Fields("JobAction").value = "U")), 0, 1)
              If strJobAction = "A" Then
                    strJobAction = "ADD"
              ElseIf strJobAction = "U" Then
                    strJobAction = "UPDATE"
              Else
                    strJobAction = "DELETE"
              End If
              
3422          lngReqID = .Fields("requisition_id").value
'3430          lngReqEmediaID = ADORS.Fields("req_emediaid").value
3430          lngReqEmediaID = .Fields("req_emediaid").value
              strListing = .Fields("listing").value & ""
3440          strJobTitle = (.Fields("title").value & "")

'3460          strHeader = (.Fields("HEADER").value & "")
'3470          strFooter = (.Fields("FOOTER").value & "")
'3480          strDuties = (.Fields("DUTIES").value & "")
'3490          strRequirements = (.Fields("REQUIREMENTS").value & "")

3530          strContactEmail = (.Fields("email_address").value & "")
'3540          strCompanyName = ADORS.Fields("listing").value
3540          strCompanyName = .Fields("listing").value
3560          strCity = (.Fields("city_name").value & "")
3570          strState = (.Fields("state_name").value & "")
              'strZipCode = (.Fields("zip_code").value & "")
              strReqCode = .Fields("requisition_code").value
              
              strPostDate = Format((.Fields("postdate").value & ""), DATE_FORMAT)
                
              'strCloseDate = Format(.Fields("closedate").value, "DD-MMM-YY")
              'strEndDate = Format(DateAdd("d", CInt(.Fields("duration").value), CDate((.Fields("postdate").value & ""))), date_format)
              
              ' ----------------------------
                strClosedDate = (.Fields("close_datetime").value & "")
                strUnpostDate = (.Fields("unpost_date").value & "")
                
                If ((LenB(strClosedDate) = 0) And (LenB(strUnpostDate) = 0)) Then
                  ' Job has not been closed or unposted
                  strEndDate = Format(DateAdd("d", CInt(.Fields("duration").value), CDate((.Fields("postdate").value & ""))), DATE_FORMAT)
                ElseIf ((LenB(strClosedDate) > 0) And (LenB(strUnpostDate) > 0)) Then
                  ' Seems job has been closed and unposted, get the earliest date of the two.
                  If (CDate(strClosedDate) < CDate(strUnpostDate)) Then
                   strEndDate = Format(strClosedDate, DATE_FORMAT)
                  ElseIf (CDate(strClosedDate) > CDate(strUnpostDate)) Then
                    strEndDate = Format(strUnpostDate, DATE_FORMAT)
                  Else
                    ' If close and unpost dates are the same then choose any.
                    strEndDate = Format(strUnpostDate, DATE_FORMAT)
                  End If
                ElseIf ((LenB(strClosedDate) = 0) And (LenB(strUnpostDate) > 0)) Then
                  ' Job was not closed, it was unposted.
                  strEndDate = Format(strUnpostDate, DATE_FORMAT)
                ElseIf ((LenB(strClosedDate) > 0) And (LenB(strUnpostDate) = 0)) Then
                  ' job was not unposted, it was closed
                  strEndDate = Format(strClosedDate, DATE_FORMAT)
                End If
              ' ----------------------------
              
'3900          strCompanyUrl = (.Fields("url_address").value & "")
'3910          strReplyUrl = clsUpload.GetDirectURL(.Fields("req_emediaID").value, clsUpload.EmediaID, .Fields("apply_statement_indicator").value, strCompanyUrl)
'3920          strReplyUrl = IIf(InStr(1, strReplyUrl, "&amp;", vbTextCompare) = 0, clsUpload.HTMLEncode(strReplyUrl), strReplyUrl)
'3930          strCompanyProfile = ""
                             
              '**************************************************************************************
              'MediaTracking Section.
              'Programmer: Cristian Rey
              'Commented out on 9/15/2005 until production is ready for it
              '**************************************************************************************
              '!!!!!!!!!!!!!!!!!!!!!!!
              'Commenting this out since it seems that there is no need or place for it in the feed file.
              'Ticket #88431
              '!!!!!!!!!!!!!!!!!!!!!!!
'              blnMediaIsProvisioned = IIf(IsNull(ADORS.Fields("Provisioned").value), False, ADORS.Fields("Provisioned").value)
'              strImageTag = ""
'
'              If blnMediaIsProvisioned = True Then
'                blnMediaSupportsImgTracking = IIf(IsNull(ADORS.Fields("MediaTracking").value), False, ADORS.Fields("MediaTracking").value)
'
'                Call clsUpload.GetMediaTrackingInfo(lngReqEmediaID, strContactEmail, strReplyUrl, blnMediaSupportsImgTracking)
'
'                strContactEmail = clsUpload.MediaContactEmail  ' colMediaTrack.Item(ITEM_EMAIL)
'                strReplyUrl = clsUpload.MediaContactURL  ' colMediaTrack.Item(ITEM_URL)
'                strImageTag = clsUpload.MediaImage  ' colMediaTrack.Item(ITEM_IMG)
'              End If
            '**************************************************************************************
      '        lngReqEmediaID = .Fields("req_emediaid").Value
              'strDescription = strDescription & strImageTag
              'strReplyUrl = IIf(strReplyUrl = "http://", "", strReplyUrl)
              'strFooter = strFooter & clsUpload.HTMLEncodeText("<BR>To Apply for this position, please <a href='" & strReplyUrl & "'><b>CLICK HERE</b></a></P> ")
             
              
3940        With objString
3950           .CLEAR
               
3960          If ((clsUpload.OkToDelete = True) And (ADORS.AbsolutePosition = 1)) Then
3970            .Append clsUpload.FormatFeedString("Action", True, DELIMITER, 0, True, False, False)
3980            .Append clsUpload.FormatFeedString("Job number", True, DELIMITER, 0, True, False, False)
3990            .Append clsUpload.FormatFeedString("Job title", True, DELIMITER, 0, True, False, False)
4000            .Append clsUpload.FormatFeedString("Job link", True, DELIMITER, 0, True, False, False)
4010            .Append clsUpload.FormatFeedString("Postdate", True, DELIMITER, 0, True, False, False)
4020            .Append clsUpload.FormatFeedString("Close date", True, DELIMITER, 0, True, False, False)
4030            .Append clsUpload.FormatFeedString("Requisition code", True, DELIMITER, 0, True, False, False)
4040            .Append clsUpload.FormatFeedString("City", True, DELIMITER, 0, True, False, False)
4050            .Append clsUpload.FormatFeedString("State", True, DELIMITER, 0, True, False, False)
4060            .Append clsUpload.FormatFeedString("Company name", True, DELIMITER, 0, True, False, False) & vbCrLf
              End If
              
4320          .Append clsUpload.FormatFeedString(strJobAction, True, DELIMITER, 0, True, False, False)
4330          .Append clsUpload.FormatFeedString(lngReqEmediaID, True, DELIMITER, 0, True, False, False)
4340          .Append clsUpload.FormatFeedString(strJobTitle, True, DELIMITER, 0, True, False, False)
4350          .Append clsUpload.FormatFeedString(clsUpload.GetAlternateApplyLink(lngReqEmediaID), True, DELIMITER, 0, True, False, False)
4351          .Append clsUpload.FormatFeedString(strPostDate, True, DELIMITER, 0, True, False, False)
4360          .Append clsUpload.FormatFeedString(strEndDate, True, DELIMITER, 0, True, False, False)
4370          .Append clsUpload.FormatFeedString(strReqCode, True, DELIMITER, 0, True, False, False)
4380          .Append clsUpload.FormatFeedString(strCity, True, DELIMITER, 0, True, False, False)
4390          .Append clsUpload.FormatFeedString(strState, True, DELIMITER, 0, True, False, False)
4400          .Append clsUpload.FormatFeedString(strListing, True, DELIMITER, 0, True, False, False)
            End With
                
4670          objTextFile.WriteLine objString.value
          
4680          clsGetFieldVals.UpdateHistory lngReqEmediaID, clsUpload.FeedID, strUDLFile, intJobAction

4690          .MoveNext
4700        Loop
4710      End With
          
'4720      objTextFile.Close
'         '******************************************************************
'
'4730      Set objTextFile = Nothing
'4740      Set objfSys = Nothing
4750    End If
        
      '        Date: 01/15/2005
      '        Originally the params were username and pwd, this was not accurate enough so it was changed to username and FID.
      '        To keep compatibility between com components the proc was changed to take an optional username and a feedid param.
      '        The proc ignores the username param completely, here only for compatibility reasons.
      '1440    clsGetFieldVals.UpdateDateSent "", CStr(clsUpload.FeedID)  ' strUserName, strPassword

        'Everything went ok creating the file
4760   ExecFeed = clsUpload.StandardReturn
        
        objTextFile.Close
          
        Set objTextFile = Nothing
        Set objfSys = Nothing

        Set objString = Nothing
        If ADORS.State = 1 Then ADORS.Close
        Set ADORS = Nothing
        Set clsGetFieldVals = Nothing
        Set clsUpload = Nothing
        
4770    Exit Function
        
      'If there is an error creating a file, we don't want the entire process for this particular user name or file
      'Trap the error and return it as a string.
ExecFeedError:
4771  clsUpload.Error_FeedClass = CLASS_NAME
4772  clsUpload.Error_FeedDescription = Err.Description
4773  clsUpload.Error_FeedNumber = Err.Number
4774  clsUpload.Error_FeedInformation = "An Error has occurred"
4775  clsUpload.Error_FeedLineNumber = CStr(Erl)
4776  clsUpload.Error_FeedReqID = lngReqID
4777  clsUpload.Error_FeedReqEmID = lngReqEmediaID
4778  clsUpload.Error_FeedUDLFile = strUDLFile
  
4779  ExecFeed = "Error"

'4780  ExecFeed = "An Error has occurred in:ExecFeed which is member of class " & CLASS_NAME & vbCrLf & "Error number:" & Err.Number & vbCrLf & "Error description:" & Err.Description & "|Requisition_id=" & ADORS.Fields("requisition_id").value & "|Req_EmediaID=" & lngReqEmediaID & "|ClientID=" & intClientID & "|ClientName=" & clsUpload.ClientName & "|" & vbCrLf & "On Code Line:" & Erl

        objTextFile.Close
          
        Set objTextFile = Nothing
        Set objfSys = Nothing

        Set objString = Nothing
        If ADORS.State = 1 Then ADORS.Close
        Set ADORS = Nothing
        Set clsGetFieldVals = Nothing
        Set clsUpload = Nothing

End Function





