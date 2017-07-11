' WWWeb Concepts wwwebconcepts.com
' James W. Threadgill james@wwwebconcepts.com
' Copyright 2017
' Version 1.0.0.0
'=========================================================

Imports System.Math
Imports System.Convert
Imports WebMatrix.Data
Imports System.Web.HttpContext

Public Class RecordSetPaging

    Public Shared Function RecordSetStats(ByVal offset As Integer, ByVal pageNum As Integer, ByVal pageSize As Integer, ByVal totalRows As Integer) As String

        Dim startRecord As Integer = 0
        If totalRows > 0 Then startRecord = offset + 1
        Dim endRecord As Integer = startRecord + pageSize - 1
        If totalRows < endRecord Then endRecord = totalRows Else endRecord = endRecord
        Dim strRecordSetStats As String = (startRecord & " - " & endRecord & " | " & totalRows)
        If startRecord < 1 Then
            strRecordSetStats = ""
        Else
            strRecordSetStats = strRecordSetStats
        End If

        Return strRecordSetStats
    End Function
    Public Shared Function GetRowCount(ByVal databaseName As String, ByVal column As String, ByVal table As String, ByVal whereClause As String) As Integer

        Dim datasource As Database = Database.Open(databaseName)
        Dim strCountSQL As String = "Select Count(" & column & ") As TotalRows FROM " & table & " WHERE " & whereClause & ""
        Dim rs_itemcount = datasource.QueryValue(strCountSQL)
        Dim itemsTotal As Integer = 0
        If Not IsDBNull(rs_itemcount) Then itemsTotal = ToInt32(rs_itemcount)

        Return itemsTotal
    End Function

    Public Shared Function RecordsetNavBar(ByVal offset As Integer, ByVal pageSize As Integer, ByVal totalRows As Integer, ByVal pageName As String, ByVal QueryParameter As String, ByVal navBarStyle As String, ByVal navBarClass As String) As String

        ' calculate pagination
        Dim totalPages As Integer = ToInt32(Ceiling(totalRows / pageSize))
        Dim pageNum As Integer = (offset / pageSize) + 1
        Dim pageNext As Integer = pageNum + 1
        Dim pagePrev As Integer = pageNum - 1
        Dim pageFirst As Integer = 1
        Dim pagelast As Integer = totalPages
        Dim strPageNumbers As String = ""
        Dim strNavBar As String = ""
        Dim strReturn As String = ""

        If pageSize < totalRows Then

            If (QueryParameter = "/") Then
                QueryParameter = QueryParameter
            ElseIf (QueryParameter <> "") And (InStr(pageName, "?") = 0) Then
                QueryParameter = ("?" & QueryParameter & "=")
            Else
                QueryParameter = ("&amp;" & QueryParameter & "=")
            End If

            ' create RecordSet navigation image & text link strings
            Dim strRSFirst As String = ""
            Dim strRSPrev As String = ""
            Dim strRSNext As String = ""
            Dim strRSLast As String = ""

            Select Case navBarStyle
                Case "blue"
                    strRSFirst = ("<img src=""icons/First record.png"" alt=""" & Current.Application("RSFirstTxt") & """ title=""" & Current.Application("RSFirstTxt") & """ />")
                    strRSPrev = ("<img src=""icons/Go back.png"" alt=""" & Current.Application("RSPrevTxt") & """ title=""" & Current.Application("RSPrevTxt") & """ />")
                    strRSNext = ("<img src=""icons/Go forward.png"" alt=""" & Current.Application("RSNextTxt") & """ title=""" & Current.Application("RSNextTxt") & """ />")
                    strRSLast = ("<img src=""icons/Last recor.png"" alt=""" & Current.Application("RSLastTxt") & """ title=""" & Current.Application("RSLastTxt") & """ />")
                Case "black"
                    strRSFirst = ("<img src=""icons/first_blk.png"" alt=""" & Current.Application("RSFirstTxt") & """ title=""" & Current.Application("RSFirstTxt") & """ />")
                    strRSPrev = ("<img src=""icons/forward_blk.png"" alt=""" & Current.Application("RSPrevTxt") & """ title=""" & Current.Application("RSPrevTxt") & """ />")
                    strRSNext = ("<img src=""icons/back_blk.png"" alt=""" & Current.Application("RSNextTxt") & """ title=""" & Current.Application("RSNextTxt") & """ />")
                    strRSLast = ("<img src=""icons/last_blk.png"" alt=""" & Current.Application("RSLastTxt") & """ title=""" & Current.Application("RSLastTxt") & """ />")
                Case "blue_modern"
                    strRSFirst = ("<img src=""icons/go-first.png"" alt=""" & Current.Application("RSFirstTxt") & """ title=""" & Current.Application("RSFirstTxt") & """ />")
                    strRSPrev = ("<img src=""icons/back.png"" alt=""" & Current.Application("RSPrevTxt") & """ title=""" & Current.Application("RSPrevTxt") & """ />")
                    strRSNext = ("<img src=""icons/forward.png"" alt=""" & Current.Application("RSNextTxt") & """ title=""" & Current.Application("RSNextTxt") & """ />")
                    strRSLast = ("<img src=""icons/go-last.png"" alt=""" & Current.Application("RSLastTxt") & """ title=""" & Current.Application("RSLastTxt") & """ />")
                Case "modern"
                    strRSFirst = ("<img src=""icons/First.gif"" alt=""" & Current.Application("RSFirstTxt") & """ title=""" & Current.Application("RSFirstTxt") & """ />")
                    strRSPrev = ("<img src=""icons/Previous.gif"" alt=""" & Current.Application("RSPrevTxt") & """ title=""" & Current.Application("RSPrevTxt") & """ />")
                    strRSNext = ("<img src=""icons/Next.gif"" alt=""" & Current.Application("RSNextTxt") & """ title=""" & Current.Application("RSNextTxt") & """ />")
                    strRSLast = ("<img src=""icons/Last.gif"" alt=""" & Current.Application("RSLastTxt") & """ title=""" & Current.Application("RSLastTxt") & """ />")
                Case "navy"
                    strRSFirst = ("<img src=""icons/gtk-goto-first-ltr.png"" alt=""" & Current.Application("RSFirstTxt") & """ title=""" & Current.Application("RSFirstTxt") & """ />")
                    strRSPrev = ("<img src=""icons/gtk-go-back-ltr.png"" alt=""" & Current.Application("RSPrevTxt") & """ title=""" & Current.Application("RSPrevTxt") & """ />")
                    strRSNext = ("<img src=""icons/gtk-go-back-rtl.png"" alt=""" & Current.Application("RSNextTxt") & """ title=""" & Current.Application("RSNextTxt") & """ />")
                    strRSLast = ("<img src=""icons/gtk-goto-first-rtl.png"" alt=""" & Current.Application("RSLastTxt") & """ title=""" & Current.Application("RSLastTxt") & """ />")
                Case "orange"
                    strRSFirst = ("<img src=""icons/go-first_orange.png"" alt=""" & Current.Application("RSFirstTxt") & """ title= """ & Current.Application("RSFirstTxt") & """ />")
                    strRSPrev = ("<img src=""icons/go-back_orange.png"" alt=""" & Current.Application("RSPrevTxt") & """ title=""" & Current.Application("RSPrevTxt") & """ />")
                    strRSNext = ("<img src=""icons/go-forward_orange.png"" alt=""" & Current.Application("RSFNextTxt") & """ title=""" & Current.Application("RSNextTxt") & """ />")
                    strRSLast = ("<img src=""icons/go-last_orange.png"" alt=""" & Current.Application("RSLastTxt") & """ title=""" & Current.Application("RSLastTxt") & """ />")
                Case "custom"
                    strRSFirst = ("<img src=""icons/" & Current.Application("RSNavCustomFirstImg") & """ alt=""" & Current.Application("RSFirstTxt") & """ title=""" & Current.Application("RSFirstTxt") & """ />")
                    strRSPrev = ("<img src=""icons/" & Current.Application("RSNavCustomPreviousImg") & """ alt=""" & Current.Application("RSPrevTxt") & """ title=""" & Current.Application("RSPrevTxt") & """ />")
                    strRSNext = ("<img src=""icons/" & Current.Application("RSNavCustomNextImg") & """ alt=""" & Current.Application("RSNextTxt") & """ title=""" & Current.Application("RSNextTxt") & """ />")
                    strRSLast = ("<img src=""icons/" & Current.Application("RSNavCustomLastImg") & """ alt=""" & Current.Application("RSLastTxt") & """ title=""" & Current.Application("RSLastTxt") & """ />")
                Case "pagenumbers" 'page numbers
                    Dim n As Integer
                    Dim pageNumber As Integer = 1
                    strPageNumbers = "<nav class=""" & navBarClass & """>" & vbCrLf & vbTab & "<ul>" & vbCrLf
                    For n = 0 To (totalPages) - 1
                        If pageNum = pageNumber Then
                            strPageNumbers &= vbTab & vbTab & "<li>" & pageNumber.ToString & "</li>" & vbCrLf
                        Else
                            strPageNumbers &= vbTab & vbTab & "<li><a href = """ & pageName & QueryParameter & pageNumber.ToString & """>" & pageNumber.ToString & "</a></li>" & vbCrLf

                        End If
                        pageNumber += 1
                    Next
                    strPageNumbers &= vbTab & "</ul>" & vbCrLf & "</nav>" & vbCrLf
                Case Else 'text strings
                    strRSFirst = ("" & Current.Application("RSFirstTxt") & "")
                    strRSPrev = ("" & Current.Application("RSPrevTxt") & "")
                    strRSNext = ("" & Current.Application("RSNextTxt") & "")
                    strRSLast = ("" & Current.Application("RSLastTxt") & "")
            End Select

            If navBarStyle = "pagenumbers" Then
                Return strPageNumbers
            Else
                ' build navbar links
                Dim strMoveRSFirst As String = vbTab & vbTab & ("<li><a href =""" & pageName & QueryParameter & pageFirst & """>" & strRSFirst & "</a></li>" & vbCrLf)
                Dim strMoveRSPrev As String = vbTab & vbTab & ("<li><a href =""" & pageName & QueryParameter & pagePrev & """>" & strRSPrev & "</a></li>" & vbCrLf)
                Dim strMoveRSNext As String = vbTab & vbTab & ("<li><a href =""" & pageName & QueryParameter & pageNext & """>" & strRSNext & "</a></li>" & vbCrLf)
                Dim strMoveRSLast As String = vbTab & vbTab & ("<li><a href =""" & pageName & QueryParameter & pagelast & """>" & strRSLast & "</a></li>" & vbCrLf)

                strNavBar = "<nav class=""" & navBarClass & """>" & vbCrLf & vbTab & "<ul>" & vbCrLf
                If pageNum > 1 And Not pageNum = totalPages Then
                    strNavBar &= strMoveRSFirst & strMoveRSPrev & strMoveRSNext & strMoveRSLast
                ElseIf pageNum < 2 And Not pageNum = totalPages Then
                    strNavBar &= strMoveRSNext & strMoveRSLast
                ElseIf pageNum = totalPages Then
                    strNavBar &= strMoveRSFirst & strMoveRSPrev
                End If
                strNavBar &= vbTab & "</ul>" & vbCrLf & "</nav>" & vbCrLf
                Return strNavBar
            End If

        End If

            Return strReturn
    End Function

    ' Get Bootstrap Column span from application column setting
    Public Shared Function GetColSpan(ByVal numCols As Integer) As String
        ' valid bootstrap span results 1,2,3,4,6,12
        Dim ColX As Integer
        ' Prevent invalid input
        If numCols > 12 Then
            numCols = 12
        ElseIf numCols < 12 And numCols > 6 Then
            numCols = 6
        End If
        ' find the boostrap column span
        ColX = Round(12 / numCols)
        Return ColX
    End Function

End Class
