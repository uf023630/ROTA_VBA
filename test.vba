Sub Check_Acc_Picture()
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\" & "PIA_Data.accdb" & ";Jet OLEDB:Database Password=uf023630;"

sql1 = "SELECT * FROM PIA_NEW"
rst1.Open sql1, cnn, adOpenKeyset, adLockOptimistic

        A = rst1.RecordCount
        rst1.MoveFirst
        For i = 0 To rst1.RecordCount - 1
                
        G = rst1.Fields("ITNO")
        
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        'If fs.FileExists("\\10.139.4.23\iktw-s004\6. ISL\ACCESS\PIA Photo\" & ITNO & ".JPG") Then
        If fs.FileExists("E:\PIA Photo\" & rst1.Fields(0) & ".JPG") Then
        
            
            
            Sql2 = "SELECT * FROM CFile WHERE ITNO='" & rst1.Fields(0) & "'"
            rst2.Open Sql2, cnn, adOpenKeyset, adLockOptimistic
            
            If rst2.RecordCount = 0 Then Addnew = True
  
            rst2.Close
            Set rst2 = Nothing
        
                If Addnew = True Then
                Sql2 = "SELECT * FROM Cfile"
                Else
                
                Sql2 = "SELECT * FROM Cfile WHERE ITNO='" & rst1.Fields(0) & "'"
                End If
                
                rst2.Open Sql2, cnn, adOpenKeyset, adLockOptimistic

                    If Addnew = True Then rst2.Addnew
                    rst2.Fields("ITNO") = rst1.Fields(0)
                    rst2.Fields("CFile") = True
                    rst2.Update
   
        
        
                rst2.Close
                Set rst2 = Nothing
       
        Else
        G = False
        End If
 
        
        rst1.MoveNext
        
        Next i

rst1.Close
Set rst1 = Nothing


cnn.Close
Set cnn = Nothing
End Sub
