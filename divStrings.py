#NamesAlgorithm.py 






if __name__ == "__main__":
    print('main')





'''
Public Sub SepararApellidos()

On Error Resume Next


            Dim Ensamblador As dao.Database
            Dim consulta, tabla As dao.Recordset
            
            Dim aux(20) As String
            Dim Num(10) As Variant
            Dim Fecha(5) As Date
            Num(1) = 0
            Num(2) = 0
            Num(6) = 0


    Set Ensamblador = CurrentDb()
    Set tabla = Ensamblador.OpenRecordset("Hoja1")



                        
                        tabla.MoveFirst
                        tabla.MoveLast
                        Num(7) = tabla.RecordCount() 'Número de registros en la tabla
 'OJO A QUÍ
                        'Primero añadimos los campos necesarios a la tabla Hoja
                        DoCmd.OpenQuery "CrearCampoNombreHoja1", , acReadOnly
                        DoCmd.OpenQuery "CrearCampoApellido1Hoja1", , acReadOnly
                        DoCmd.OpenQuery "CrearCampoApellido2Hoja1", , acReadOnly
                        DoCmd.OpenQuery "CampoRestoHoja1", , acReadOnly
                        
                        MsgBox ("Casi seguro que no funciona ejecuta la consulta  crear campos..., las cuatro")
                        'Si no funciona que es posible, ejecuta las cuatro consultas

                        

'MyComp = StrComp(MyStr1, MyStr2, 1)    Returns 0.

                      
                        tabla.MoveFirst
                        While Not tabla.EOF
              

                                            tabla.Edit
                                              Dim Cadena As String
                                              Dim NomApells As String
                                              Dim SeparaNombreYApellidos
                                              'NomApells = tabla!Nombreyapellidos
                                                'Cadena = Split(SeparaNombreYApellidos(NomApells), ",")
                                                
                                                
                                                    'Función del excel
                                                    Dim rng As String
                                                    Dim nombreArr() As String
                                                    Dim nuevaCadena As String
                                                    Dim i As Integer
                                                     
                                                    'Dvidir el nombre por palabras en un arreglo
                                                    rng = tabla!Nombreyapellidos
                                                    nombreArr = Split(Trim(rng))
                                                    
                                                     nuevaCadena = ""
                                                    'Analizar cada palabra dentro del arreglo
                                                    For i = 0 To UBound(nombreArr)
                                                        Select Case LCase(nombreArr(i))
                                                             
                                                            'Palabras que forman parte de un apellido compuesto
                                                            'Agregar nuevas palabras separadas por una coma
                                                            Case "de", "del", "la", "las", "los", "san", "Mª", " "
                                                                'Insertar espacio en blanco
                                                                
                                                                nuevaCadena = nuevaCadena & nombreArr(i) & " "
                                                            Case Else
                                                                'Insertar caracter delimitador
                                                                
                                                                nuevaCadena = nuevaCadena & nombreArr(i) & "@"
                                                         
                                                        End Select
                                                    Next
                                                     
                                                    'Remover el último caracter delimitador de la cadena
                                                    If Right(nuevaCadena, 1) = "@" Then
                                                        nuevaCadena = Left(nuevaCadena, Len(nuevaCadena) - 1)
                                                    End If
                                                     
                                                   ' SepararApellidos = nuevaCadena
                                                     
  
                                            
                                            
                                            
                                            'Función de VBA
                                            Dim NyAp As String
                                            NyAp = Trim(nuevaCadena)
                                                                                        
                                            If NyAp <> "" Then
                                            'Not IsNull(NyAp) Then
                                                    Dim MArray
                                                    MArray = Split(NyAp, "@")
                                                    If UBound(MArray) > 2 Then
                                                       If Len(MArray(1)) = 2 Then
                                                          SeparaNombreYApellidos = MArray(0) & "," & MArray(1) & " " & MArray(2) & "," & MArray(3)
                                                       Else
                                                          SeparaNombreYApellidos = MArray(0) & " " & MArray(1) & "," & MArray(2) & "," & MArray(3)
                                                       End If
                                                    Else
                                                          SeparaNombreYApellidos = MArray(0) & "," & MArray(1) & "," & MArray(2)
                                                    End If
                                                
                                                
                                                'Tercer procedimiento
                                                
                                        If UBound(MArray) >= 3 Then
                                        
                                                'Select Case Trim(MArray(1))
                                                    'Case "LUIS", "JOSEFA", "MARÍA", "JOSEFA", "JOSEFA", "ÁNGELES" & _
                                                  ' ",  'JOSÉ'"

                                                
                                                        tabla!Nombre = MArray(0) & " " & MArray(1)
                                                        tabla!Apellido1 = MArray(2)
                                                        tabla!Apellido2 = MArray(3)
                                                        tabla!Otros = MArray(3)
                                                    
                                                
                                                'End Select
                                                
                                         Else
                                                
                                                
                                                tabla!Nombre = MArray(0)
                                                tabla!Apellido1 = MArray(1)
                                                tabla!Apellido2 = MArray(2)
                                                tabla!Resto = MArray(3)
                                        End If

                                            End If

                                            
                                            tabla.Update
                                            Num(3) = Num(3) + 1
                                          
                                            
                                            
                                       nuevaCadena = ""
                                 

                                    Num(2) = Num(2) + 1
                                    If Num(2) = 1000 Then
                                        Num(2) = 0
                                    End If
                                    
                                   
                                    
                          
                            tabla.MoveNext
                            Num(9) = Num(9) + 1
                            
                        Wend
     tabla.Close
     
End Sub
'''