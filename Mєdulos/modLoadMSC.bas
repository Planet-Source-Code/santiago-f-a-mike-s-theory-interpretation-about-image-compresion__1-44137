Attribute VB_Name = "modLoadMSC"
Option Explicit

' Con este procedimiento cargo un archivo del tipo MSC en un picturebox
Public Sub LoadMSC(Path As String, pPic As PictureBox, LoadMaster As Boolean, LoadClones As Boolean, LoadFPixels As Boolean)
    Dim i As Integer, j As Integer, k As Integer, X As Integer, Y As Integer
    ' Usada para acumular los pixels de la imagen
    Dim pPixels() As Byte
    ' Usada para marcar las pos de pixels usados
    Dim UsedPixels() As Boolean
    ' Usadas para la lectura del archivo
    Dim Header As tHeader
    Dim MasterBlock As tMasterBlock
    Dim SlavePreInfo As tPreClone
    Dim SlaveBlock As tCloneBlock
    Dim PixelInfo As tPixelInfo

    ' Abro el archivo
    Open Path For Binary Access Read As #2
        ' Levanto el encabezado del archivo
        Get #2, , Header
        
        ' Redimensiono el picturebox
        pPic.Width = Header.ImgWidth
        pPic.Height = Header.ImgHeight
        
        ' Redimensiono la matriz de pixeles y pixeles usados
        ReDim pPixels(3, Header.ImgWidth - 1, Header.ImgHeight - 1) As Byte
        ReDim UsedPixels(Header.ImgWidth - 1, Header.ImgHeight - 1) As Boolean
        ReDim MasterBlock.Info(Header.BlockWidth - 1, Header.BlockHeight - 1, 2) As Byte
        
        ' Si hay por lo menos un master
        For j = 0 To Header.TotalMasters - 1
            ' Leo la info del master
            Get #2, , MasterBlock
            ' Cargo los pixels del master en la matriz
            For X = MasterBlock.X To MasterBlock.X + Header.BlockWidth - 1
                For Y = MasterBlock.Y To MasterBlock.Y + Header.BlockHeight - 1
                    For k = 0 To 2 ' RGB Colors
                    
                        ' Cargo el pixel
                        If LoadMaster = True Then pPixels(k, X, Y) = MasterBlock.Info(X - MasterBlock.X, Y - MasterBlock.Y, k)
                    
                    Next k
                    ' Marco este pixel usado (para luego cargar los pixels libres
                    UsedPixels(X, Y) = True
                Next Y
            Next X
            ' Leo la cant de clons
            Get #2, , SlavePreInfo
            ' Redimensiono la matriz de los clons (o slaves)
            ReDim SlaveBlock.Info(SlavePreInfo.Cant - 1) As tClonePos
            ' Cargo todas las posiciones de los clones
            For i = 0 To UBound(SlaveBlock.Info())
                Get #2, , SlaveBlock.Info(i)
            Next i
            ' Cargo la misma info que el master (es como que igualo clon=master)
            For i = 0 To UBound(SlaveBlock.Info())
                For X = SlaveBlock.Info(i).X To SlaveBlock.Info(i).X + Header.BlockWidth - 1
                    For Y = SlaveBlock.Info(i).Y To SlaveBlock.Info(i).Y + Header.BlockWidth - 1
                        For k = 0 To 2 ' RGB Colors
                        
                            ' Cargo el pixel
                            If LoadClones = True Then pPixels(k, X, Y) = MasterBlock.Info(X - SlaveBlock.Info(i).X, Y - SlaveBlock.Info(i).Y, k)
                        
                        Next k
                        ' Marco este pixel usado (para luego cargar los pixels libres)
                        UsedPixels(X, Y) = True
                    Next Y
                Next X
            Next i
        Next j
    
        ' Inicializo valores de X e Y
        X = 0: Y = 0
        ' Ahora una vez cargados los Master y Salves cargo los Pixels Libres
        For Y = 0 To Header.ImgHeight - 1
            For X = 0 To Header.ImgWidth - 1
                ' Si el pixel no estaba en uso por algún master o slave...
                If UsedPixels(X, Y) = False Then
                    ' Cargo el valor del pixel
                    Get #2, , PixelInfo
                    ' y lo meto en la matriz de pixeles
                    If LoadFPixels = True Then
                        pPixels(2, X, Y) = PixelInfo.Red
                        pPixels(1, X, Y) = PixelInfo.Green
                        pPixels(0, X, Y) = PixelInfo.Blue
                    End If
                End If
            Next X
        Next Y
    Close #2

    SetDIs pPic, pPixels()
End Sub
